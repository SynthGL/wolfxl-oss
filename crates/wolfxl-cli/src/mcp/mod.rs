//! File-based MCP over newline-delimited JSON-RPC on standard input/output.

mod files;
mod schema;
mod tools;

use std::collections::HashMap;
use std::io::{self, BufRead, Write};
use std::path::PathBuf;

use anyhow::{bail, Context, Result};
use serde_json::{json, Value};

use crate::McpArgs;

const PROTOCOL: &str = "2025-06-18";
/// Undo snapshots kept per workbook; the oldest are dropped first.
const HISTORY_ENTRIES: usize = 32;
/// Snapshot bytes kept per workbook; the newest snapshot is kept even when larger.
const HISTORY_BYTES: usize = 256 << 20;

#[derive(Default)]
struct History {
    undo: Vec<Vec<u8>>,
    redo: Vec<Vec<u8>>,
    revision: String,
}

impl History {
    /// Drop the oldest undo snapshots beyond the entry and byte budgets.
    fn trim(&mut self) {
        let mut bytes: usize = self.undo.iter().map(Vec::len).sum();
        let mut dropped = 0;
        while self.undo.len() - dropped > 1
            && (self.undo.len() - dropped > HISTORY_ENTRIES || bytes > HISTORY_BYTES)
        {
            bytes -= self.undo[dropped].len();
            dropped += 1;
        }
        self.undo.drain(..dropped);
    }
}

struct Server {
    root: Option<PathBuf>,
    tools: Vec<Value>,
    initialized: bool,
    history: HashMap<PathBuf, History>,
}

pub(crate) fn run(args: McpArgs) -> Result<()> {
    let root = args
        .root
        .map(|path| path.canonicalize().context("resolving --root"))
        .transpose()?;
    if root.as_ref().is_some_and(|path| !path.is_dir()) {
        bail!("--root must name a directory");
    }
    let mut server = Server {
        root,
        tools: schema::tools(args.read_only),
        initialized: false,
        history: HashMap::new(),
    };
    let stdin = io::stdin();
    let mut stdout = io::stdout().lock();
    for line in stdin.lock().lines() {
        let line = line.context("reading MCP request")?;
        let response = match serde_json::from_str::<Value>(&line) {
            Ok(request) => server.handle(request),
            Err(error) => Some(rpc_error(
                Value::Null,
                -32700,
                &format!("Parse error: {error}"),
            )),
        };
        if let Some(response) = response {
            serde_json::to_writer(&mut stdout, &response).context("encoding MCP response")?;
            stdout.write_all(b"\n").context("writing MCP response")?;
            stdout.flush().context("flushing MCP response")?;
        }
    }
    Ok(())
}

fn rpc_error(id: Value, code: i32, message: &str) -> Value {
    json!({"jsonrpc":"2.0", "id":id, "error":{"code":code, "message":message}})
}

impl Server {
    fn handle(&mut self, request: Value) -> Option<Value> {
        let id = request.get("id").cloned().unwrap_or(Value::Null);
        if !request.is_object()
            || request["jsonrpc"] != "2.0"
            || !request["method"].is_string()
            || !(id.is_null() || id.is_string() || id.is_number())
        {
            return Some(rpc_error(
                Value::Null,
                -32600,
                "Invalid Request: expected a JSON-RPC 2.0 request",
            ));
        }
        // Notifications never receive a response, including unknown methods.
        request.get("id")?;
        if request
            .get("params")
            .is_some_and(|params| !params.is_object())
        {
            return Some(rpc_error(id, -32602, "Invalid params: expected an object"));
        }
        let empty = json!({});
        let params = request.get("params").unwrap_or(&empty);
        let result = match request["method"].as_str().unwrap_or_default() {
            "initialize" => {
                let Some(version) = params["protocolVersion"].as_str() else {
                    return Some(rpc_error(
                        id,
                        -32602,
                        "Invalid params: protocolVersion must be a string",
                    ));
                };
                let version = match version {
                    "2025-06-18" | "2025-03-26" | "2024-11-05" => version,
                    _ => PROTOCOL,
                };
                self.initialized = true;
                json!({"protocolVersion":version,"capabilities":{"tools":{"listChanged":false}},"serverInfo":{"name":"wolfxl","version":env!("CARGO_PKG_VERSION")},"instructions":"Read get_workbook before editing. File revisions are SHA-256 of exact file bytes. Use expected_revision to guard edits and dry_run to preview them. History lasts only for this server process."})
            }
            "ping" => json!({}),
            "tools/list" | "tools/call" if !self.initialized => {
                return Some(rpc_error(
                    id,
                    -32002,
                    "Server not initialized; call initialize first",
                ));
            }
            "tools/list" => {
                if params.get("cursor").is_some() {
                    return Some(rpc_error(
                        id,
                        -32602,
                        "Invalid params: this server does not paginate tools",
                    ));
                }
                json!({"tools":self.tools})
            }
            "tools/call" => {
                let Some(name) = params["name"].as_str() else {
                    return Some(rpc_error(
                        id,
                        -32602,
                        "Invalid params: name must be a string",
                    ));
                };
                let Some(tool) = self.tools.iter().find(|tool| tool["name"] == name) else {
                    return Some(rpc_error(
                        id,
                        -32602,
                        &format!("Unknown or unavailable tool: {name}"),
                    ));
                };
                let arguments = params.get("arguments").unwrap_or(&empty);
                if let Err(error) = schema::validate(&tool["inputSchema"], arguments, "arguments") {
                    return Some(rpc_error(id, -32602, &format!("Invalid params: {error}")));
                }
                let (document, is_error) = match self.call(name, arguments) {
                    Ok(document) => (document, false),
                    Err(error) => (json!({"error":format!("{error:#}")}), true),
                };
                json!({"content":[{"type":"text","text":document.to_string()}],"structuredContent":document,"isError":is_error})
            }
            _ => return Some(rpc_error(id, -32601, "Method not found")),
        };
        Some(json!({"jsonrpc":"2.0","id":id,"result":result}))
    }
}
