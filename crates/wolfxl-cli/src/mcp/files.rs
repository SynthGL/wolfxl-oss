use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use anyhow::{anyhow, bail, Context, Result};
use serde_json::{json, Value};
use sha2::{Digest, Sha256};

use super::{History, Server};

static TEMP_ID: AtomicU64 = AtomicU64::new(0);

pub(super) fn revision(bytes: &[u8]) -> String {
    base16ct::lower::encode_string(&Sha256::digest(bytes))
}

/// Write already-rendered package bytes so dry-run and committed revisions agree.
fn atomic_write(path: &Path, bytes: &[u8]) -> Result<()> {
    let parent = path
        .parent()
        .ok_or_else(|| anyhow!("workbook has no parent directory"))?;
    let temp = parent.join(format!(
        ".wolfxl-mcp-{}-{}.tmp",
        std::process::id(),
        TEMP_ID.fetch_add(1, Ordering::Relaxed)
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&temp)
        .context("creating temporary workbook")?;
    let outcome = (|| -> Result<()> {
        file.set_permissions(fs::metadata(path)?.permissions())?;
        file.write_all(bytes)
            .context("writing temporary workbook")?;
        file.sync_all().context("syncing temporary workbook")?;
        fs::rename(&temp, path).context("replacing workbook")?;
        Ok(())
    })();
    if outcome.is_err() {
        let _ = fs::remove_file(&temp);
    }
    outcome
}

impl Server {
    pub(super) fn path(&self, argument: &Value, name: &str) -> Result<PathBuf> {
        let raw = argument[name]
            .as_str()
            .ok_or_else(|| anyhow!("{name} must be a path string"))?;
        let input = Path::new(raw);
        let input = match &self.root {
            Some(root) if input.is_relative() => root.join(input),
            _ => input.to_path_buf(),
        };
        let path = input
            .canonicalize()
            .with_context(|| format!("resolving {name} {raw:?}"))?;
        if self
            .root
            .as_ref()
            .is_some_and(|root| !path.starts_with(root))
        {
            bail!("{name} is outside --root: {raw:?}");
        }
        if !path.is_file() {
            bail!("{name} must name an existing regular file: {raw:?}");
        }
        Ok(path)
    }

    pub(super) fn check_revision(arguments: &Value, current: &str) -> Result<()> {
        if let Some(expected) = arguments["expected_revision"].as_str() {
            if expected != current {
                bail!("revision conflict: expected {expected}, current {current}; read the workbook again before editing");
            }
        }
        Ok(())
    }

    pub(super) fn commit_edit(
        &mut self,
        path: &Path,
        before: Vec<u8>,
        after: Vec<u8>,
        arguments: &Value,
        change: Value,
    ) -> Result<Value> {
        let before_revision = revision(&before);
        let after_revision = revision(&after);
        let dry_run = arguments["dry_run"].as_bool().unwrap_or(false);
        Self::check_revision(arguments, &before_revision)?;
        if !dry_run {
            if revision(&fs::read(path).context("checking workbook before write")?)
                != before_revision
            {
                bail!("revision conflict: workbook changed while preparing the edit");
            }
            atomic_write(path, &after)?;
            let history = self.history.entry(path.to_path_buf()).or_default();
            if history.revision != before_revision {
                *history = History::default();
            }
            history.undo.push(before);
            history.redo.clear();
            history.revision.clone_from(&after_revision);
        }
        Ok(
            json!({"path":path,"revision":after_revision,"previous_revision":before_revision,"dry_run":dry_run,"change":change}),
        )
    }

    pub(super) fn restore(&mut self, path: &Path, arguments: &Value, redo: bool) -> Result<Value> {
        let before = fs::read(path).context("reading workbook history target")?;
        let before_revision = revision(&before);
        Self::check_revision(arguments, &before_revision)?;
        let history = self
            .history
            .get_mut(path)
            .ok_or_else(|| anyhow!("no history for this path in this server process"))?;
        if history.revision != before_revision {
            bail!("revision conflict: workbook changed outside this server; undo/redo refused");
        }
        let source = if redo { &history.redo } else { &history.undo };
        let after = source
            .last()
            .ok_or_else(|| anyhow!("nothing to {}", if redo { "redo" } else { "undo" }))?;
        let after_revision = revision(after);
        let dry_run = arguments["dry_run"].as_bool().unwrap_or(false);
        if !dry_run {
            if revision(&fs::read(path)?) != before_revision {
                bail!("revision conflict: workbook changed while restoring history");
            }
            atomic_write(path, after)?;
            if redo {
                history.redo.pop();
                history.undo.push(before);
            } else {
                history.undo.pop();
                history.redo.push(before);
            }
            history.revision.clone_from(&after_revision);
        }
        Ok(
            json!({"path":path,"revision":after_revision,"previous_revision":before_revision,"dry_run":dry_run,"change":{"operation":if redo {"redo"} else {"undo"}}}),
        )
    }
}
