//! Consumers use discovery metadata to distinguish volatile formulas from
//! formulas whose values can be reused between recalculations.

use assert_cmd::Command;
use serde_json::Value;

#[test]
fn discovery_identifies_only_volatile_functions() {
    let output = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(["list-functions", "--json"])
        .output()
        .expect("run function discovery");
    assert!(output.status.success());
    let document: Value = serde_json::from_slice(&output.stdout).expect("discovery JSON");
    let functions = document["functions"].as_array().expect("functions array");
    let volatile: Vec<_> = functions
        .iter()
        .filter(|function| {
            function["volatile"]
                .as_bool()
                .expect("each function advertises volatility")
        })
        .map(|function| function["name"].as_str().expect("function name"))
        .collect();
    assert_eq!(volatile, ["NOW", "OFFSET", "RAND", "RANDBETWEEN", "TODAY"]);
}
