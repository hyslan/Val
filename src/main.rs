#![windows_subsystem = "windows"]
use std::process::Command;

fn main() {
    Command::new("python")
        .args([&"-m", "src.main"])
        .spawn()
        .expect("failed to execute process");

}
