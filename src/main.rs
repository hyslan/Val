#![windows_subsystem = "windows"]
use std::process::Command;

fn main() {
    // ordem dos argumentos: session, option, contrato, familia (opcional) 
    Command::new("python")
        .args([&"-m", "src.main", "-s", "0",
            "-o", "9",
            "-c", "4600041302",
        "-f", "hidrometro"])
        .spawn()
        .expect("failed to execute process");

}
