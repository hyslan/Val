#![windows_subsystem = "windows"]

use std::process::Command;
use std::thread;
use std::time::Duration;

// Contratos
const _NOVASPMLG: &str = "4600041302";
const _GBITAQUERA: &str = "4600042888";
const _NORTESULMLG: &str = "4600043760";
const _NORTESULMLN: &str = "4600045267";
const _NORTESULMLN2: &str = "4600046036";
const _NORTESULMLQ: &str = "4600043654";
const _ZCMLN: &str = "4600042975";
const _RECAPEMLN: &str = "4600044787";
const _RECAPEMLQ: &str = "4600044777";
const _RECAPEMLG: &str = "4600044782";
const _ZIGURATEMLQ: &str = "4600056089";

// Familias
const FAMILY: [&str; 5]  = ["cavalete", "religacao", "supressao", "hidrometro", "poco"];

fn main() {
    Command::new("python")
        .args([&"-m", "src.sap_connection"])
        .spawn()
        .expect("failed to execute process");

    thread::sleep(Duration::from_secs(60));
    //ordem dos argumentos: session, option, contrato, familia (opcional), senha, revalorar
    Command::new("python")
        .args([&"-m", "src.main", "-s", "0",
            "-o", "9",
            "-c", _NOVASPMLG,
        "-f", FAMILY[0],
        "-p", "alefafa"
        ])
        .spawn()
        .expect("failed to execute process");
    Command::new("python")
        .args([&"-m", "src.main", "-s", "1",
            "-o", "9",
            "-c", _NOVASPMLG,
        "-f", FAMILY[1],
        "-p", "alefafa"
        ])
        .spawn()
        .expect("failed to execute process");
    Command::new("python")
        .args([&"-m", "src.main", "-s", "2",
            "-o", "9",
            "-c", _NOVASPMLG,
        "-f", FAMILY[2],
        "-p", "alefafa"
        ])
        .spawn()
        .expect("failed to execute process");

    Command::new("python")
        .args([&"-m", "src.main", "-s", "3",
            "-o", "9",
            "-c", _ZIGURATEMLQ,
        "-f", FAMILY[0],
        "-p", "alefafa"
        ])
        .spawn()
        .expect("failed to execute process");
    Command::new("python")
        .args([&"-m", "src.main", "-s", "4",
            "-o", "9",
            "-c", _ZIGURATEMLQ,
        "-f", FAMILY[1],
        "-p", "alefafa"
        ])
        .spawn()
        .expect("failed to execute process");
    Command::new("python")
        .args([&"-m", "src.main", "-s", "5",
            "-o", "9",
            "-c", _ZIGURATEMLQ,
        "-f", FAMILY[2],
        "-p", "alefafa"
        ])
        .spawn()
        .expect("failed to execute process");
}
