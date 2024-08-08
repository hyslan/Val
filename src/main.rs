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
const _FAMILY: [&str; 5] = ["cavalete", "religacao", "supressao", "hidrometro", "poco"];

fn main() {
    // Lista de contratos
    let contracts = [_ZCMLN, _NOVASPMLG, _ZIGURATEMLQ];
    let mut session: u32 = 0;
    let mut str_session: String = session.to_string();
    Command::new("python")
        .args([&"-m", "python.src.charging_sessions"])
        .spawn()
        .expect("failed to execute process");

    print!("Iniciando as sessões das conexões...");
    thread::sleep(Duration::from_secs(60));
    // ordem dos argumentos: session, option, contrato, familia (opcional), senha, revalorar
    for &contract in &contracts {
        let args = vec![
            "-m",
            "main",
            "-s",
            str_session.as_str(),
            "-o",
            "9",
            "-c",
            contract, // Catch all family possible.
            "-p",
            "alefafa",
        ];

        // Adiciona os elementos da família ao vetor de argumentos
        //         for &family in FAMILY.iter() {
        //             args.push(family);
        //         }

        Command::new("python")
            .args(&args)
            .spawn()
            .expect("failed to execute process");

        session += 1;
        str_session = session.to_string();

        print!(
            "Iniciando a sessão {} para o contrato {}...",
            str_session, contract
        );
        print!("Aguardando 40 segundos para iniciar a próxima sessão...");
        thread::sleep(Duration::from_secs(40));
    }
}
