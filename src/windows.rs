use crate::{declare::PrintOptions, fsys::remove_file};
use std::env;
use std::fs::File;
use std::io::Write;
use std::os::windows::process::CommandExt;
use std::path::Path;
use std::process::Command;

const CREATE_NO_WINDOW: u32 = 0x08000000;

fn create_file(dir: &Path, bin: &[u8]) -> std::io::Result<()> {
    let mut f = File::create(dir.join("sm.exe"))?;
    f.write_all(bin)?;
    f.sync_all()?;
    Ok(())
}

fn run_powershell(script_body: &str) -> std::io::Result<std::process::Output> {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);

    let script = format!(
        "$OutputEncoding = [System.Text.Encoding]::UTF8; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; {}",
        script_body
    );

    cmd.args(["-NoProfile", "-NonInteractive", "-Command", script.as_str()]);
    cmd.output()
}

fn decode_output(bytes: &[u8]) -> String {
    String::from_utf8_lossy(bytes).trim().to_string()
}

fn escape_ps_single_quotes(value: &str) -> String {
    value.replace('\'', "''")
}

fn run_powershell_json(script_body: &str, fallback_json: &str) -> String {
    match run_powershell(script_body) {
        Ok(output) => {
            if !output.status.success() {
                return fallback_json.to_string();
            }

            let stdout = decode_output(&output.stdout);
            if stdout.is_empty() {
                fallback_json.to_string()
            } else {
                stdout
            }
        }
        Err(_) => fallback_json.to_string(),
    }
}

pub fn init_windows() {
    let sm = include_bytes!("bin/sm");
    let dir: std::path::PathBuf = env::temp_dir();
    if let Err(error) = create_file(&dir, sm) {
        eprintln!("Failed to initialize printer helper executable: {}", error);
    }
}

pub fn get_printers() -> String {
    run_powershell_json(
        "Get-Printer | Select-Object Name, DriverName, JobCount, PrintProcessor, PortName, ShareName, ComputerName, PrinterStatus, Shared, Type, Priority | ConvertTo-Json -Compress",
        "[]",
    )
}

pub fn get_printers_by_name(printername: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    run_powershell_json(
        format!(
            "Get-Printer -Name '{}' | Select-Object Name, DriverName, JobCount, PrintProcessor, PortName, ShareName, ComputerName, PrinterStatus, Shared, Type, Priority | ConvertTo-Json -Compress",
            escaped_name
        )
        .as_str(),
        "null",
    )
}

pub fn print_pdf(options: PrintOptions) -> String {
    println!("options id {}", options.id);
    println!("options print_setting {}", options.print_setting);

    let dir: std::path::PathBuf = env::temp_dir();

    let remove_after_print = options.remove_after_print;
    let path_to_remove = options.path.clone();

    let exe_path = format!("{}sm.exe", dir.display());
    let mut cmd = Command::new(&exe_path);

    if options.id.is_empty() {
        cmd.arg("-print-to-default");
    } else {
        let printer_name = options.id.trim_matches('"');
        cmd.arg("-print-to");
        cmd.arg(printer_name);
    }

    if !options.print_setting.is_empty() {
        for s in options.print_setting.split_whitespace() {
            cmd.arg(s);
        }
    }

    cmd.arg("-silent");
    cmd.arg(&options.path);
    cmd.creation_flags(CREATE_NO_WINDOW);

    println!("Executing command: {:?}", cmd);

    let result = match cmd.output() {
        Ok(output) => {
            if output.status.success() {
                let stdout = decode_output(&output.stdout);
                if stdout.is_empty() {
                    "OK".to_string()
                } else {
                    stdout
                }
            } else {
                let stdout = decode_output(&output.stdout);
                let stderr = decode_output(&output.stderr);
                format!(
                    "Command failed with status: {}\\nStdout:\\n{}\\nStderr:\\n{}",
                    output.status, stdout, stderr
                )
            }
        }
        Err(error) => format!("Failed to execute command: {}", error),
    };

    if remove_after_print {
        if let Err(error) = remove_file(&path_to_remove) {
            eprintln!("Failed to remove temp print file '{}': {}", path_to_remove, error);
        }
    }

    result
}


pub fn get_jobs(printername: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    run_powershell_json(
        format!(
            "Get-PrintJob -PrinterName '{}' | Select-Object DocumentName,Id,TotalPages,Position,Size,SubmmitedTime,UserName,PagesPrinted,JobTime,ComputerName,Datatype,PrinterName,Priority,SubmittedTime,JobStatus | ConvertTo-Json -Compress",
            escaped_name
        )
        .as_str(),
        "[]",
    )
}

pub fn get_jobs_by_id(printername: String, jobid: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    let escaped_jobid = escape_ps_single_quotes(&jobid);
    run_powershell_json(
        format!(
            "Get-PrintJob -PrinterName '{}' -ID '{}' | Select-Object DocumentName,Id,TotalPages,Position,Size,SubmmitedTime,UserName,PagesPrinted,JobTime,ComputerName,Datatype,PrinterName,Priority,SubmittedTime,JobStatus | ConvertTo-Json -Compress",
            escaped_name, escaped_jobid
        )
        .as_str(),
        "null",
    )
}


pub fn resume_job(printername: String, jobid: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    let escaped_jobid = escape_ps_single_quotes(&jobid);
    run_powershell_json(
        format!(
            "Resume-PrintJob -PrinterName '{}' -ID '{}' | Out-String",
            escaped_name, escaped_jobid
        )
        .as_str(),
        "",
    )
}

pub fn restart_job(printername: String, jobid: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    let escaped_jobid = escape_ps_single_quotes(&jobid);
    run_powershell_json(
        format!(
            "Restart-PrintJob -PrinterName '{}' -ID '{}' | Out-String",
            escaped_name, escaped_jobid
        )
        .as_str(),
        "",
    )
}

pub fn pause_job(printername: String, jobid: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    let escaped_jobid = escape_ps_single_quotes(&jobid);
    run_powershell_json(
        format!(
            "Suspend-PrintJob -PrinterName '{}' -ID '{}' | Out-String",
            escaped_name, escaped_jobid
        )
        .as_str(),
        "",
    )
}

pub fn remove_job(printername: String, jobid: String) -> String {
    let escaped_name = escape_ps_single_quotes(&printername);
    let escaped_jobid = escape_ps_single_quotes(&jobid);
    run_powershell_json(
        format!(
            "Remove-PrintJob -PrinterName '{}' -ID '{}' | Out-String",
            escaped_name, escaped_jobid
        )
        .as_str(),
        "",
    )
}
