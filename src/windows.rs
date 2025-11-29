
use std::process::Command;
use std::{sync::mpsc, io::Write};
use std::thread;
use std::fs::File;
use std::env;
use crate::{declare::PrintOptions, fsys::remove_file};
use std::os::windows::process::CommandExt;

const CREATE_NO_WINDOW: u32 = 0x08000000;

/**
 * Create sm.exe to temp
 */
fn create_file(path: String, bin: &[u8]) -> std::io::Result<()> {
    let mut f = File::create(format!("{}sm.exe", path))?;
    f.write_all(bin)?;
  
    f.sync_all()?;
    Ok(())
}

  
/**
 * init sm.exe
 */
pub fn init_windows() {
    let sm = include_bytes!("bin/sm");
    let dir: std::path::PathBuf = env::temp_dir();
    let result: Result<(), std::io::Error>  = create_file(dir.display().to_string(),sm);
    if result.is_err() {
        panic!("Gagal")
    }
}

/**
 * Get printers on windows using powershell
 */
pub fn get_printers() -> String {
    // Create a channel for communication
    let (sender, receiver) = mpsc::channel();

    // Spawn a new thread
    thread::spawn(move || {
        let mut cmd = Command::new("powershell");
        cmd.creation_flags(CREATE_NO_WINDOW);
        let output = cmd.args(["Get-Printer | Select-Object Name, DriverName, JobCount, PrintProcessor, PortName, ShareName, ComputerName, PrinterStatus, Shared, Type, Priority | ConvertTo-Json"]).output().unwrap();

        sender.send(String::from_utf8(output.stdout).unwrap()).unwrap();
    });

    // Do other non-blocking work on the main thread

    // Receive the result from the spawned thread
    let result: String = receiver.recv().unwrap();


    return result;
}

/**
 * Get printers by name on windows using powershell
 */
pub fn get_printers_by_name(printername: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Get-Printer -Name \"{}\" | Select-Object Name, DriverName, JobCount, PrintProcessor, PortName, ShareName, ComputerName, PrinterStatus, Shared, Type, Priority | ConvertTo-Json", printername)]).output().unwrap();

    return String::from_utf8(output.stdout).unwrap();
}

/**
 * Print pdf file 
 */
pub fn print_pdf (options: PrintOptions) -> String {
    println!("options id {}", options.id);
    println!("options print_setting {}", options.print_setting);

    let dir: std::path::PathBuf = env::temp_dir();
    
    let remove_after_print = options.remove_after_print;
    let path_to_remove = options.path.clone();

    // Create a channel for communication
    let (sender, receiver) = mpsc::channel();
    
    // Spawn a new thread
    thread::spawn(move || {
        let exe_path = format!("{}sm.exe", dir.display());
        let mut cmd = Command::new(&exe_path);

        if options.id.is_empty() {
            cmd.arg("-print-to-default");
        } else {
            cmd.arg("-print-to");
            cmd.arg(&options.id);
        }

        if !options.print_setting.is_empty() {
            // Split settings by comma and then by equals sign for key-value pairs
            for part in options.print_setting.split(',') {
                cmd.arg(part.trim());
            }
        }

        cmd.arg("-silent");
        cmd.arg(&options.path);

        // Hide console window
        cmd.creation_flags(CREATE_NO_WINDOW);
        
        println!("Executing command: {:?}", cmd);

        // --- Start of fix: Detailed error reporting ---
        let output_result = cmd.output();

        let response = match output_result {
            Ok(output) => {
                if output.status.success() {
                    String::from_utf8(output.stdout).unwrap_or_else(|_| "Successfully executed, but stdout is not valid UTF-8".to_string())
                } else {
                    let stdout = String::from_utf8(output.stdout).unwrap_or_else(|_| "stdout: Invalid UTF-8".to_string());
                    let stderr = String::from_utf8(output.stderr).unwrap_or_else(|_| "stderr: Invalid UTF-8".to_string());
                    format!(
                        "Command failed with status: {}\\nStdout:\\n{}\\nStderr:\\n{}",
                        output.status, stdout, stderr
                    )
                }
            }
            Err(e) => {
                format!("Failed to execute command: {}", e)
            }
        };
        // --- End of fix ---

        sender.send(response).unwrap();
    });

    // Do other non-blocking work on the main thread

    // Receive the result from the spawned thread
    let result = receiver.recv().unwrap();
    
    if remove_after_print {
        let _ = remove_file(&path_to_remove);
    }
    
    return result;
}


/**
 * Get printer job on windows using powershell
 */
pub fn get_jobs(printername: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Get-PrintJob -PrinterName \"{}\"  | Select-Object DocumentName,Id,TotalPages,Position,Size,SubmmitedTime,UserName,PagesPrinted,JobTime,ComputerName,Datatype,PrinterName,Priority,SubmittedTime,JobStatus | ConvertTo-Json", printername)]).output().unwrap();
    return String::from_utf8(output.stdout).unwrap();
}

/**
 * Get printer job by id on windows using powershell
 */
pub fn get_jobs_by_id(printername: String, jobid: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Get-PrintJob -PrinterName \"{}\" -ID \"{}\"  | Select-Object DocumentName,Id,TotalPages,Position,Size,SubmmitedTime,UserName,PagesPrinted,JobTime,ComputerName,Datatype,PrinterName,Priority,SubmittedTime,JobStatus | ConvertTo-Json", printername, jobid)]).output().unwrap();
    return String::from_utf8(output.stdout).unwrap();
}


/**
 * Resume printers job on windows using powershell
 */
pub fn resume_job(printername: String, jobid: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Resume-PrintJob -PrinterName \"{}\" -ID \"{}\" ", printername, jobid)]).output().unwrap();
    return String::from_utf8(output.stdout).unwrap();
}

/**
 * Restart printers job on windows using powershell
 */
pub fn restart_job(printername: String, jobid: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Restart-PrintJob -PrinterName \"{}\" -ID \"{}\" ", printername, jobid)]).output().unwrap();
    return String::from_utf8(output.stdout).unwrap();
}

/**
 * pause printers job on windows using powershell
 */
pub fn pause_job(printername: String, jobid: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Suspend-PrintJob -PrinterName \"{}\" -ID \"{}\" ", printername, jobid)]).output().unwrap();
    return String::from_utf8(output.stdout).unwrap();
}

/**
 * remove printers job on windows using powershell
 */
pub fn remove_job(printername: String, jobid: String) -> String {
    let mut cmd = Command::new("powershell");
    cmd.creation_flags(CREATE_NO_WINDOW);
    let output = cmd.args([format!("Remove-PrintJob -PrinterName \"{}\" -ID \"{}\" ", printername, jobid)]).output().unwrap();
    return String::from_utf8(output.stdout).unwrap();
}
