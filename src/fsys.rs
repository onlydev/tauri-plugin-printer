
use std::fs::{File, remove_file as rmf};

use std::io::Write;
use std::path::Path;
use base64::{Engine as _, engine::{general_purpose}};


pub fn create_file_from_base64(base64_string: &str, file_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut buffer = Vec::<u8>::new();
    general_purpose::STANDARD
        .decode_vec(base64_string, &mut buffer,)?;

    let path = Path::new(file_path);
    let mut file = File::create(&path)?;

    file.write_all(&buffer)?;

    Ok(())
}

pub fn remove_file (file_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    rmf(file_path)?;
    Ok(())
}
