use std::{ffi::OsStr,  fs, io,};
use umya_spreadsheet::{self, reader::{self}, writer, Spreadsheet};
use chrono::Local;
static mut CURRENT_FILE_NUMBER :u32 = 0 ;





pub fn create_output_file ( output_ss: &mut Spreadsheet,  output_ss_path: &std::path::Path,  ) -> Result<(), writer::xlsx::XlsxError> {
    
    if output_ss_path.exists() {
        let _ = fs::remove_file(output_ss_path);
    } else {
        let _ = fs::create_dir_all(output_ss_path.parent().unwrap());
    }

    let combine_ss_ws = output_ss.get_sheet_mut(&0).unwrap();
    
    combine_ss_ws.get_cell_value_mut((1, 1)).set_value("Date Modified");
    combine_ss_ws.get_cell_value_mut((2, 1)).set_value("Number of Files");
    combine_ss_ws.get_cell_value_mut((3, 1)).set_value("Series Number");
    combine_ss_ws.get_cell_value_mut((4, 1)).set_value("Counter Number for Each File");
    combine_ss_ws.get_cell_value_mut((5, 1)).set_value("File Name");

    writer::xlsx::write(output_ss, output_ss_path)

}

pub fn get_files_in_directory(path: &str) -> io::Result<Vec<String>> {
    // Get a list of all entries in the folder
    let entries = fs::read_dir(path)?;
  
    // Extract the filenames from the directory entries and store them in a vector
    let file_names: Vec<String> = entries
        .filter_map(|entry| {
            let path = entry.ok()?.path();
            if path.is_file() && path.extension().and_then(OsStr::to_str) == Some("xlsx"){
                path.to_str().map(|s| s.to_owned())
            } else {

                None
            }
        })
        .collect();

    Ok(file_names)
}

pub fn insert_data_from_file (from_file: &std::path::Path, to_file: &std::path::Path) -> Result<(), std::io::Error> {

    unsafe { CURRENT_FILE_NUMBER = CURRENT_FILE_NUMBER + 1; }
    let mut count_number_for_each_file: String = unsafe {CURRENT_FILE_NUMBER.to_string()+"-"+ &1.to_string()};
    let  file_name = from_file.file_name().unwrap().to_owned().into_string().unwrap();

    let mut output_ss = reader::xlsx::read(to_file).unwrap();
    let output_ws = output_ss.get_sheet_mut(&0).unwrap();
     
     let input_ss = reader::xlsx::read(from_file).unwrap();
     let input_ws = input_ss.get_sheet(&0).unwrap();

    let (output_ws_highest_col_index, output_ws_highest_row_index) = output_ws.get_highest_column_and_row();
    let (input_ws_highest_col_index, input_ws_highest_row_index) = input_ws.get_highest_column_and_row();

    let mut output_ws_current_row_index: u32 = 1;
    let mut output_ws_current_col_index: u32 = 5;

    let mut input_ws_current_col_index = 0;
    let mut input_ws_current_row_index = 1;
    

    
    if output_ws_highest_row_index != 1 {
        output_ws_current_row_index = output_ws_highest_row_index+1;
    
    } 
    

    println!("mergin file: {:?}, .............", from_file );
    println!("combine worksheet current dimension is: {:?}", (output_ws_highest_col_index, output_ws_highest_row_index));
    println!("merging file dimenstion is: {:?}" , (input_ws_highest_col_index, input_ws_highest_row_index));

    
    

    //insert data from file into combine spreadsheet
    for cell in  input_ws.get_cell_value_by_range(&input_ws.calculate_worksheet_dimension()) {

        

        

        if input_ws_current_col_index < input_ws_highest_col_index {
            input_ws_current_col_index += 1;
        } else {
            input_ws_current_col_index = 1;
            input_ws_current_row_index +=1;

        }

        
        if output_ws_highest_row_index ==1 || input_ws_current_row_index > 1  {

            if output_ws_current_col_index < input_ws_highest_col_index+5 {
                output_ws_current_col_index += 1;
               
            } else {
                output_ws_current_col_index = 6;
                output_ws_current_row_index += 1;
                count_number_for_each_file = unsafe {CURRENT_FILE_NUMBER.to_string()+"-"+ &(( output_ws_current_row_index  -  output_ws_highest_row_index).to_string()) };
            }

            if output_ws_current_row_index > 1 {

            let  now: chrono::DateTime<Local> = chrono::Local::now();
            output_ws.get_cell_value_mut((1, output_ws_current_row_index)).set_value(now.format("%Y-%m-%d %T").to_string());
            unsafe {
            output_ws.get_cell_value_mut((2, output_ws_current_row_index)).set_value_number(CURRENT_FILE_NUMBER);
            }
            output_ws.get_cell_value_mut((3, output_ws_current_row_index)).set_value_number(output_ws_current_row_index-1); 
            output_ws.get_cell_value_mut((4, output_ws_current_row_index)).set_value(&count_number_for_each_file);
            output_ws.get_cell_value_mut((5, output_ws_current_row_index)).set_value(&file_name);
            }
        
            output_ws.get_cell_value_mut((output_ws_current_col_index, output_ws_current_row_index)).set_value(cell.get_value());
        }

        println!("ws dimensions: ( {}, {})", input_ws_current_col_index, input_ws_current_row_index);
        println!("combine ws dimensions: ( {}, {})", output_ws_current_col_index, output_ws_current_row_index);
    }
    

    println!("combine worksheet new dimension is: {:?}", output_ws.get_highest_column_and_row());
    let _ = writer::xlsx::write(&output_ss, to_file);
   

    Ok(())
}