use umya_spreadsheet;
use std::path::Path;
use excle_file_merger::*;



fn main() {

    //extract all excel (.xlsx) files from the given directory
    let directory_path = "./sample_files";
    let combine_spreadsheet_path = format!("{}{}{}", directory_path, "/output/", "output.xlsx");
    let combine_spreadsheet_path = Path::new(&combine_spreadsheet_path);


    let mut book_combine = umya_spreadsheet::new_file();
    let _ = create_output_file(&mut book_combine, combine_spreadsheet_path);

    let files = get_files_in_directory(directory_path).unwrap();
    println!("files: {:?}", files);

    

    for file in files {
        println!("file: {}", file);
        let spreadsheet_path = Path::new(&file);
        let _ = insert_data_from_file(spreadsheet_path, combine_spreadsheet_path);
    }



}


