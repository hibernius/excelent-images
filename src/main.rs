use clap::{Arg, Command};
use image::io::Reader as ImageReader;
use image::{DynamicImage, GenericImageView, ImageError};
use rust_xlsxwriter::{ColNum, Format, RowNum, Workbook};
use std::path::{Path, PathBuf};

fn main() {
    let opts = get_opts();

    let input_filepath = opts
        .get_one::<String>("input_file")
        .expect("Error: Input file argument is missing. Please provide a valid file path.");
    let output_directory = opts
        .get_one::<String>("output_directory")
        .expect("Error: Output argument is missing. Please provide a valid output directory name.");

    let image = match get_image(input_filepath) {
        Ok(img) => img,
        Err(error) => {
            eprintln!("Error opening or decoding image: {:?}", error);
            return;
        }
    };

    let (width, height) = get_image_dimensions(image.clone());
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for col in 0..width {
        worksheet
            .set_column_width(col as ColNum, 2)
            .expect("Error setting column width.");
        for row in 0..height {
            let pixel = image.get_pixel(col, row);
            let color = format!("#{:02X}{:02X}{:02X}", pixel[0], pixel[1], pixel[2]);
            let format = Format::new().set_background_color(&*color);
            let _ = worksheet.write_string_with_format(row as RowNum, col as ColNum, ' ', &format);
        }
    }

    let filename = get_filename(input_filepath).unwrap_or_else(|| "unknown".to_string());
    replace_extension(&filename, "xlsx");
    let output_directory = enforce_trailing_slash(output_directory);

    workbook
        .save(format!("{}{}", output_directory, filename))
        .expect("Error saving workbook.");
}

fn get_opts() -> clap::ArgMatches {
    Command::new("ExcelentImages")
        .about("Converts an image file to an Excel file.")
        .arg(
            Arg::new("input_file")
                .short('f')
                .long("file")
                .value_name("INPUT_FILE")
                .help("Path to the image file")
                .value_parser(clap::value_parser!(String))
                .required(true),
        )
        .arg(
            Arg::new("output_directory")
                .short('d')
                .long("dir")
                .value_name("OUTPUT_DIR")
                .default_value("output/")
                .help("Output directory path (include trailing slash)")
                .value_parser(clap::value_parser!(String))
                .required(false),
        )
        .get_matches()
}

fn get_image(path: &str) -> Result<DynamicImage, ImageError> {
    let reader = ImageReader::open(path)?;
    let image = reader.decode()?;
    Ok(image)
}

fn get_image_dimensions(image: DynamicImage) -> (u32, u32) {
    image.dimensions()
}

fn get_filename(path: &str) -> Option<String> {
    let path = Path::new(path);
    path.file_stem()
        .map(|os_str| os_str.to_string_lossy().to_string())
}

fn replace_extension(path: &str, new_extension: &str) -> String {
    let mut path_buf = PathBuf::from(path);
    path_buf.set_extension(new_extension);
    path_buf.to_string_lossy().to_string()
}

fn enforce_trailing_slash(dir_path: &str) -> String {
    if dir_path.ends_with('/') {
        dir_path.to_string()
    } else {
        format!("{}/", dir_path)
    }
}
