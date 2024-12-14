use image::io::Reader as ImageReader;
use image::{DynamicImage, GenericImageView, ImageError};
use rust_xlsxwriter::{ColNum, Format, RowNum, Workbook};

fn main() {
    let image = open_image("image-name.ext");
    let mut workbook = Workbook::new();

    let image = match image {
        Ok(img) => img,
        Err(error) => {
            eprintln!("Error opening or decoding image: {:?}", error);
            return;
        }
    };

    let (width, height) = get_image_dimensions(image.clone());
    let worksheet = workbook.add_worksheet();

    for col in 0..width {
        worksheet.set_column_width(col as ColNum, 2).expect("Error setting column width.");
        for row in 0..height {
            let pixel = image.get_pixel(col, row);
            let color = format!("#{:02X}{:02X}{:02X}", pixel[0], pixel[1], pixel[2]);
            let format = Format::new().set_background_color(&*color);
            let _ = worksheet.write_string_with_format(row as RowNum, col as ColNum, ' ', &format);
        }
    }

    workbook.save("output.xlsx").expect("Error saving workbook.");
}

fn open_image(path: &str) -> Result<DynamicImage, ImageError> {
    return ImageReader::open(path)?.decode();
}

fn get_image_dimensions(image: DynamicImage) -> (u32, u32) {
    return image.dimensions();
}

