import os
import warnings
import sys
import time
import click
import pytesseract
import fitz
from PIL import Image
import xlsxwriter
import shutil


# Uncomment for Windows
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


@click.command()
@click.argument("input_path")
@click.argument("output_path")
def main(input_path: str, output_path: str) -> None:
    """
    Orchestrates the execution of the PDF text extraction process.

    Parameters:
    input_path (str): The path to the directory containing the input PDF files.
    output_path (str): The path to the directory where the output Excel file
    will be saved.

    Returns:
    None
    """

    files = prepare_files(input_path, output_path)
    extract_text(input_path, output_path, files)

    click.echo(f"Your data is ready in {output_path}")


def prepare_files(input_path: str, output_path: str) -> list:
    """
    Takes in the input and output directory paths and returns a list of PDF
    files in the input directory.

    Parameters:
    input_path (str): The directory path for input files.
    output_path (str): The directory path for output files.

    Returns:
    list: A list of PDF files in the input directory.

    Raises:
    ValueError: If the input path is not valid or there are no PDF files
    in the input directory.
    """

    # Validate input path
    if not os.path.exists(input_path):
        raise ValueError("Please use a valid input directory path.")

    # Handle output path senarios
    if not os.path.exists(output_path):
        try:
            os.mkdir(output_path)
        except:
            raise ValueError("Please use a valid output directory path.")

    # Raise warining for typical input and output paths
    if input_path == output_path:
        warnings.warn("You are using the same path as the input and output folder.")

    # Keep pdf files only
    input_files = os.listdir(input_path)
    files = [file for file in input_files if file[-3:].lower() == "pdf"]

    # Check if PDF files exist
    if len(files) == 0:
        raise ValueError("No PDF files in the input directory.")

    return files


def extract_text(input_path: str, output_path: str, files: list) -> None:
    """
    Takes in the input and output directory paths and a list of PDF files,
    extracts the text from the PDF files, and writes the text to an Excel file.

    Parameters:
    input_path (str): The directory path for input files.
    output_path (str): The directory path for output files.
    files (list): A list of PDF files in the input directory.

    Returns:
    None.

    Raises:
    None.
    """
    # Show process start
    sys.stdout.write("Validating your files...\n")
    # Get seperation code parts
    code = "4444XUJY76TFG543ED67"
    code_parts = [code[i : i + 6] for i in range(len(code) - 6)]

    # Run the writer
    df_path = os.path.join(output_path, "extracted_data.xlsx")
    with xlsxwriter.Workbook(
        df_path,
        {"strings_to_formulas": False, "constant_memory": True, "encoding": "utf-8"},
        # {"strings_to_formulas": False, "encoding": "utf-8"},
    ) as workbook:
        worksheet = workbook.add_worksheet()
        # Customize row and column size
        worksheet.set_column_pixels(0, 0, 300)
        worksheet.set_column_pixels(0, 1, 300)
        worksheet.set_default_row(200)
        # Set row height for row 0
        worksheet.set_row(0, 25)
        # Write header
        worksheet.write(0, 0, "File Name")
        worksheet.write(0, 1, "Content")

        # Get files
        files_len = len(files)
        sys.stdout.write("Running the OCR...\n")
        # Read the files and run OCR
        count = 1  # Excel files counter
        for f in range(files_len):
            try:
                img_list, temp_list, text_list, = (
                    [],
                    [],
                    [],
                )
                file_path = os.path.join(input_path, files[f])
                doc = fitz.open(file_path)

                for p in range(len(doc)):
                    page = doc.load_page(p)
                    pix = page.get_pixmap(dpi=200)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    # Write "" if OCR is failed
                    try:
                        tess_txt = pytesseract.image_to_string(img, lang="eng+ara")
                    except:
                        tess_txt = ""

                    # Check if the code exists in current page
                    tess_text_len = len(tess_txt)
                    if (
                        (tess_text_len < 25)
                        and (tess_text_len > 10)
                        and any(part for part in code_parts if part in tess_txt)
                    ):
                        img_list.append(None)
                        text_list.append(None)
                    else:
                        img_list.append(img)
                        text_list.append(tess_txt[:-1])

                # Check if the file will not be split
                none_idxs = [i for i, n in enumerate(img_list) if n is None]

                # Split file data
                worksheet, count = split_data(
                    img_list,
                    temp_list,
                    output_path,
                    none_idxs,
                    files[f],
                    text_list,
                    worksheet,
                    count,
                )
                # Copy file if not split
                if len(none_idxs) == 0:
                    shutil.copy(file_path, os.path.join(output_path, files[f]))

            except:
                sys.stdout.write(f"\nReading or writing {files[f]} has failed.\n")
            # Print progress bar
            time.sleep(0.01)
            sys.stdout.write("\r")
            sys.stdout.write(
                "Your files are being processed... {:.0f}%".format(
                    (((f + 1) / (files_len)) * 100)
                )
            )
            sys.stdout.flush()
    sys.stdout.write("\n")


def write_files(
    temp_list: list,
    img_list: list,
    text_list: list,
    file: str,
    output_path: str,
    i: str,
    worksheet,
    count,
) -> tuple:
    """
    Writes a PDF file using the provided lists of image and text data.

    Parameters:
    temp_list (list): A list of integers representing the indices of images
    to be used in the PDF.
    img_list (list): A list of images to be included in the PDF.
    text_list (list): A list of strings representing the text in the PDF.
    file (str): The name of the input file.
    output_path (str): The path to the output directory.
    i (int): The current index for naming the output file.
    worksheet: An Excel worksheet object.
    count (int): An integer representing the current row in the Excel worksheet.

    Returns:
    tuple.

    Raises:
    None.
    """

    new_fname = f"{file[:-4]}-{i+1}.pdf"
    pdf_output_path = os.path.join(output_path, new_fname)
    temp_list_len = len(temp_list)
    # Handle saving cases based on the number of pages
    if temp_list_len == 0:
        pass
    elif temp_list_len == 1:
        img_list[temp_list[0]].save(pdf_output_path)
        txt = [text_list[temp_list[0]]]
    else:
        if temp_list_len == 2:
            img_list[temp_list[0]].save(
                pdf_output_path, save_all=True, append_images=[img_list[temp_list[1]]]
            )
        elif temp_list[-1] == (len(img_list) - 1):
            img_list[temp_list[0]].save(
                pdf_output_path, save_all=True, append_images=img_list[temp_list[1] :]
            )
        else:
            img_list[temp_list[0]].save(
                pdf_output_path,
                save_all=True,
                append_images=img_list[temp_list[1] : temp_list[-1] + 1],
            )
        txt = text_list[temp_list[0] : temp_list[-1] + 1]
    # Write data based on number of pages
    if temp_list_len > 0:
        for i in range(len(txt)):
            worksheet.write(count, 0, new_fname)
            worksheet.write(count, 1, txt[i])
            count += 1

    return worksheet, count


def split_data(
    img_list: list,
    temp_list: list,
    output_path: str,
    none_idxs: list,
    file: str,
    text_list: list,
    worksheet,
    count,
) -> tuple:
    """
    Splits data and writes PDF files using the provided lists of image and text data.

    Parameters:
    img_list (list): A list of images to be included in the PDF.
    temp_list (list): A list of integers representing the indices of images to be
    used in the PDF.
    output_path (str): The path to the output directory.
    none_idxs (list): list of None elements that exist in image list.
    file (str): The name of the input file.
    text_list (list): A list of strings representing the text to be included in the PDF.
    worksheet: An Excel worksheet object.
    count: An integer representing the current row in the Excel worksheet.

    Returns:
    tuple.

    Raises:
    None.
    """
    temp_list = []
    # Drop None if in the beginning or end of lists.
    # Code must be ignored in both cases.
    if img_list[0] is None:
        del img_list[0]
        del text_list[0]
        del none_idxs[0]

    if img_list[-1] is None:
        del img_list[-1]
        del text_list[-1]
        del none_idxs[-1]
    none_idxs_len = len(none_idxs)
    # Write text as is if no separator. Else, split the files.
    if none_idxs_len == 0:
        for i in range(len(text_list)):
            worksheet.write(count, 0, file)
            worksheet.write(count, 1, text_list[i])
            count += 1
    else:
        img_list_len = len(img_list)
        for i, n in enumerate(none_idxs):
            if i == 0:
                temp_list = [j for j in range(0, n)]
                worksheet, count = write_files(
                    temp_list,
                    img_list,
                    text_list,
                    file,
                    output_path,
                    i,
                    worksheet,
                    count,
                )
            elif i == none_idxs_len - 1:
                temp_list = [j for j in range((none_idxs[i - 1]) + 1, n)]
                worksheet, count = write_files(
                    temp_list,
                    img_list,
                    text_list,
                    file,
                    output_path,
                    i,
                    worksheet,
                    count,
                )
                temp_list = [j for j in range(n + 1, img_list_len)]
                worksheet, count = write_files(
                    temp_list,
                    img_list,
                    text_list,
                    file,
                    output_path,
                    i + 1,
                    worksheet,
                    count,
                )
            else:
                temp_list = [j for j in range((none_idxs[i - 1]) + 1, n)]
                worksheet, count = write_files(
                    temp_list,
                    img_list,
                    text_list,
                    file,
                    output_path,
                    i,
                    worksheet,
                    count,
                )

    return worksheet, count


if __name__ == "__main__":
    main()
