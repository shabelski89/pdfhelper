"""
Provides to convert docx file(s) into pdf using docx2pdf
"""
__author__ = "Aleksandr Shabelsky"
__version__ = "1.0.1"
__email__ = "a.shabelsky@gmail.com"
# Requirement docx2pdf: pip install docx2pdf
# Usage python docx_to_pdf.py -i file1.docx file2.docx

import os
import time
import argparse
from datetime import datetime
from docx2pdf import convert
import multiprocessing as mp
from functools import partial


def date_time():
    """
    Function return string formatted datetime
    :return: str
    """
    return datetime.now().strftime('%d.%m.%Y %H:%M:%S')


def flatten_list(nested_list):
    """
    Function to flatten nested lists
    :param nested_list: list of nested lists
    :return: list
    """
    if not(bool(nested_list)):
        return nested_list
    if isinstance(nested_list[0], list):
        return flatten_list(*nested_list[:1]) + flatten_list(nested_list[1:])
    return nested_list[:1] + flatten_list(nested_list[1:])


def convert_docx2pdf(file, **kwargs):
    """
    Function to convert docx file to pdf
    :param file: name of docx file
    :return: pdf file
    """
    file_dirname = os.path.dirname(file)
    filename = os.path.basename(file)
    output_path_pdf = os.path.join(file_dirname, filename.replace('.docx', ''), 'pdf')
    try:
        os.makedirs(output_path_pdf, exist_ok=True)
    except Exception as E:
        print(E)
    export_filename = filename.replace('docx', 'pdf')
    full_pdf_filename = os.path.join(output_path_pdf, export_filename)
    try:
        convert(file, full_pdf_filename)
        return full_pdf_filename
    except Exception as E:
        print(E)


def main(files, **kwargs):
    """
    Main function
    :param files: files to convert from pdf to png
    :param kwargs:
    """
    print('#' * 100)
    print(f'[{date_time()}] - Program start')
    start = time.time()

    #  make list of files is flat
    files2convert = [file for file in flatten_list(files) if file.endswith('.docx') and os.path.exists(file)]
    if files2convert:
        for file in files2convert:
            print(f'[{date_time()}] - Converting file - {file}')

        multiprocessing = kwargs.get('multiprocessing', False)
        if multiprocessing:
            with mp.Pool() as pool:
                print(f'[{date_time()}] - Converting file with multi processing')
                converted_files = pool.map(partial(convert_docx2pdf, **kwargs), files2convert)
        else:
            print(f'[{date_time()}] - Converting file with single processing')
            converted_files = [convert_docx2pdf(file=file, **kwargs) for file in files2convert]

        for file in converted_files:
            print(f'[{date_time()}] - Converted file - {file}')
    else:
        print(f'[{date_time()}] - There are no files to converting in {files2convert}')
    end_time = int(time.time() - start)
    print(f"[{date_time()}] - Time elapsed {end_time} seconds")
    print(f'[{date_time()}] - Program end')
    print('#' * 100)


if __name__ == '__main__':
    # command line argument parser with help message
    desc_msg = "docx2pdf"
    help_msg = "Docx files converting to PDF files"
    arg_parser = argparse.ArgumentParser(description=desc_msg, formatter_class=argparse.RawTextHelpFormatter)
    arg_parser.add_argument("-i", dest="input", required=True, nargs='+', action='append', help=help_msg)
    arg_parser.add_argument("-m", dest="multiprocessing", required=False, type=bool, default=False)
    args = arg_parser.parse_args()
    main(args.input)
