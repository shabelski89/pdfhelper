"""
Provides to convert pdf file(s) into png image using pdf2image
"""
__author__ = "Aleksandr Shabelsky"
__version__ = "1.0.0"
__email__ = "a.shabelsky@gmail.com"
# Requirement pdf2image: pip install pdf2image
# Usage python pdf_to_png.py -i file1.pdf file2.pdf

import os
import time
import argparse
from datetime import datetime
from pdf2image import convert_from_path
import multiprocessing as mp
from tqdm import tqdm


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


def convert_pdf2png(file):
    """
    Function to convert pdf file to png images
    :param file: name of pdf file
    :return: image files
    """
    pbar = tqdm(total=1)
    file_dirname = os.path.dirname(file)
    filename = os.path.basename(file)
    output_path_png = os.path.join(file_dirname, filename.replace('.pdf', ''), 'png')
    try:
        os.makedirs(output_path_png, exist_ok=True)
    except Exception as E:
        print(E)
    try:
        images = convert_from_path(file)

        output_files = []
        for i in range(len(images)):
            out_file = os.path.join(output_path_png, f"{filename.replace('.pdf', '')}_{i}.png")
            images[i].save(out_file, 'PNG')
            output_files.append(out_file)
        pbar.update(1)
        return output_files
    except Exception as E:
        print(E)


def main(arguments):
    """
    Main function
    :param arguments: cli arguments
    """
    print('#' * 100)
    print(f'[{date_time()}] - Program start')
    start = time.time()

    #  make list of files is flat
    files = [file for file in flatten_list(arguments) if file.endswith('.pdf') and os.path.exists(file)]
    if files:
        for file in files:
            print(f'[{date_time()}] - Converting file - {file}')

        with mp.Pool() as pool:
            converted_files = pool.map(convert_pdf2png, files)

        for files in converted_files:
            print(f'[{date_time()}] - Converted {len(files)} files to folder - {os.path.dirname(files[0])}')
    else:
        print(f'[{date_time()}] - There are no files to converting in {arguments}')
    end_time = int(time.time() - start)
    print(f"[{date_time()}] - Time elapsed {end_time} seconds")
    print(f'[{date_time()}] - Program end')
    print('#' * 100)


if __name__ == '__main__':
    # command line argument parser with help message
    desc_msg = "pdf2png"
    help_msg = "PDF files converting to PNG files"
    arg_parser = argparse.ArgumentParser(description=desc_msg, formatter_class=argparse.RawTextHelpFormatter)
    arg_parser.add_argument("-i", dest="input", required=True, nargs='+', action='append', help=help_msg)
    args = arg_parser.parse_args()
    main(args.input)
