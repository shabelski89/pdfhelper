"""
Provides to insert png file(s) into docx using docx
"""
__author__ = "Aleksandr Shabelsky"
__version__ = "1.0.0"
__email__ = "a.shabelsky@gmail.com"
__status__ = "Dev"

# Requirement pdf2image: pip install python-docx
# Usage python png_to_docx.py -i file1.docx=file1_png_folder

import os
import time
import docx
import docx.shared
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import argparse
from datetime import datetime
import multiprocessing as mp
from pprint import pprint


def parse_var(s):
    """
    Parse a key, value pair, separated by '='
    That's the reverse of ShellArgs.
    On the command line (argparse) a declaration will typically look like:
        foo=hello
    """
    items = s.split('=')
    key = items[0].strip()
    if len(items) > 1:
        value = '='.join(items[1:])
        return key, value


def parse_vars(items):
    """
    Parse a series of key-value pairs and return a dictionary
    """
    d = {}
    if items and isinstance(items, list):
        for item in items:
            try:
                key, value = parse_var(item)
                d[key] = value
            except TypeError as e:
                print(help_msg)
                exit(1)
    return d


def date_time():
    """
    Function return string formatted datetime
    :return: str
    """
    return datetime.now().strftime('%d.%m.%Y %H:%M:%S')


def get_files_list(path: str):
    return [os.path.join(path, f) for f in os.listdir(path)]


def get_sorted_dict(files: list):
    img_dict = {}
    for filename in files:
        if filename.endswith('.png'):
            num_index = filename.rfind('-') + 1
            ext_index = filename.rfind('.png')
            num = filename[num_index:ext_index]
            try:
                n = int(num)
                img_dict[n] = filename
            except Exception as E:
                print(E)
    return img_dict


def png2docx(docx_file: str, png_files: dict):
    doc = docx.Document(docx_file)

    picture_width = 148
    picture_height = 210
    pictures_dirname = os.path.dirname(png_files[1])

    picture_num = 0
    paragraph_index = 0

    for paragraph in doc.paragraphs:
        paragraph_index += 1

        if 1 < paragraph_index < 4:
            continue

        if picture_num >= len(png_files):
            if picture_num == 0:
                paragraph.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)

            paragraph = paragraph._element
            paragraph.getparent().remove(paragraph)
            paragraph._p = paragraph._element = None

        else:
            paragraph.alignment = docx.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
            my_run = paragraph.add_run()
            my_run.add_picture(png_files[picture_num + 1],
                               width=docx.shared.Mm(picture_width),
                               height=docx.shared.Mm(picture_height))

        picture_num += 1

    for table in doc.tables:
        paragraph = doc.add_paragraph()
        table._element.addprevious(paragraph._p)
        run = paragraph.add_run()
        run.add_break(docx.enum.text.WD_BREAK.PAGE)

    export_path = os.path.join(pictures_dirname, "Edited")

    try:
        os.makedirs(export_path, exist_ok=True)
    except Exception as E:
        print(E)

    docx_rendered_file = os.path.join(export_path, os.path.basename(docx_file).replace('A4(src)', ''))
    doc.save(docx_rendered_file)
    return docx_rendered_file


def main(arguments):
    """
    Main function
    :param arguments: cli arguments
    """
    print('#' * 100)
    print(f'[{date_time()}] - Program start')
    start = time.time()

    #  make dicts {file.docx:{number: png_file_number}}
    compared_files = {k: get_sorted_dict(get_files_list(v)) for k, v in arguments.items()
                      if os.path.exists(k) and os.path.exists(v) and get_sorted_dict(get_files_list(v))}
    if compared_files.values():

        print(f'[{date_time()}] - Insert PNG to DOCX')
        pprint(compared_files)

        for f, img in compared_files.items():
            with mp.Pool() as pool:
                converted_files = pool.starmap(png2docx, [(f, img)])

        for file in converted_files:
            print(f'[{date_time()}] - Converted file - {file}')
    else:
        print(f'[{date_time()}] - There are no files to converting')
        pprint(arguments)

    end_time = int(time.time() - start)
    print(f"[{date_time()}] - Time elapsed {end_time} seconds")
    print(f'[{date_time()}] - Program end')
    print('#' * 100)


if __name__ == '__main__':
    # command line argument parser with help message
    desc_msg = "png2docx"
    help_msg = """
    Set a number of key-value pairs 
    (do not put spaces before or after the = sign). 
    If a value contains spaces, you should define it with double quotes: 
    'file1.docx=file1_png_images_folder' 
    Note that values are always treated as strings.
    """
    arg_parser = argparse.ArgumentParser(description=desc_msg, formatter_class=argparse.RawTextHelpFormatter)
    arg_parser.add_argument("-i", dest="input", metavar="KEY=VALUE", nargs='+', help=help_msg)
    args = arg_parser.parse_args()
    values = parse_vars(args.input)
    if values:
        main(values)
    else:
        print(help_msg)
