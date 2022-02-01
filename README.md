### Установка


```shell
python setup.py install
```

### Ручная сборка
```shell
    pip install docx2pdf
    pip install pdf2image
    pip install wxpython
    pip install objectlistview
```

### Запуск с GUI
`python pydocs.py`

Выбираются один или несколько фалов для конвертации и тип конвертации файла:
* docx2pdf
* pdf2png


### Запуск без GUI
###### docx2pdf
`python docx_to_pdf.py -i file1.docx file2.docx`

Один обязательный аргумент **-i** с указанием одного или нескольких файлов для конвертации.

###### pdf2png
`python pdf_to_png.py -i file1.pdf file2.pdf`

Один обязательный аргумент **-i** с указанием одного или нескольких файлов для конвертации.
