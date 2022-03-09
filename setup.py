from setuptools import setup

setup(
    name='pdfhelper',
    version='1.1.0',
    packages=['wxgui'],
    url='https://github.com/shabelski89/pdfhelper',
    license='',
    author='Aleksandr Shabelsky',
    author_email='a.shabelsky@gmail.com',
    description='GUI to convert docx2pdf and pdf2png',
    install_requires=['docx2pdf', 'pdf2image', 'wxpython', 'objectlistview'],
    python_requires='>=3.6'
)
