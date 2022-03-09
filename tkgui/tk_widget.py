import os
from tkinter import *
from tkinter.ttk import Treeview, Style
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from converter import docx_to_pdf, pdf_to_png, png_to_docx


class InterFace:
    def __init__(self, main):
        self.root = main
        self.initialize_user_interface()

    def initialize_user_interface(self):
        self.root.title("Конвертер")
        self.root.geometry("950x500")
        self.root.resizable(False, False)
        self.root.config()

        self.root.frame_top = Frame()
        self.root.frame_bottom = Frame()
        self.root.frame_top.pack(side='top', padx=2)
        self.root.frame_bottom.pack(side='bottom', padx=2, expand=1, fill=Y, anchor=CENTER)

        self.root.btn_open = Button(self.root.frame_top, text="Открыть файлы...", command=self.select_files)
        self.root.btn_open.grid(row=0, column=0, pady=6, padx=6, sticky='w')

        self.root.radio_var = StringVar()
        self.root.radio_btn_1 = Radiobutton(self.root.frame_top, text='docx2pdf',
                                            variable=self.root.radio_var, value='docx2pdf')
        self.root.radio_btn_1.grid(row=0, column=1, pady=6, padx=2)
        self.root.radio_btn_2 = Radiobutton(self.root.frame_top, text='pdf2png',
                                            variable=self.root.radio_var, value='pdf2png')
        self.root.radio_btn_2.grid(row=0, column=2, pady=6, padx=2)
        self.root.radio_btn_3 = Radiobutton(self.root.frame_top, text='png2docx',
                                            variable=self.root.radio_var, value='png2docx')
        self.root.radio_btn_3.grid(row=0, column=3, pady=6, padx=2)
        self.root.radio_var.set('docx2pdf')

        self.root.btn_start = Button(self.root.frame_top, text="Конвертировать", command=self.on_start, state=DISABLED)
        self.root.btn_start.grid(row=0, column=4, pady=6, padx=6)

        self.root.var = IntVar()
        self.root.check_box = Checkbutton(self.root.frame_top, text="Multiprocessing", variable=self.root.var)
        self.root.check_box.grid(row=0, column=5, pady=6, padx=6)

        self.root.style = Style(self.root)
        self.root.style.configure('Treeview', rowheight=36)

        self.root.tree = Treeview(self.root.frame_bottom, height=11, columns="f", selectmode=BROWSE, show='headings')

        self.root.tree.column(f'#1', minwidth=800, width=900, stretch=YES, anchor=W)
        self.root.tree.heading(f'#1', text="Файл")

        self.root.tree.grid(row=10, columnspan=5, sticky='nsew')
        self.root.treeview = self.root.tree

    def select_files(self):
        self.root.tree.delete(*self.root.tree.get_children())
        filetypes = (('All files', '*.*'), ('pdf files', '*.pdf'), ('word files', '*.docx'))
        filenames = fd.askopenfilenames(title='Открыть файлы...', initialdir='/', filetypes=filetypes)
        for el in filenames:
            self.root.treeview.insert('', 'end', text='', values=el)
        self.root.btn_start["state"] = NORMAL
        self.root.update()

    def on_start(self):
        files_from_view = [" ".join([str(x) for x in self.root.tree.item(child)["values"]])
                           for child in self.root.tree.get_children()]
        alert_msg = 'Файлы не могут быть конвертированы:\n'
        result_msg = 'Файлы cконвертированы:\n'
        extension = {'docx2pdf': 'docx', 'pdf2png': 'pdf', 'png2docx': 'docx'}
        apps = {'docx2pdf': docx_to_pdf.main, 'pdf2png': pdf_to_png.main, 'png2docx': png_to_docx.main}
        ext = extension[self.root.radio_var.get()]
        app = apps[self.root.radio_var.get()]

        files = [x for x in files_from_view if x.endswith(ext)]
        depot_files = [os.path.basename(x) for x in files_from_view if not x.endswith(ext)]

        if depot_files:
            showinfo(title='Предупреждение', message=alert_msg + "\n".join(depot_files))

        if files:
            multiprocessing = bool(self.root.var.get())
            app(files, multiprocessing=multiprocessing)
            showinfo(title='Предупреждение', message=result_msg + "\n".join(files))

        self.root.tree.delete(*self.root.tree.get_children())
        self.root.btn_start["state"] = DISABLED


if __name__ == '__main__':
    root = Tk()
    application = InterFace(root)
    application.root.mainloop()
