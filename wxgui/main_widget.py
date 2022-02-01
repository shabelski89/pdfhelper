import os
import wx
from ObjectListView import ObjectListView, ColumnDefn


class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, id=wx.ID_ANY, title="Конвертер", size=(950, 500))
        self.sb = self.CreateStatusBar()
        panel = MainPanel(self)

    def change_statusbar(self, msg: str):
        self.sb.SetStatusText(f'Количество файлов: {msg}')


class MainPanel(wx.Panel):
    wildcard = "All files (*.*)|*.*|" "MS Word files (*.docx)|*.docx|" "PDF files (*.pdf)|*.pdf"
    headers = {'path': '', 'filename': '', 'extension': '', 'size': ''}

    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)
        self.parent = parent
        self.data_main_Olv = ObjectListView(self, wx.ID_ANY, style=wx.LC_REPORT | wx.SUNKEN_BORDER)
        self.data_main_Olv.SetEmptyListMsg("Выберите файлы для конвертации")
        self.current_dir = os.getcwd()
        self.set_olv_table()
        self.main_table = False

        # create the buttons and bindings
        self.open_file_dlg_btn = wx.Button(self, label="Открыть файлы...")
        self.open_file_dlg_btn.Bind(wx.EVT_BUTTON, self.on_open_file)
        self.radio1 = wx.RadioButton(self, label="docx2pdf", name='docx2pdf', style=wx.RB_GROUP)
        self.radio2 = wx.RadioButton(self, label="pdf2png", name='pdf2png', )
        self.start_btn = wx.Button(self, label="Конвертировать")
        self.start_btn.Bind(wx.EVT_BUTTON, self.on_start)
        self.start_btn.Disable()

        # Create sizers
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        button_sizer.Add(self.open_file_dlg_btn, proportion=0, flag=wx.ALL | wx.EXPAND, border=8)
        button_sizer.Add(self.radio1, proportion=0, flag=wx.ALL | wx.EXPAND, border=8)
        button_sizer.Add(self.radio2, proportion=0, flag=wx.ALL | wx.EXPAND, border=8)
        button_sizer.Add(self.start_btn, proportion=0, flag=wx.ALL | wx.EXPAND, border=8)
        main_sizer.Add(button_sizer, proportion=0, flag=wx.ALL | wx.EXPAND, border=8)

        table_sizer = wx.BoxSizer(wx.VERTICAL)
        self.data_main_OlvLabel = wx.StaticText(self, label="Файлы для обработки")
        table_sizer.Add(self.data_main_OlvLabel, proportion=0, flag=wx.ALL | wx.EXPAND, border=8)
        table_sizer.Add(self.data_main_Olv, proportion=1, flag=wx.ALL | wx.EXPAND, border=8)
        main_sizer.Add(table_sizer, proportion=1, flag=wx.ALL | wx.EXPAND, border=8)

        self.SetSizer(main_sizer)

    def set_olv_table(self):
        """
        Set empty OLV table
        """
        self.data_main_Olv.SetColumns([ColumnDefn(title='Файл', align="left", width=900, valueGetter='filename')])

    def update_table(self, event, data):
        """
        Push data to OLV table
        """
        self.data_main_Olv.SetObjects(data)
        count = self.data_main_Olv.ItemCount
        self.set_status(self, str(count))
        self.main_table = True
        self.start_btn.Enable()

    def set_status(self, event, msg):
        """
        Set count files to footer
        """
        self.parent.change_statusbar(msg=msg)

    def on_open_file(self, event):
        """
        Create and show the Open FileDialog
        """
        operation = event.GetEventObject().GetLabel()
        dlg = wx.FileDialog(
            self, message=operation,
            defaultDir=self.current_dir,
            defaultFile="",
            wildcard=self.wildcard,
            style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR
        )
        if dlg.ShowModal() == wx.ID_OK:
            data = [dict(filename=x) for x in dlg.GetPaths()]
            self.update_table(event=event, data=data)
            self.current_dir = dlg.GetDirectory()
        dlg.Destroy()

    def on_start(self, event):
        files_from_olv = [x['filename'] for x in self.data_main_Olv.GetObjects()]
        alert_msg = 'Файлы не могут быть конвертированы:\n'
        if self.radio1.GetValue():
            files = [x for x in files_from_olv if x.endswith('docx')]
            depot_files = [os.path.basename(x) for x in files_from_olv if not x.endswith('docx')]
            if depot_files:
                answer = wx.MessageBox(alert_msg + "\n".join(depot_files), "Предупреждение", wx.OK | wx.CANCEL)
                if answer == wx.OK:
                    from wxgui import docx_to_pdf
                    docx_to_pdf.main(files)
        elif self.radio2.GetValue():
            files = [x for x in files_from_olv if x.endswith('pdf')]
            depot_files = [x for x in files_from_olv if not x.endswith('pdf')]
            if depot_files:
                answer = wx.MessageBox(alert_msg + "\n".join(depot_files), "Предупреждение", wx.OK | wx.CANCEL)
                if answer == wx.OK:
                    from wxgui import pdf_to_png
                    pdf_to_png.main(files)

        #  clear table and disable button
        self.data_main_Olv.DeleteAllItems()
        self.start_btn.Disable()
