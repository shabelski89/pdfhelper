"""
Provides to convert docx file(s) into pdf using docx2pdf
"""
__author__ = "Aleksandr Shabelsky"
__version__ = "0.9.0"
__email__ = "a.shabelsky@gmail.com"

import cmd
import sys


class PydocsCmd(cmd.Cmd):
    intro = 'Welcome to the pydocs shell.\nType help or ? to list commands.\n'
    prompt = '(pydocs) '

    def do_wx(self, arg):
        """Run pydocs in wxPython GUI"""
        from wxgui import wx_gui
        wx_gui.main()
        return self

    def do_tk(self, line):
        """Run pydocs in Tkinter GUI"""
        return self

    def do_exit(self, line):
        """Exit and close program"""
        return self


if __name__ == '__main__':
    my_cmd = PydocsCmd()

    if len(sys.argv) > 1:
        my_cmd.onecmd(' '.join(sys.argv[1:]))
    else:
        my_cmd.cmdloop()
