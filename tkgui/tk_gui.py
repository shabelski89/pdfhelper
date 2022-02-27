import tkinter as tk
from tkgui.tk_widget import InterFace


def main():
    root = tk.Tk()
    application = InterFace(root)
    application.root.mainloop()


if __name__ == "__main__":
    main()
