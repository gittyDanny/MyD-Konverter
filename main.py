"""MYD SAP Migration Checker.By Daniil Ioffe und Claudi(nicht im Auftrag von PwC)"""
import tkinter as tk
from gui.app import XMLtoExcelApp

if __name__ == '__main__':
    root = tk.Tk()
    app = XMLtoExcelApp(root)
    root.mainloop()
