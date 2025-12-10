import tkinter as tk
from tkinter import messagebox
import sys

print("Testing tkinter with PyInstaller...")

root = tk.Tk()
root.withdraw()  # Hide main window

messagebox.showinfo("Test", f"Python: {sys.version}\nTkinter test successful!")

print("Test completed")
root.destroy()
