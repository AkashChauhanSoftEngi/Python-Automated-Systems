#from win32 import win32api
from tkinter import *
#import tkinter.filedialog
from tkinter import filedialog
import os
import xlrd, xlwt
import pandas as pd
import xlutils.filter
from xlutils.copy import copy as xl_copy
import tkinter as tk
import openpyxl
import re
import string

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title='xxxxxxx')
#Rest content deleted!
