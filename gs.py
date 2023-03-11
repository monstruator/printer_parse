import win32print
import win32api

GHOSTSCRIPT_PATH = "C:\\printer_parse\\gs\\bin\\gswin64.exe"
GSPRINT_PATH = "C:\\printer_parse\\Ghostgum\\gsview\\gsprint.exe"
#C:\Program Files\gs\gs10.00.0\bin
# YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT
currentprinter = win32print.GetDefaultPrinter()
Str = '-ghostscript "'+GHOSTSCRIPT_PATH+'" -printer "'+currentprinter+'" "Папка №2.pdf"' + " -color"
print(Str)

win32api.ShellExecute(0, 'open', GSPRINT_PATH, Str, '.', 0)