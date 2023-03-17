import win32print
import win32api

GHOSTSCRIPT_PATH = "C:\\gs\\bin\\gswin64.exe"
GSPRINT_PATH = "C:\\Ghostgum\\gsview\\gsprint.exe"
#C:\Program Files\gs\gs10.00.0\bin
# YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT
currentprinter = win32print.GetDefaultPrinter()
test = [printer[2] for printer in win32print.EnumPrinters(2)]
for i in test:
    print(i)

printer_name = 'Xerox WorkCentre 73=7245 PCL 6'
#printer_srt = '"-sstdout=filename -sOutputFile="%printer%' + printer_name + '" '
printer_srt = '"-sstdout=filename -printer "' + printer_name + '" '
file_name = 'img01644.pdf'
file_str = '"' + file_name + '"'
#Str = '-ghostscript "'+GHOSTSCRIPT_PATH+'" -printer "'+currentprinter+'" "Папка №2.pdf"' + " -color"
Str = '-ghostscript "' + GHOSTSCRIPT_PATH + printer_srt + file_str + " -color"
print(Str)

win32api.ShellExecute(0, 'open', GSPRINT_PATH, Str, '.', 0)
