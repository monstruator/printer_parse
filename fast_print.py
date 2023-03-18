import win32print
import win32api
import time
import argparse

GHOSTSCRIPT_PATH = "C:\\Program Files\\WinFast\\gs\\bin\\gswin64c.exe"
GSPRINT_PATH = "C:\\Program Files\\WinFast\\Ghostgum\\gsview\\gsprint.exe"

right_printer = ''
#currentprinter = win32print.GetDefaultPrinter()
rp = ' '

def createParser():
    parser = argparse.ArgumentParser()
    parser.add_argument('name', nargs='?')
    return parser

if __name__ == '__main__':
    parser = createParser()
    namespace = parser.parse_args()
    if namespace.name:
        print(namespace.name)
        str1 = str(namespace.name)
        #str1 = str1.encode('utf-8')
        find_printer = 0
        all_printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        for i in all_printers:
            if rp in i:
                print("OK")
                right_printer = i
                find_printer = 1
    
        if find_printer == 1:
            printer_name = right_printer#.decode() #'Xerox WorkCentre 73=7245 PCL 6'
            printer_srt = '" -printer "' + printer_name + '" '
            #file_name = str1
            file_str = '"' + str1 + '"'
            Str = '-ghostscript "' + GHOSTSCRIPT_PATH + printer_srt + file_str + " -color"
            print(Str)

            win32api.ShellExecute(0, 'open', GSPRINT_PATH, Str, '.', 0)