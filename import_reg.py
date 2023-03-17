import winreg as reg
import os
import sys

cwd = os.getcwd()

python_exe = sys.executable

#HKEY_CLASSES_ROOT\SystemFileAssociations\.png
key_path = r"SystemFileAssociations\\.pdf\\shell\\FastPrint"
#key_path = r".pdf\\shell\\FastPrint"

key = reg.CreateKeyEx(reg.HKEY_CLASSES_ROOT, key_path)
#key2 = reg.CreateKeyEx(reg.HKEY_CLASSES_ROOT, key_path)

reg.SetValueEx(key,'',0,reg.REG_SZ, '&Fast print')
reg.SetValueEx(key,'icon',0,reg.REG_SZ, f'{cwd}\\fast_print.exe')

key1 = reg.CreateKeyEx(key, r"command")
reg.SetValue(key1, '', reg.REG_SZ, f'{cwd}\\fast_print.exe ' + '"%1"')