import win32print
import win32api

GHOSTSCRIPT_PATH = "./GHOSTSCRIPT\\bin\\gswin32.exe"
GSPRINT_PATH = "./GSPRINT\\gsprint.exe"

# YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT
currentprinter = win32print.GetDefaultPrinter()

win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH+'"-color -printer "'+currentprinter+'" "'+ 'Draw 2 Result.pdf' + '', '.', 0)
