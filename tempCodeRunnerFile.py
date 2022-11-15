]
        win32print.SetDefaultPrinter(impressora[2])
        win32api.ShellExecute(0, "print", diretorio, None, diretorio[:-17], 0)