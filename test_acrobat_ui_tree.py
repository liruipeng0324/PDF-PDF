from pywinauto import Application


app = Application(backend="uia").connect(title_re=".*Adobe Acrobat.*")
win = app.window(title_re=".*Adobe Acrobat.*")
win.set_focus()
win.print_control_identifiers()
