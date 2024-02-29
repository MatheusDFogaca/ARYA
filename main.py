
"pyinstaller --additional-hooks-dir . --noconfirm  --onefile main.py"

from gui import MainGUI

if __name__ == '__main__':
    app = MainGUI()
    app.resizable(True, True)
    app.mainloop()