"""Export through a separately owned, hidden Word process without desktop input."""
import argparse
from pathlib import Path


def export(docx, pdf):
    if Path(docx).resolve() == Path(pdf).resolve() or Path(pdf).exists():
        raise ValueError('PDF export requires a new path distinct from the input')
    import pythoncom
    import win32com.client
    import win32gui
    import win32process
    app = document = None
    owned = False
    pythoncom.CoInitialize()
    try:
        existing = set(win32process.EnumProcesses())
        app = win32com.client.DispatchEx('Word.Application')
        windows = []

        def collect(hwnd, _):
            if win32gui.GetClassName(hwnd) == 'OpusApp':
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid and pid not in existing:
                    windows.append(hwnd)

        win32gui.EnumWindows(collect, None)
        if len(windows) != 1 or win32gui.IsWindowVisible(windows[0]) or app.Visible:
            raise RuntimeError('Word did not create one separate hidden window; refusing to change or close a shared application')
        owned = True
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        document = app.Documents.Open(str(Path(docx).resolve()), ReadOnly=True, Visible=False,
                                      AddToRecentFiles=False, ConfirmConversions=False)
        document.Repaginate()
        document.ExportAsFixedFormat(str(Path(pdf).resolve()), 17)
    finally:
        try:
            if document is not None:
                document.Close(False)
        finally:
            try:
                if owned and app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('pdf')
    args = parser.parse_args()
    export(args.docx, args.pdf)
