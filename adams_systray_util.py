from io import StringIO
import os
import subprocess
import sys
from pathlib import Path
from typing import Callable, List, Tuple
import pandas as pd

import psutil
from PyQt5 import QtGui, QtWidgets, QtCore

ADAMS_ICON = Path('icons/adams.ico')

SOLVER_IMAGE = 'solver.exe'
AVIEW_IMAGE = 'aview.exe'

STARTUPINFO = subprocess.STARTUPINFO()
STARTUPINFO.dwFlags |= subprocess.STARTF_USESHOWWINDOW

class Menu(QtWidgets.QMenu):
    
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.aboutToShow.connect(self.populate)

    def get_proc_table(self):
        raise NotImplementedError

    def populate(self):
        
        # Set the wait cursor
        proc_table = self.get_proc_table()
        
        self.clear()

        for proc, pid, name, working_dir in proc_table:
            menu = QtWidgets.QMenu(f'{name} [{pid}]', self)

            open_dir = menu.addAction('Go to')
            open_dir.triggered.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(working_dir))))

            kill = menu.addAction('Kill')
            kill.triggered.connect(lambda: proc.terminate())
            
            self.addMenu(menu)

        # Add a disabled action if there are no processes
        if len(proc_table) == 0:
            noneAction = self.addAction('None')
            noneAction.setEnabled(False)

class SolverMenu(Menu):
    def get_proc_table(self):
        return get_solver_table()

class AviewMenu(Menu):
    def get_proc_table(self):
        return get_aview_table()

class SystemTrayIcon(QtWidgets.QSystemTrayIcon):

    def __init__(self, icon, parent:QtWidgets.QWidget=None):
        QtWidgets.QSystemTrayIcon.__init__(self, icon, parent)
        menu = QtWidgets.QMenu(parent)

        exitAction = menu.addAction('Exit')
        exitAction.triggered.connect(parent.close)

        menu.addSeparator()

        killAllSolversAction = menu.addAction('Kill all solver.exe')
        killAllSolversAction.triggered.connect(kill_all_solver)

        killAllAviewAction = menu.addAction('Kill all aview.exe')
        killAllAviewAction.triggered.connect(kill_all_aview)
        
        menu.addSeparator()

        menu.addMenu(SolverMenu('solver processes', parent=parent))
        menu.addMenu(AviewMenu('aview processes', parent=parent))

        self.setContextMenu(menu)

def get_solver_table()->List[Tuple[psutil.Process, int, str, Path]]:
    
    with subprocess.Popen(f'tasklist /FO csv /FI "imagename eq {SOLVER_IMAGE}"', 
                          startupinfo=STARTUPINFO, 
                          stdout=subprocess.PIPE, text=True) as proc:
        out, err = proc.communicate()

    if err is not None:
        raise RuntimeError(err)
    
    table = []
    for _, row in pd.read_csv(StringIO(out)).iterrows():
        proc = psutil.Process(int(row['PID']))
        pid = int(row['PID'])
        res_file = next(Path(p.path) for p in proc.open_files() if Path(p.path).suffix == '.res')
        working_dir = res_file.parent.absolute()
        ans_name = res_file.stem

        table.append([proc, pid, ans_name, working_dir])
    
    return sorted(table, key=lambda x: x[2], reverse=True)

def get_aview_table()->List[Tuple[psutil.Process, int, str, Path]]:

    with subprocess.Popen(f'tasklist /FO csv /FI "imagename eq {AVIEW_IMAGE}"', 
                          startupinfo=STARTUPINFO, 
                          stdout=subprocess.PIPE, text=True) as proc:
        out, err = proc.communicate()

    if err is not None:
        raise RuntimeError(err)
    
    table = []
    for _, row in pd.read_csv(StringIO(out)).iterrows():
        proc = psutil.Process(int(row['PID']))
        pid = int(row['PID'])
        working_dir = proc.cwd()

        table.append([proc, pid, Path(working_dir).as_posix(), working_dir])
    
    return sorted(table, key=lambda x: x[2], reverse=True)

def kill_all_solver():
    os.system('taskkill /f /fi "imagename eq solver.exe"')

def kill_all_aview():
    os.system('taskkill /f /fi "imagename eq aview.exe"')

def main():
    app = QtWidgets.QApplication(sys.argv)

    w = QtWidgets.QWidget()
    trayIcon = SystemTrayIcon(QtGui.QIcon(str(ADAMS_ICON)), w)


    trayIcon.show()
    sys.exit(app.exec_())



if __name__ == '__main__':
    main()
