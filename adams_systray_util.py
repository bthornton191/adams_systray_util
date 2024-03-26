import logging
import os
import platform
import subprocess
import sys
from io import StringIO
from logging.handlers import WatchedFileHandler
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import psutil
from PyQt5 import QtCore, QtGui, QtWidgets

ADAMS_ICON = Path('icons/adams.ico')

SOLVER_IMAGE = 'solver.exe'
AVIEW_IMAGE = 'aview.exe'

STARTUPINFO = subprocess.STARTUPINFO()
STARTUPINFO.dwFlags |= subprocess.STARTF_USESHOWWINDOW


def setup_logging():
    handler = WatchedFileHandler(f'{Path(__file__).stem}.log')
    formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s', '%Y-%m-%d %H:%M:%S')
    handler.setFormatter(formatter)
    root = logging.getLogger()
    root.setLevel('DEBUG')
    # Remove existing handlers for this file name, if any
    for old_handler in [h for h in root.handlers if (isinstance(h, WatchedFileHandler)
                                                     and h.baseFilename == handler.baseFilename)]:
        root.handlers.remove(old_handler)
    root.addHandler(handler)
    return logging.getLogger(__name__)


LOG = setup_logging()


class Menu(QtWidgets.QMenu):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.aboutToShow.connect(self.populate)

    def get_proc_table(self) -> List[Tuple[psutil.Process, int, str, Path]]:
        raise NotImplementedError

    def populate(self):

        # Set the wait cursor
        proc_table = self.get_proc_table()

        self.clear()

        for proc, pid, name, path in proc_table:
            menu = QtWidgets.QMenu(f'{name} [{pid}]', self)

            open_dir = menu.addAction('Go to')
            open_dir.triggered.connect(lambda: goto(Path(path)))

            kill = menu.addAction('Kill')
            kill.triggered.connect(lambda: terminate_process(proc))

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

    def __init__(self, icon, parent: QtWidgets.QWidget = None):

        LOG.debug('Initializing tray icon...')
        QtWidgets.QSystemTrayIcon.__init__(self, icon, parent)
        menu = QtWidgets.QMenu(parent)

        exitAction = menu.addAction('Exit')
        exitAction.triggered.connect(parent.close)

        menu.addSeparator()

        # ------------------------------------------------------------------------------------------
        # TODO: Implement Start with Windows feature
        # ------------------------------------------------------------------------------------------
        # startupAction = QtWidgets.QAction('Start with Windows', checkable=True)
        # startupAction.triggered.connect(run_at_startup)
        # menu.addAction(startupAction)
        # menu.addSeparator()
        # ------------------------------------------------------------------------------------------

        killAllSolversAction = menu.addAction('Kill all solver.exe')
        killAllSolversAction.triggered.connect(kill_all_solver)

        killAllAviewAction = menu.addAction('Kill all aview.exe')
        killAllAviewAction.triggered.connect(kill_all_aview)

        menu.addSeparator()

        menu.addMenu(SolverMenu('solver processes', parent=parent))
        menu.addMenu(AviewMenu('aview processes', parent=parent))

        self.setContextMenu(menu)


def terminate_process(proc: psutil.Process):
    try:
        proc.terminate()
    except psutil.NoSuchProcess:
        pass


def goto(path: Path):
    if not path.exists():
        return

    if path.is_dir():
        goto_dir(path)

    elif platform.system() != 'Windows':
        goto_dir(path.parent)

    else:
        with subprocess.Popen(' '.join(['explorer', f'/select,"{path.absolute().resolve()}"']),
                              stdout=subprocess.PIPE,
                              stderr=subprocess.STDOUT,
                              text=True) as proc:
            for line in proc.stdout:
                LOG.debug(line.rstrip())


def goto_dir(dir_: Path):
    QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(dir_)))


def get_solver_table() -> List[Tuple[psutil.Process, int, str, Path]]:

    with subprocess.Popen(f'tasklist /FO csv /FI "imagename eq {SOLVER_IMAGE}"',
                          startupinfo=STARTUPINFO,
                          stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE,
                          text=True) as proc:
        out, err = proc.communicate()

    if err:
        LOG.error(err)
        return []

    table = []
    for _, row in pd.read_csv(StringIO(out)).iterrows():
        proc = psutil.Process(int(row['PID']))
        pid = int(row['PID'])
        res_file = next(Path(p.path) for p in proc.open_files() if Path(p.path).suffix == '.res')
        msg_file = res_file.with_suffix('.msg')
        ans_name = res_file.stem

        table.append([proc, pid, ans_name, msg_file])

    return sorted(table, key=lambda x: x[2], reverse=True)


def get_aview_table() -> List[Tuple[psutil.Process, int, str, Path]]:

    with subprocess.Popen(f'tasklist /FO csv /FI "imagename eq {AVIEW_IMAGE}"',
                          startupinfo=STARTUPINFO,
                          stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE,
                          text=True) as proc:
        out, err = proc.communicate()

    if err:
        LOG.error(err)
        return []

    table = []
    for _, row in pd.read_csv(StringIO(out)).iterrows():
        proc = psutil.Process(int(row['PID']))
        pid = int(row['PID'])
        working_dir = proc.cwd()

        table.append([proc, pid, Path(working_dir).as_posix(), working_dir])

    return sorted(table, key=lambda x: x[2], reverse=True)


def kill_all_solver():
    subprocess.Popen(f'taskkill /f /fi "imagename eq {SOLVER_IMAGE}"', startupinfo=STARTUPINFO)


def kill_all_aview():
    subprocess.Popen(f'taskkill /f /fi "imagename eq {AVIEW_IMAGE}"', startupinfo=STARTUPINFO)


def run_at_startup(enabled: bool):
    startup_path = Path(os.getenv('APPDATA')) / 'Microsoft/Windows/Start Menu/Programs/Startup'
    shortcut_path = startup_path / 'adams_systray_util.lnk'

    if enabled and shortcut_path.exists():
        return

    target_path = Path(sys.executable).absolute()

    ps_cmds = ['$WshShell = New-Object -comObject WScript.Shell',
               f'$Shortcut = $WshShell.CreateShortcut("{shortcut_path}")']

    if 'python' in target_path.stem:
        target_path = target_path.parent / 'pythonw.exe'

        ps_cmds += [f'$Shortcut.Arguments = "{Path(__file__).absolute().resolve()}"']

    ps_cmds += [f'$Shortcut.TargetPath = "{target_path}"',
                '$Shortcut.Save()']

    with subprocess.Popen(['powershell.exe',
                           '; '.join(ps_cmds)],
                          startupinfo=STARTUPINFO,
                          stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE,
                          text=True) as proc:
        out, err = proc.communicate()

    if err:
        raise RuntimeError(err)

    LOG.debug(out)
    LOG.info('Created startup shortcut')


def excepthook(exc_type, exc_value, exc_traceback):
    LOG.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


def main():

    LOG.info('Starting...')
    app = QtWidgets.QApplication(sys.argv)

    w = QtWidgets.QWidget()
    trayIcon = SystemTrayIcon(QtGui.QIcon(str(ADAMS_ICON)), w)

    trayIcon.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    sys.excepthook = excepthook
    main()
