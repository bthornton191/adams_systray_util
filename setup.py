import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {
    'packages': ['pandas', 'psutil', 'PyQt5'],
    'excludes': [],
    'include_files': ['icons'],
}

bdist_msi_options = {
    'upgrade_code': '{90df58c3-ab91-4c02-84c6-416e658b6aeb}',
    'data': {
        'Icon': [('IconId', 'icons/adams.ico')],
    },
    'target_name': 'adams_systray_util',
    'install_icon': 'icons/adams.ico',
    'initial_target_dir': r'[ProgramFilesFolder]\AdamsSysTrayUtil',
}

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('adams_systray_util.py',
               base=base,
               icon='icons/adams.ico')
]

setup(name='adams_systray_util',
      version='1.0',
      description='A system tray utility for the Hexagon Adams user.',
      options={'build_exe': build_options,
               'bdist_msi': bdist_msi_options},
      executables=executables)
