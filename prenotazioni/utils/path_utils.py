import sys
from pathlib import Path, PurePath
from shutil import copyfile


def get_folder_path():

    if getattr(sys, 'frozen', False):
        folder_path = Path(sys.executable).parent.resolve()
        # print(f'Path dell\'eseguibile: {folder_path}')
        return folder_path

    else:
        folder_path = Path(__file__).parent.parent.parent.resolve()
        # print(f'Path dello script: {folder_path}')
        return folder_path


def create_copy(folder_path, name_file):
    complete_path = str(PurePath(folder_path).joinpath(name_file))
    backup_file_path = str(PurePath.joinpath(folder_path, 'BACKUP ' + name_file))
    copyfile(complete_path, backup_file_path)

