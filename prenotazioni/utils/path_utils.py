import sys
from pathlib import Path, PurePath
from shutil import copyfile


def get_folder_path():
    '''
    Get the path of current folder
    :return: the path of the folder
    '''

   # if is an executable get the path of current folder
    if getattr(sys, 'frozen', False):
        folder_path = Path(sys.executable).parent.resolve()
        return folder_path

    # else get the path of three top-level folder
    else:
        folder_path = Path(__file__).parent.parent.parent.resolve()
        return folder_path


def create_copy(folder_path, name_file):
    '''
    Create a copy of the file from complete_path to backup_file_path for security purpose
    :param folder_path: path of the folder
    :param name_file: name of excel file
    '''
    complete_path = str(PurePath(folder_path).joinpath(name_file))
    backup_file_path = str(PurePath.joinpath(folder_path, 'BACKUP ' + name_file))
    copyfile(complete_path, backup_file_path)

