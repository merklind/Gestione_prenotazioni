import sys
from pathlib import Path
from shutil import copyfile


def _is_frozen():
    return getattr(sys, 'frozen', False)


def get_root_path() -> Path:
    '''
    Get the path of current folder
    :return: the path of the folder
    '''

    # if is an executable get the path of current folder
    if _is_frozen():
        folder_path = Path(sys.executable).parent.resolve()

    # else get the path of three top-level folder
    else:
        folder_path = Path(__file__).parent.parent.parent.resolve()

    return folder_path


def build_file_path(filename: str) -> Path:
    folder_path = get_root_path()
    file_path = folder_path.joinpath(filename)

    return file_path


def create_copy(file_path: Path) -> None:
    '''
    Create a copy of the file from complete_path to backup_file_path for security purpose
    :param folder_path: path of the folder
    :param name_file: name of excel file
    '''
    folder_path = file_path.parent
    file_name = file_path.name

    backup_file_path = folder_path.joinpath(f'BACKUP {file_name}')
    copyfile(file_path, backup_file_path)
