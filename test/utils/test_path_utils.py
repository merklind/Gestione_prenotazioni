from pathlib import Path

import pytest
from pytest_mock import MockFixture

from utils.path_utils import build_file_path, get_root_path, create_copy


def test_get_folder_path_executable(mocker: MockFixture):
    mock_is_frozen = mocker.patch('utils.path_utils._is_frozen', return_value=True)
    mock_sys = mocker.patch('utils.path_utils.sys')
    mock_resolve = mocker.patch('utils.path_utils.Path.resolve', return_value=Path("absolute/path/of/the/executable"))

    mock_sys.executable = "path/of/the/executable"

    result = get_root_path()

    mock_resolve.assert_called()
    mock_is_frozen.assert_called()
    assert result == Path("absolute/path/of/the/executable")


def test_get_folder_path_not_executable(mocker: MockFixture):
    mock_is_frozen = mocker.patch('utils.path_utils._is_frozen', return_value=True)
    mock_sys = mocker.patch('utils.path_utils.sys')
    mock_sys.executable = "path/of/the/main/entrypoint"
    mock_resolve = mocker.patch('utils.path_utils.Path.resolve',
                                return_value=Path("absolute/path/of/the/main/entrypoint"))

    result = get_root_path()

    mock_is_frozen.assert_called()
    mock_resolve.assert_called()
    assert result == Path("absolute/path/of/the/main/entrypoint")


def test_build_file_path(mocker: MockFixture):
    mock_get_root_path = mocker.patch("utils.path_utils.get_root_path", return_value=Path("/mocked/folder/path"))

    # Arrange
    filename = "test_file.txt"
    expected_file_path = Path("/mocked/folder/path/test_file.txt")

    # Act
    result = build_file_path(filename)

    # Assert
    assert result == expected_file_path
    mock_get_root_path.assert_called_once()


def test_create_copy(mocker: MockFixture):
    mock_copyfile = mocker.patch('utils.path_utils.copyfile')

    file_path = Path("path/to/folder/filename.xlsx")

    create_copy(file_path)

    mock_copyfile.assert_called_with(file_path,
                                     Path("path/to/folder/BACKUP filename.xlsx"))


if __name__ == "__main__":
    pytest.main()
