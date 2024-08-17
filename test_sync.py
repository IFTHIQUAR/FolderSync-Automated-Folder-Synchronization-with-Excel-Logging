import pytest
import os
import pandas as pd
import tempfile
from main import update_snapshot_excel, create_initial_excel
from unittest.mock import patch, mock_open
from main import (
    is_file_accessible,
    check_and_create_folders,
    create_initial_excel,
    read_folder_paths,
    read_ignored_files,
    create_snapshot,
    compare_snapshots,
    sync_files,
    update_snapshot_excel
)


# Test for is_file_accessible
def test_is_file_accessible():
    with patch("builtins.open", mock_open(read_data="data")) as mock_file:
        assert is_file_accessible("dummy_path") is True
    with patch("builtins.open", mock_open()) as mock_file:
        mock_file.side_effect = IOError()
        assert is_file_accessible("dummy_path") is False


# Test for check_and_create_folders
def test_check_and_create_folders(tmp_path):
    client_folder = tmp_path / "ClientFolder"
    dev_folder = tmp_path / "DevFolder"
    check_and_create_folders(client_folder, dev_folder)
    assert os.path.exists(client_folder)
    assert os.path.exists(dev_folder)


# Test for create_initial_excel
def test_create_initial_excel(tmp_path):
    excel_path = tmp_path / "input.xlsx"
    create_initial_excel(excel_path)
    assert os.path.exists(excel_path)


# Test for read_folder_paths
def test_read_folder_paths(tmp_path):
    excel_path = tmp_path / "input.xlsx"
    create_initial_excel(excel_path)
    client_folder, dev_folder = read_folder_paths(excel_path)
    assert client_folder == 'D:\\ClientFolder'
    assert dev_folder == 'D:\\DevFolder'


# Test for read_ignored_files
def test_read_ignored_files(tmp_path):
    excel_path = tmp_path / "input.xlsx"
    create_initial_excel(excel_path)
    ignored_files = read_ignored_files(excel_path)
    assert ignored_files == ['ignore.txt']


# Test for create_snapshot
def test_create_snapshot(tmp_path):
    ignored_files = []
    test_file = tmp_path / "test.txt"
    test_file.write_text("content")
    snapshot = create_snapshot(tmp_path, ignored_files)
    assert len(snapshot) == 1
    assert snapshot[0][0] == "test.txt"


# Test for compare_snapshots
def test_compare_snapshots():
    client_snapshot = [("file1.txt", "2023-08-17 10:00:00")]
    dev_snapshot = [("file1.txt", "2023-08-17 10:00:00")]
    combined_df = compare_snapshots(client_snapshot, dev_snapshot)
    assert combined_df.iloc[0]['Status'] == 'Same File'


# Test for sync_files (mocking shutil.copy)
def test_sync_files(tmp_path):
    combined_df = pd.DataFrame({
        'File Name': ['file1.txt'],
        'Status': ['Modified in Client Only']
    })
    client_folder = tmp_path / "ClientFolder"
    dev_folder = tmp_path / "DevFolder"
    client_folder.mkdir()
    dev_folder.mkdir()
    (client_folder / 'file1.txt').write_text("content")

    with patch("shutil.copy") as mock_copy:
        sync_files(combined_df, client_folder, dev_folder)
        mock_copy.assert_called_once()


# Test for update_snapshot_excel (mocking Excel writing)
def test_update_snapshot_excel():
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, 'temp_excel.xlsx')
        create_initial_excel(excel_path)

        combined_df = pd.DataFrame({
            'File Name': ['file1.txt'],
            'Status': ['Modified in Client Only']
        })

        # Call the function to test
        update_snapshot_excel(combined_df, excel_path)

        # Read the updated Excel file to verify changes
        updated_df = pd.read_excel(excel_path, sheet_name='Last Snapshot')

        # Perform your assertions
        assert updated_df.equals(combined_df)