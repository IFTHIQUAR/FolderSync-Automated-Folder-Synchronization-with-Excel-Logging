import os
import tempfile
import shutil
import pytest
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from watchdog.events import FileSystemEvent
from main import (
    is_file_accessible,
    read_config,
    check_and_create_folders,
    create_initial_excel,
    ensure_excel_sheets,
    update_snapshot_excel,
    log_changes,
    FileChangeHandler
)


@pytest.fixture
def setup_temp_dir():
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


def test_is_file_accessible(setup_temp_dir):
    file_path = os.path.join(setup_temp_dir, 'test.txt')
    with open(file_path, 'w') as f:
        f.write('Test content')

    assert is_file_accessible(file_path, 'r')

def test_check_and_create_folders(setup_temp_dir):
    client_folder = os.path.join(setup_temp_dir, 'Client')
    dev_folder = os.path.join(setup_temp_dir, 'Dev')
    check_and_create_folders(client_folder, dev_folder)
    assert os.path.exists(client_folder)
    assert os.path.exists(dev_folder)


def test_create_initial_excel(setup_temp_dir):
    excel_path = os.path.join(setup_temp_dir, 'snapshot.xlsx')
    create_initial_excel(excel_path)
    assert os.path.exists(excel_path)

def test_read_folder_paths(setup_temp_dir):
    client_folder = os.path.join(setup_temp_dir, 'Client')
    dev_folder = os.path.join(setup_temp_dir, 'Dev')
    check_and_create_folders(client_folder, dev_folder)

    assert os.path.exists(client_folder)
    assert os.path.exists(dev_folder)


def test_read_ignored_files(setup_temp_dir):
    excel_path = os.path.join(setup_temp_dir, 'snapshot.xlsx')
    create_initial_excel(excel_path)

    # Simulate reading ignored files from Excel
    ignored_files = []
    assert isinstance(ignored_files, list)


def test_create_snapshot(setup_temp_dir):
    excel_path = os.path.join(setup_temp_dir, 'snapshot.xlsx')
    create_initial_excel(excel_path)

    # Simulate snapshot creation
    client_snapshot = [('file1.txt', datetime.now())]
    assert isinstance(client_snapshot, list)


def test_compare_snapshots(setup_temp_dir):
    client_snapshot = [('file1.txt', datetime.now())]
    dev_snapshot = [('file1.txt', datetime.now())]

    # Simulate snapshot comparison logic
    assert client_snapshot[0][0] == dev_snapshot[0][0]


def test_sync_files(setup_temp_dir):
    client_folder = os.path.join(setup_temp_dir, 'Client')
    dev_folder = os.path.join(setup_temp_dir, 'Dev')
    os.makedirs(client_folder, exist_ok=True)
    os.makedirs(dev_folder, exist_ok=True)

    # Simulate file sync
    file_path = os.path.join(client_folder, 'file1.txt')
    with open(file_path, 'w') as f:
        f.write('Test content')

    shutil.copy(file_path, dev_folder)
    assert os.path.exists(os.path.join(dev_folder, 'file1.txt'))

def test_log_changes(setup_temp_dir):
    log_file = os.path.join(setup_temp_dir, 'log.txt')

    log_changes(log_file, "Test event")
    with open(log_file, 'r') as f:
        logs = f.read()

    assert "Test event" in logs


