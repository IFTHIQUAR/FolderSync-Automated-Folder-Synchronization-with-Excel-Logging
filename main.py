import os
import pandas as pd
from datetime import datetime
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import openpyxl


# Function to check if a file is accessible
def is_file_accessible(file_path, mode='r'):
    try:
        with open(file_path, mode):
            pass
    except IOError:
        return False
    return True


# Check and create necessary folders
def check_and_create_folders(client_folder, dev_folder):
    os.makedirs(client_folder, exist_ok=True)
    os.makedirs(dev_folder, exist_ok=True)
    print(f"Ensured that folders exist:\nClient: {client_folder}\nDev: {dev_folder}")


# Check and create the initial Excel file with required sheets if they do not exist
def create_initial_excel(file_path='D:\\input.xlsx'):
    if not os.path.exists(file_path):
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Create Input sheet with sample folder paths
            input_df = pd.DataFrame({
                'Client Folder': ['D:\\ClientFolder'],
                'Dev Folder': ['D:\\DevFolder']
            })
            input_df.to_excel(writer, sheet_name='Input', index=False)

            # Create Ignored Files sheet with sample data
            ignored_files_df = pd.DataFrame({
                'Ignored Files': ['ignore.txt']
            })
            ignored_files_df.to_excel(writer, sheet_name='Ignored Files', index=False)

            # Create Log sheet for recording changes
            log_df = pd.DataFrame(columns=['Timestamp', 'Event'])
            log_df.to_excel(writer, sheet_name='Log', index=False)

        print(f"Initial Excel file created at {file_path}")


# Ensure that the Excel file has the required sheets
def ensure_excel_sheets(file_path='D:\\input.xlsx'):
    if not is_file_accessible(file_path, 'a'):
        raise PermissionError(f"File {file_path} is not accessible. Please close it if it is open in another program.")

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        workbook = writer.book
        if 'Input' not in workbook.sheetnames:
            input_df = pd.DataFrame({
                'Client Folder': ['D:\\ClientFolder'],
                'Dev Folder': ['D:\\DevFolder']
            })
            input_df.to_excel(writer, sheet_name='Input', index=False)
            print(f"Created 'Input' sheet in {file_path}")

        if 'Ignored Files' not in workbook.sheetnames:
            ignored_files_df = pd.DataFrame({
                'Ignored Files': ['ignore.txt']
            })
            ignored_files_df.to_excel(writer, sheet_name='Ignored Files', index=False)
            print(f"Created 'Ignored Files' sheet in {file_path}")

        if 'Log' not in workbook.sheetnames:
            log_df = pd.DataFrame(columns=['Timestamp', 'Event'])
            log_df.to_excel(writer, sheet_name='Log', index=False)
            print(f"Created 'Log' sheet in {file_path}")


# Read folder paths from the Excel file
def read_folder_paths(file_path='D:\\input.xlsx'):
    df = pd.read_excel(file_path, sheet_name='Input')
    return df['Client Folder'][0], df['Dev Folder'][0]


# Read the list of ignored files from the Excel file
def read_ignored_files(file_path='D:\\input.xlsx'):
    try:
        df = pd.read_excel(file_path, sheet_name='Ignored Files')
        return df['Ignored Files'].tolist()
    except Exception as e:
        print(f"Error reading ignored files: {e}")
        return []


# Create a snapshot of files in a folder
def create_snapshot(folder_path, ignored_files):
    snapshot = []
    for file_name in os.listdir(folder_path):
        if file_name in ignored_files:
            continue
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            modified_time = os.path.getmtime(file_path)
            snapshot.append((file_name, datetime.fromtimestamp(modified_time)))
    return snapshot


# Compare the snapshots of the two folders
def compare_snapshots(client_snapshot, dev_snapshot):
    client_df = pd.DataFrame(client_snapshot, columns=['File Name', 'Client Modified'])
    dev_df = pd.DataFrame(dev_snapshot, columns=['File Name', 'Dev Modified'])

    combined_df = pd.merge(client_df, dev_df, on='File Name', how='outer', indicator=True)
    combined_df['Status'] = combined_df['_merge'].map({
        'both': 'Same File',
        'left_only': 'Modified in Client Only',
        'right_only': 'Modified in Dev Only'
    })

    combined_df.sort_values(by=['Client Modified', 'Dev Modified'], ascending=False, inplace=True)
    return combined_df


# Synchronize files based on the comparison
def sync_files(combined_df, client_folder, dev_folder):
    for _, row in combined_df.iterrows():
        file_name = row['File Name']
        client_file_path = os.path.join(client_folder, file_name)
        dev_file_path = os.path.join(dev_folder, file_name)

        if row['Status'] == 'Modified in Client Only':
            if os.path.exists(client_file_path):
                shutil.copy(client_file_path, dev_folder)
                print(f"Copied {file_name} from Client to Dev.")
            else:
                if os.path.exists(dev_file_path):
                    os.remove(dev_file_path)
                    print(f"Deleted {file_name} from Dev because it was deleted in Client.")
        elif row['Status'] == 'Same File':
            if os.path.exists(client_file_path) and not os.path.exists(dev_file_path):
                shutil.copy(client_file_path, dev_folder)
                print(f"Copied {file_name} from Client to Dev because it was missing in Dev.")


# Update the Excel file with the snapshot data and highlight differences
def update_snapshot_excel(combined_df, excel_path):
    if not is_file_accessible(excel_path, 'a'):
        raise PermissionError(f"File {excel_path} is not accessible. Please close it if it is open in another program.")

    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        combined_df.to_excel(writer, sheet_name='Last Snapshot', index=False)

        # Add conditional formatting for highlighting differences
        workbook = writer.book
        worksheet = writer.sheets['Last Snapshot']
        for index, row in combined_df.iterrows():
            cell = worksheet.cell(row=index + 2, column=1)
            if row['Status'] == 'Same File':
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00",
                                                        fill_type="solid")  # Green
            elif row['Status'] == 'Modified in Client Only':
                cell.fill = openpyxl.styles.PatternFill(start_color="FFCC00", end_color="FFCC00",
                                                        fill_type="solid")  # Yellow
            elif row['Status'] == 'Modified in Dev Only':
                cell.fill = openpyxl.styles.PatternFill(start_color="FF0000", end_color="FF0000",
                                                        fill_type="solid")  # Red

        print(f"Snapshot updated in {excel_path}")


# Log changes to the Excel file
def log_changes(file_path, event):
    log_df = pd.read_excel(file_path, sheet_name='Log')
    new_log = pd.DataFrame([{
        'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Event': event
    }])
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        log_df = pd.concat([log_df, new_log], ignore_index=True)
        log_df.to_excel(writer, sheet_name='Log', index=False)
    print(f"Logged change: {event}")


# Run the synchronization process and update the snapshot in the Excel file
def run_sync(client_folder, dev_folder, ignored_files, excel_path):
    client_snapshot = create_snapshot(client_folder, ignored_files)
    dev_snapshot = create_snapshot(dev_folder, ignored_files)
    combined_df = compare_snapshots(client_snapshot, dev_snapshot)
    sync_files(combined_df, client_folder, dev_folder)
    update_snapshot_excel(combined_df, excel_path)
    return combined_df


# Event handler for detecting changes in the folders
class SyncEventHandler(FileSystemEventHandler):
    def __init__(self, client_folder, dev_folder, ignored_files, log_file):
        super().__init__()
        self.client_folder = client_folder
        self.dev_folder = dev_folder
        self.ignored_files = ignored_files
        self.log_file = log_file

    def on_modified(self, event):
        if event.is_directory:
            return
        event_message = f'File modified: {event.src_path}'
        print(event_message)
        log_changes(self.log_file, event_message)
        run_sync(self.client_folder, self.dev_folder, self.ignored_files, self.log_file)

    def on_created(self, event):
        if event.is_directory:
            return
        event_message = f'File created: {event.src_path}'
        print(event_message)
        log_changes(self.log_file, event_message)
        run_sync(self.client_folder, self.dev_folder, self.ignored_files, self.log_file)

    def on_deleted(self, event):
        if event.is_directory:
            return
        event_message = f'File deleted: {event.src_path}'
        print(event_message)
        log_changes(self.log_file, event_message)
        run_sync(self.client_folder, self.dev_folder, self.ignored_files, self.log_file)


if __name__ == "__main__":
    excel_path = 'D:\\input.xlsx'

    # Create the initial Excel file if it does not exist
    create_initial_excel(excel_path)

    # Ensure the Excel file has the required sheets
    ensure_excel_sheets(excel_path)

    # Read folder paths and ignored files
    client_folder, dev_folder = read_folder_paths(excel_path)
    ignored_files = read_ignored_files(excel_path)

    # Check and create necessary folders
    check_and_create_folders(client_folder, dev_folder)

    # Initial sync
    run_sync(client_folder, dev_folder, ignored_files, excel_path)

    # Set up the observer for real-time monitoring
    event_handler = SyncEventHandler(client_folder, dev_folder, ignored_files, excel_path)
    observer = Observer()
    observer.schedule(event_handler, client_folder, recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
