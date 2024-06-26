import shutil
import os
import datetime
import subprocess
import time
import stat

default_backup_base_folder = r'C:\path\to\your\backup\folder'
default_solution_base_folder = r'C:\path\to\your\solution\folder'

def close_visual_studio():
    try:
        subprocess.run(["taskkill", "/F", "/IM", "devenv.exe"], check=True)
        print("Visual Studio closed successfully.")
    except subprocess.CalledProcessError as e:
        print("Error closing Visual Studio:", e)

def is_visual_studio_running():
    try:
        result = subprocess.run(["tasklist", "/FI", "IMAGENAME eq devenv.exe"], capture_output=True, text=True)
        return "devenv.exe" in result.stdout
    except subprocess.CalledProcessError as e:
        print("Error checking Visual Studio process:", e)
        return False

def get_open_solution_path():
    try:
        result = subprocess.run(["powershell", "-File", "get_open_solution.ps1"], capture_output=True, text=True)
        solution_path = result.stdout.strip()
        if not solution_path or "No open solution found." in solution_path or "No Visual Studio instance found." in solution_path:
            raise Exception(solution_path)
        return solution_path
    except Exception as e:
        print(f"Error retrieving open solution path: {e}")
        return None

def get_path(prompt, default_path):
    new_path = input(f"Enter the {prompt} or press Enter to use default [{default_path}]: ").strip()
    return new_path if new_path else default_path

def list_folders(path):
    folders = [f for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))]
    if not folders:
        print(f"No folders found in '{path}'.")
        return None
    print("Available folders:")
    for idx, folder in enumerate(folders):
        print(f"{idx + 1}. {folder}")
    return folders

def choose_folder(folders, prompt):
    choice = input(f"Enter the number of the {prompt} you want to select: ").strip()
    if not choice.isdigit() or int(choice) < 1 or int(choice) > len(folders):
        print("Invalid choice.")
        return None
    return folders[int(choice) - 1]

def clear_directory(directory):
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path, onerror=handle_remove_readonly)
        except Exception as e:
            print(f"Failed to delete {item_path}. Reason: {e}")

def handle_remove_readonly(func, path, exc_info):
    # Remove read-only attribute if present and retry
    os.chmod(path, stat.S_IWRITE)
    func(path)

def backup_solution(solution_path, solution_name, backup_base_folder):
    print(f"Backing up solution '{solution_name}' from '{solution_path}'...")

    now = datetime.datetime.now()
    date_str = now.strftime('%Y-%m-%d')

    backup_folder = os.path.join(backup_base_folder, date_str)
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
        print(f"Created backup folder: {backup_folder}")
    else:
        print(f"Backup folder already exists: {backup_folder}")

    copy_number = 1
    while True:
        copy_folder_name = f'{solution_name}_copy_{copy_number}'
        copy_folder_path = os.path.join(backup_folder, copy_folder_name)
        if not os.path.exists(copy_folder_path):
            break
        copy_number += 1

    if is_visual_studio_running():
        print("Closing Visual Studio...")
        close_visual_studio()
        while is_visual_studio_running():
            print("Waiting for Visual Studio to close completely...")
            time.sleep(5)
        print("Visual Studio is fully closed. Proceeding with backup...")

    shutil.copytree(solution_path, os.path.join(copy_folder_path, solution_name))

    description_file_name = f'{date_str}_copy_{copy_number}_{solution_name}_readme.txt'
    description_file_path = os.path.join(copy_folder_path, description_file_name)
    with open(description_file_path, 'w') as description_file:
        description_file.write(f'Original Path: {solution_path}\n')
        description_file.write(f'Backup Path: {copy_folder_path}\n')
        description_file.write(f'Solution Name: {solution_name}\nDate: {date_str}\nTime: {now.strftime("%H:%M:%S")}\n\n')
        description = input(f"Enter the description of changes and next steps for '{solution_name}': ").strip()
        description_file.write(description)

    print(f"Backup of solution '{solution_name}' completed successfully.")

def interactive_backup_solution():
    print("Retrieving the path of the currently open solution in Visual Studio...")
    solution_path = get_open_solution_path()
    if solution_path is None or solution_path == '':
        print("No open solution found.")
        base_path = get_path("solution base folder path", default_solution_base_folder)
        folders = list_folders(base_path)
        if folders is None:
            print("Exiting...")
            return
        solution_name = choose_folder(folders, "solution folder")
        if solution_name is None:
            print("Exiting...")
            return
        solution_path = os.path.join(base_path, solution_name)
    else:
        solution_dir = os.path.dirname(solution_path)
        solution_name = os.path.basename(solution_path).replace(".sln", "")

    print(f"Solution folder path retrieved: '{solution_path}'")
    print(f"Solution name: {solution_name}")

    backup_base_folder = get_path("backup folder path", default_backup_base_folder)
    backup_solution(solution_path, solution_name, backup_base_folder)

def backup_all_solutions():
    base_path = get_path("solution base folder path", default_solution_base_folder)
    folders = list_folders(base_path)
    if folders is None:
        print("Exiting...")
        return
    backup_base_folder = get_path("backup folder path", default_backup_base_folder)
    for folder in folders:
        solution_path = os.path.join(base_path, folder)
        backup_solution(solution_path, folder, backup_base_folder)

def restore_solution():
    def get_backup_folder_path():
        return default_backup_base_folder

    def list_backups(backup_folder_path):
        backups = [d for d in os.listdir(backup_folder_path) if os.path.isdir(os.path.join(backup_folder_path, d))]
        if not backups:
            print(f"No backups found in '{backup_folder_path}'.")
            exit(1)
        return backups

    def choose_backup(backups, backup_folder_path):
        backups.sort(key=lambda x: os.path.getmtime(os.path.join(backup_folder_path, x)), reverse=True)
        print("Available backups:")
        for idx, backup in enumerate(backups):
            print(f"{idx + 1}. {backup}")
        choice = input(f"Enter the number of the backup you want to restore (or press Enter to use the last modified backup '{backups[0]}'): ").strip()
        if choice == '':
            return backups[0]
        if not choice.isdigit() or int(choice) < 1 or int(choice) > len(backups):
            print("Invalid choice.")
            exit(1)
        return backups[int(choice) - 1]

    def list_subfolders(backup_path):
        subfolders = [d for d in os.listdir(backup_path) if os.path.isdir(os.path.join(backup_path, d))]
        if not subfolders:
            print(f"No subfolders found in '{backup_path}'.")
            exit(1)
        return subfolders

    def choose_subfolder(subfolders, backup_path):
        subfolders.sort(key=lambda x: os.path.getmtime(os.path.join(backup_path, x)), reverse=True)
        print("Available subfolders:")
        for idx, subfolder in enumerate(subfolders):
            print(f"{idx + 1}. {subfolder}")
        choice = input(f"Enter the number of the subfolder you want to restore (or press Enter to use the last modified subfolder '{subfolders[0]}'): ").strip()
        if choice == '':
            return subfolders[0]
        if not choice.isdigit() or int(choice) < 1 or int(choice) > len(subfolders):
            print("Invalid choice.")
            exit(1)
        return subfolders[int(choice) - 1]

    def get_original_path_from_description(description_file_path):
        try:
            with open(description_file_path, 'r') as file:
                lines = file.readlines()
                for line in lines:
                    if line.startswith("Original Path:"):
                        return line.split(":", 1)[1].strip()
        except Exception as e:
            print(f"Error reading description file: {e}")
        return None

    def restore_backup(backup_path, subfolder):
        backup_subfolder_path = os.path.join(backup_path, subfolder)
        description_file_name = [f for f in os.listdir(backup_subfolder_path) if f.endswith('_readme.txt')][0]
        description_file_path = os.path.join(backup_subfolder_path, description_file_name)
        solution_folder_path = get_original_path_from_description(description_file_path)
        
        if not solution_folder_path:
            print(f"Could not determine the original solution path from the description file '{description_file_path}'.")
            return

        if is_visual_studio_running():
            close_vs = input("Visual Studio is currently running. Would you like to close it to proceed with the restore? (y/n): ").strip().lower()
            if close_vs == 'y':
                close_visual_studio()
                while is_visual_studio_running():
                    print("Waiting for Visual Studio to close completely...")
                    time.sleep(5)
            else:
                print("Cannot restore while Visual Studio is open. Exiting...")
                return
        
        print(f"Restoring backup from '{backup_subfolder_path}' to '{solution_folder_path}'...")
        try:
            # Clear the original solution folder
            if os.path.exists(solution_folder_path):
                clear_directory(solution_folder_path)
            else:
                os.makedirs(solution_folder_path)
            # Copy only the contents of the backup folder, excluding the description file
            solution_name = subfolder.split('_copy_')[0]
            backup_solution_path = os.path.join(backup_subfolder_path, solution_name)
            for item in os.listdir(backup_solution_path):
                s = os.path.join(backup_solution_path, item)
                d = os.path.join(solution_folder_path, item)
                if item == description_file_name:
                    continue
                if os.path.isdir(s):
                    shutil.copytree(s, d, dirs_exist_ok=True)
                else:
                    shutil.copy2(s, d)
            print("Restore completed successfully.")
        except Exception as e:
            print(f"Error during restore: {e}")

    backup_folder_path = get_backup_folder_path()
    backups = list_backups(backup_folder_path)
    backup = choose_backup(backups, backup_folder_path)
    subfolders = list_subfolders(os.path.join(backup_folder_path, backup))
    subfolder = choose_subfolder(subfolders, os.path.join(backup_folder_path, backup))
    restore_backup(os.path.join(backup_folder_path, backup), subfolder)

def main():
    print("Choose an option:")
    print("1. Backup Solution")
    print("2. Restore Solution")
    print("3. Backup All Solutions")
    choice = input("Enter the number of your choice (or press Enter for Backup): ").strip() or '1'
    if choice == '1':
        interactive_backup_solution()
    elif choice == '2':
        restore_solution()
    elif choice == '3':
        backup_all_solutions()
    else:
        print("Invalid choice. Exiting...")

if __name__ == "__main__":
    main()
