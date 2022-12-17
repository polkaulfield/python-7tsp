#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, shutil, subprocess, ctypes, sys, psutil
        
# Declare variables and arrays

script_dir = os.path.dirname(sys.executable)
iconpacks_dir = os.path.join(script_dir, "iconpacks")
reshacker_path = os.path.join(script_dir, "ResourceHacker.exe")
unpatched_files_dir = os.path.join(script_dir, "unpatched")
patched_files_dir = os.path.join(script_dir, "patched")
system_mun_dir = os.path.join("C:\\", "Windows", "SystemResources")
tmp_dir = os.path.join(script_dir, "tmp")
resources_dir = os.path.join(tmp_dir, "Resources")
szip_path = os.path.join(script_dir, "7za.exe")

# Needed functions

def check_exe_dependencies():
    files = os.listdir(script_dir)
    if "7za.exe" not in files or "ResourceHacker.exe" not in files:
        input("Some executables missing, cannot continue...")
        sys.exit(1)

def create_needed_folders():
    if not os.path.exists(iconpacks_dir):
        os.makedirs(iconpacks_dir)
        input("Created iconpacks folder. Place your 7tsp packs there.")
        sys.exit(0)

def get_by_filetype(path, filetype):
    file_array = []
    for file in os.listdir(path):
        if file.endswith(filetype):
            file_array.append(file)
    return file_array

def select_file_menu(msg, file_array):
    print(msg)
    for file in file_array:
        print(str(file_array.index(file)) + ": " + file)    
    while True:
        try:
            selection = int(input("Type a number:"))
            if selection <= len(file_array):
                return file_array[selection]
            else:
                print("Value not in the list!")
        except ValueError:
            print("Input a number!")

def extract_7z(file, extract_path):
    subprocess.call((szip_path,"x",file,"-o" + extract_path))

def restore_unpatched():
    for unpatched_file in os.listdir(unpatched_files_dir):
        source_path = os.path.join(unpatched_files_dir, unpatched_file)
        target_path = os.path.join(system_mun_dir, unpatched_file)
        print("Changing Permissions of " + target_path)
        subprocess.call(("takeown", "/f",target_path))
        subprocess.call(("icacls",target_path ,"/grant","Administrators:rw"))
        print("Restoring unpatched " + unpatched_file)
        shutil.copy(source_path, target_path)

def copy_patched():
    for patched_file in os.listdir(patched_files_dir):
        source_path = os.path.join(patched_files_dir, patched_file)
        target_path = os.path.join(system_mun_dir, patched_file)
        print("Changing Permissions of " + target_path)
        subprocess.call(("takeown", "/f",target_path))
        subprocess.call(("icacls",target_path ,"/grant","Administrators:rw"))
        print("Copying patched " + patched_file)
        shutil.copy(source_path, target_path)

def kill_processes(files_to_replace):
    for proc in psutil.process_iter():
        proc_running = True
        try:
            for open_file_path in proc.open_files():
                if proc_running == False:
                    break

                for filename in files_to_replace:
                    if filename in open_file_path.path:
                        try:
                            proc.kill()
                            print("Killed " + proc.name())
                        except:
                            print("Couldn't kill " + proc.name())
                        proc_running = False
                        break
        except:
            pass

def rebuild_icon_cache():
    iconcache_db = os.path.join(os.getenv("LOCALAPPDATA"), "IconCache.db")
    iconcache_dir = os.path.join(os.getenv("LOCALAPPDATA"), "Microsoft", "Windows", "Explorer")
    subprocess.call(("ie4uinit.exe","-show"))
    subprocess.call(("taskkill", "/f","/im","explorer.exe"))
    try:
        os.remove(iconcache_db)
    except:
        print("Couldn´t delete " + iconcache_db)
    
    for db_file in get_by_filetype(iconcache_dir, "db"):
        db_file_path = os.path.join(iconcache_dir, db_file)
        try:
            os.remove(db_file_path)
            print(db_file_path + " deleted!")
        except:
            print("Couldn´t delete " + db_file_path)

def isAdmin():
    return ctypes.windll.shell32.IsUserAnAdmin() != 0

def windows_logout():
    subprocess.call(("shutdown.exe","/l"))

# Main
def main():

    if not isAdmin():
        input("Run this program as administrator")
        sys.exit(1)
    
    check_exe_dependencies()
    create_needed_folders()
    
    print("python-7tsp by polkaulfield\n")
    print("Save everything important before running this! It will kill a some processes and logout you\n")

    if os.path.exists(unpatched_files_dir):

        print("0: Restore unpatched files")
        print("1: Apply icon pack")
   
        while True:       
                choice = int(input("Select what you wanna do: "))
                if choice == 0:
                    sys_file_list = get_by_filetype(unpatched_files_dir, "mun")
                    kill_processes(sys_file_list)
                    restore_unpatched()
                    rebuild_icon_cache()
                    print("Files restored!")
                    input("Press enter to logout...")
                    windows_logout()
                    sys.exit(0)
                if choice == 1:
                    break

    iconpacks_list = get_by_filetype(iconpacks_dir, ".7z")
    if len(iconpacks_list) == 0:
        input("No icon packs found! Place them in the iconpacks folder.")
        sys.exit(1)

    selected_pack = select_file_menu("Select an icon pack:", iconpacks_list)

    # Clear tmp folder
    if os.path.exists(tmp_dir):
        print("Deleting tmp dir")
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # Clear patched folder
    if os.path.exists(patched_files_dir):
        print("Deleting patched dir")
        shutil.rmtree(patched_files_dir)
    os.makedirs(patched_files_dir)

    extract_7z(os.path.join(iconpacks_dir, selected_pack), tmp_dir)

    res_file_list = get_by_filetype(resources_dir, "mun.res")
    
    # Populate list of system files based on the .res files from the icon pack
    sys_file_list = []
    for file in res_file_list:
        sys_file_list.append(file[:-4]) # Removing .res extension
        
    # Create folder if it doesnt exists
    if not os.path.exists(unpatched_files_dir):
        os.makedirs(unpatched_files_dir)

    # Make a backup of the unpatched system files at unpatched dir
    for res_file in res_file_list:
        system_mun_file = res_file[:-4] # Removing extension length
        system_mun_file_path = os.path.join(system_mun_dir, system_mun_file)
        unpatched_file_path = os.path.join(unpatched_files_dir, system_mun_file)

        if os.path.exists(system_mun_file_path) and not os.path.exists(unpatched_file_path):
            print("Copying " + system_mun_file_path + " into " + unpatched_file_path)
            shutil.copyfile(system_mun_file_path, unpatched_file_path)

        # Start patching
        print("Patching " + system_mun_file)
        patched_file_path = os.path.join(patched_files_dir, system_mun_file)
        res_file_path = os.path.join(resources_dir, res_file)
        subprocess.call((reshacker_path,"-open",unpatched_file_path,"-save",patched_file_path,"-action","addoverwrite","-res",res_file_path))

    # Kill processes that prevent patching
    kill_processes(sys_file_list)

    # Restore backed up system muns (Just in case the new pack doesnt have an already patched file)
    restore_unpatched()

    # Copy the newly patched files
    copy_patched()

    # Purge all the icon cache
    rebuild_icon_cache()

    input("Press enter to logout...")

    windows_logout()
    sys.exit(0)

if __name__ == "__main__":
    main()