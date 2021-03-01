import winreg
import os
import time

def findBestInstallString(uninst_str, source_str, loc_str):
    final_str = ''
    found_good_str = False

    # its probably cleaner to put this in a while loop but this works for now
    if uninst_str.lower() == 'msiexec.exe':
        found_good_str = True
        final_str = 'C:\\Windows\\System32\\msiexec.exe'

    if uninst_str[:2].lower() == 'c:':
        found_good_str = os.access(uninst_str, os.F_OK)
        if found_good_str:
            final_str = uninst_str

    if not found_good_str:
        if source_str[:2].lower() == 'c:':
            found_good_str = os.access(source_str, os.F_OK)
            if found_good_str:
                final_str = uninst_str

    if not found_good_str:
        if loc_str[:2].lower() == 'c:':
            found_good_str = os.access(loc_str, os.F_OK)
            if found_good_str:
                final_str = uninst_str

    return final_str


def formatInstallString(file_str):
    if not file_str:
        return ''

    # remove any input args to the .exe
    if 'exe' in file_str:
        exe_ind = file_str.index('exe')
        temp_str = file_str[exe_ind:]
        arg_ind = temp_str.index(' ')
        file_str = file_str[:exe_ind+arg_ind]

    if file_str.lower == 'msiexec.exe':
        file_str = 'C:\\Windows\\System32\\msiexec.exe'

    if file_str[0] == '\"':
        # Strip double quotes
        file_str = file_str.strip('\"')

    return file_str


def getInstallTime(file_str):

    if not file_str:
        return ''

    inst_time = os.path.getmtime(file_str)
    local_time = time.ctime(inst_time)
    return local_time

def printToCSV(sw_list, version_list, time_list, pub_list, save_name):
    with open(save_name, 'w') as f:
        f.write("Name\tVersion\tInstall Date\tPublisher\n")
        for ind in range(len(sw_list)):

            if sw_list[ind] == '':
                continue

            s = sw_list[ind] + '\t' + version_list[ind] + '\t' + time_list[ind] + '\t' + pub_list[ind] + '\n'
            f.write(s)


if __name__ == '__main__':

    uninstall_loc = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\uninstall"

    save_file_path = os.path.dirname(__file__)

    software_list_no_updates = []
    software_list = []
    uninstall_no_updates = []
    uninstall = []
    publisher_no_updates = []
    publisher = []
    displayVersion = []
    displayVersion_no_updates = []

    hKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_loc)
    subkey_ind = 0
    sw_ind = 0
    other_ind = 0
    while 1:
        try:
            hKey_sub = winreg.EnumKey(hKey, subkey_ind)
            hKey_sub_path = uninstall_loc + '\\' + hKey_sub
            prop_ind = 0
            str_DisplayName = ''
            str_DisplayName_no_update = ''
            str_UninstallString = ''
            str_Publisher = ''
            str_DisplayVersion = ''
            str_InstallLocation = ''
            str_InstallSource = ''

            while 1:
                try:
                    sub_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, hKey_sub_path)
                    name, value, type_ind = winreg.EnumValue(sub_key, prop_ind)

                    if name == 'DisplayName':

                        if 'Update' not in value:
                            str_DisplayName_no_update = value

                        str_DisplayName = value

                    if name == 'UninstallString':
                        str_UninstallString = formatInstallString(value)

                    if name == 'InstallLocation':
                        str_InstallLocation = formatInstallString(value)

                    if name == 'InstallSource':
                        str_InstallSource = formatInstallString(value)

                    if name == 'DisplayVersion':
                        str_DisplayVersion = value

                    if name == 'Publisher':
                        str_Publisher = value

                    prop_ind += 1
                except Exception as e:
                    if str(e) == '[WinError 259] No more data is available':
                        winreg.CloseKey(sub_key)

                        if str_DisplayName:
                            software_list.append(str_DisplayName)

                            install_time_str = findBestInstallString(str_UninstallString, str_InstallSource, str_InstallLocation)
                            if install_time_str:
                                install_time = getInstallTime(install_time_str)
                                uninstall.append(install_time)
                                uninstall_no_updates.append(install_time)
                            else:
                                uninstall.append('')
                                uninstall_no_updates.append('')
                            publisher.append(str_Publisher)
                            displayVersion.append(str_DisplayVersion)
                            software_list_no_updates.append(str_DisplayName_no_update)
                            publisher_no_updates.append(str_Publisher)
                            displayVersion_no_updates.append(str_DisplayVersion)

                        break
                    prop_ind += 1

            subkey_ind += 1
        except Exception as e:
            if str(e) == '[WinError 259] No more data is available':
                break
            subkey_ind += 1

    # now just print to csv or text
    file_with_updates = save_file_path + '/software_list_with_updates.txt'
    file_without_updates = save_file_path + '/software_list_no_updates.txt'
    file_all_keys = save_file_path + '/software_list_keys.txt'

    printToCSV(software_list, displayVersion, uninstall, publisher, file_with_updates)
    printToCSV(software_list_no_updates, displayVersion_no_updates, uninstall_no_updates, publisher_no_updates, file_without_updates)
