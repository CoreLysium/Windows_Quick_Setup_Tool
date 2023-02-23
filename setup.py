import winreg
import os
import subprocess
import shutil

W  = '\033[0m'  # white (normal)
R  = '\033[31m' # red
G  = '\033[32m' # green
O  = '\033[33m' # orange
B  = '\033[34m' # blue
P  = '\033[35m' # purple

def toggle(value):
    if value == 'ENABLED':
        return 'DISABLED'
    else:
        return 'ENABLED'
    
def boolstralts(value):
    if value == 'ENABLED':
        return 'Yes'
    else:
        return 'No'

def boolstrtoint(value):
    if value == 'ENABLED':
        return 1
    else:
        return 0

textfile = open('settings.txt', 'r')
settings = textfile.readlines()
textfile.close()

textfile = open('config.txt', 'r')
config = textfile.readlines()
textfile.close()

settings = [item.strip() for item in settings]
config = [item.strip() for item in config]

os.system('cls')
os.system('title Windows Quick Setup Tool')
os.system('ver')
print('Windows Quick Setup Tool Version 1.0 \nCopyright(c) CoreLysium. All rights reserved.')

option = input('\nSelect an action from the list below:\n1. Run Windows Quick Setup.\n2. Change default software list.\n3. Change default windows configuration.\n4. Change tool settings.\n0. Exit\n\n')

while not option == '0':
    match option:
        case '1':
            os.system('cls')

            useRecommendedSettings = input('\nWould you like to change windows settings to the recommended defaults? [Y]es or [N]o? \n')

            if useRecommendedSettings == 'Y' or useRecommendedSettings == 'y':

                print('\nUpdating Theme...')

                try:

                    user = winreg.HKEY_CURRENT_USER
                    theme = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\Windows\\CurrentVersion\\Themes\\Personalize", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(theme, "AppsUseLightTheme", 0, winreg.REG_DWORD, boolstrtoint(config[0]))
                    winreg.SetValueEx(theme, "SystemUsesLightTheme", 0, winreg.REG_DWORD, boolstrtoint(config[0]))

                    if theme:
                        winreg.CloseKey(theme)

                    print(G + 'Updated Theme Successfully!' + W)
                
                except Exception as e:
                    print(R + str(e) + W)

                print('\nUpdating file explorer options...')

                try:

                    explorer = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(explorer, "HideFileExt", 0, winreg.REG_DWORD, boolstrtoint(config[1]))
                    winreg.SetValueEx(explorer, "Hidden", 0, winreg.REG_DWORD, boolstrtoint(config[2]))
                    winreg.SetValueEx(explorer, "LaunchTo", 0, winreg.REG_DWORD, boolstrtoint(config[3]))

                    if explorer:
                        winreg.CloseKey(explorer)

                    explorer = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(explorer, "ShowFrequent", 0, winreg.REG_DWORD, boolstrtoint(config[4]))
                    winreg.SetValueEx(explorer, "ShowRecent", 0, winreg.REG_DWORD, boolstrtoint(config[5]))

                    if explorer:
                        winreg.CloseKey(explorer)

                    print(G + 'Updated Explorer Options Successfully!' + W)
                except Exception as e:
                    print(R + str(e) + W)

                print('\nUpdating Power Plan...')

                try:

                    os.system('powercfg -restoredefaultschemes')
                    os.system('powercfg -import "' + os.getcwd() + '\\"' + config[6] + ' 4a48d9fe-6e0a-4d02-bbb0-79c32d314f39')
                    os.system('powercfg -SETACTIVE 4a48d9fe-6e0a-4d02-bbb0-79c32d314f39')

                    print(G + 'Updated Power Plan Successfully!' + W)
                except Exception as e:
                    print(R + str(e) + W)

                print('\nUpdating Network Rules...')

                try:

                    os.system('netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=' + boolstralts(config[7]))
                    os.system('netsh advfirewall firewall set rule group="Network Discovery" new enable=' + boolstralts(config[8]))

                    print(G + 'Updated Network Options Successfully!' + W)
                except Exception as e:
                    print(R + str(e) + W)

                print('\nUpdating Notification Settings...')

                try:

                    policies = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PushNotifications", 0, winreg.KEY_ALL_ACCESS)

                    winreg.SetValueEx(policies, "ToastEnabled", 0, winreg.REG_DWORD, boolstrtoint(config[9]))

                    if policies:
                        winreg.CloseKey(policies)

                    policies = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager", 0, winreg.KEY_ALL_ACCESS)

                    winreg.SetValueEx(policies, "SubscribedContent-310093Enabled", 0, winreg.REG_DWORD, boolstrtoint(config[9]))
                    winreg.SetValueEx(policies, "SubscribedContent-338389Enabled", 0, winreg.REG_DWORD, boolstrtoint(config[9]))

                    if policies:
                        winreg.CloseKey(policies)

                    policies = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\UserProfileEngagement", 0, winreg.KEY_ALL_ACCESS)

                    winreg.SetValueEx(policies, "ScoobeSystemSettingEnabled", 0, winreg.REG_DWORD, boolstrtoint(config[9]))

                    if policies:
                        winreg.CloseKey(policies)

                    print(G + 'Updated Notification Settings Successfully!' + W)
                
                except Exception as e:
                    print(R + str(e) + W)

                print('\nUpdating Desktop Icons...')

                try:

                    icons_1 = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\HideDesktopIcons\\NewStartPanel", 0, winreg.KEY_SET_VALUE)
                    icons_2 = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\HideDesktopIcons\\ClassicStartMenu", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(icons_1, "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_1, "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_1, "{59031a47-3f72-44a7-89c5-5595fe6b30ee}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_1, "{645FF040-5081-101B-9F08-00AA002F954E}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_1, "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))

                    winreg.SetValueEx(icons_2, "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_2, "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_2, "{59031a47-3f72-44a7-89c5-5595fe6b30ee}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_2, "{645FF040-5081-101B-9F08-00AA002F954E}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))
                    winreg.SetValueEx(icons_2, "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", 0, winreg.REG_DWORD, boolstrtoint(config[10]))

                    if icons_1:
                        winreg.CloseKey(icons_1)
                    if icons_2:
                        winreg.CloseKey(icons_2)

                    print(G + 'Updated Desktop Icons Successfully!' + W)
                except Exception as e:
                    print(R + str(e) + W)

                print('\nUpdating taskbar settings...')

                try:

                    taskbar = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(taskbar, "TaskbarAl", 0, winreg.REG_DWORD, boolstrtoint(config[11]))
                    winreg.SetValueEx(taskbar, "ShowCortanaButton", 0, winreg.REG_DWORD, 0)
                    winreg.SetValueEx(taskbar, "ShowTaskViewButton", 0, winreg.REG_DWORD, 0)

                    if taskbar:
                        winreg.CloseKey(taskbar)

                    taskbar = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\Windows\\CurrentVersion\\Search", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(taskbar, "SearchboxTaskbarMode", 0, winreg.REG_DWORD, boolstrtoint(config[12]))

                    if taskbar:
                        winreg.CloseKey(taskbar)

                    taskbar = winreg.CreateKeyEx(user, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Feeds", 0, winreg.KEY_SET_VALUE)

                    winreg.SetValueEx(taskbar, 'ShellFeedsTaskbarViewMode', 0, winreg.REG_DWORD, 2)

                    if taskbar:
                        winreg.CloseKey(taskbar)
                    
                    print(G + 'Taskbar Settings Updated Successfully!' + W)
                except Exception as e:
                    print(R + str(e) + W)

            installOptions = input('Would you like to install recommended software? [Y]es or [N]o? \n')
            if installOptions == 'Y' or installOptions == 'y':
                textfile = open('software.txt', 'r')

                subprocess.call('powershell Set-ExecutionPolicy AllSigned')
                subprocess.call("powershell Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))")
                subprocess.call('RefreshEnv.cmd')
                subprocess.call('choco feature enable -n=allowGlobalConfirmation')
                subprocess.call('choco install ' + textfile.readline())
                
                additionalcmds = textfile.readline()
                while not additionalcmds == '':
                    subprocess.call(additionalcmds)
                    additionalcmds = textfile.readline()

                textfile.close()
                if settings[0] == 'ENABLED':
                    print('Removing chocolatey and chocolatey created files...')
                    try:
                    
                        shutil.rmtree('C:\\ProgramData\\chocolatey')
                        shutil.rmtree('C:\\Users\\' + os.getlogin() +'\\AppData\\Local\\Temp\\chocolatey')

                        print(G + 'Chocolatey removed successfully!' + W)
                    
                    except Exception as e:
                        print(R + str(e) + W)

            input(P + '*PRESS ENTER TO CONTINUE*' + W)
        case '2':
            os.system('cls')

            print(P + '*DISCLAIMER*\n' + W + 'Software is retrieved using chocolatey and the chocolatey public repository.\nAll rights belong to the softwares respective owners.\n')
            textfile = open('software.txt', 'r')
            text = textfile.readlines()
            if len(text) > 1:
                print('Software List: ' + text[0])
                print('Additional Chocolatey Commands: ')
                for i in range(1,len(text)):
                    print(text[i])

            option = input('\nWould you like to change the list? [Y]es or [N]o?\n')

            if option == 'y' or option == 'Y':
                new_list = input('Enter new software list separating each program with a space: \n')
                text[0] = new_list

            option = input('Would you like to change the additional commands? [Y]es or [N]o?\n')
            counter = 1
            while option == 'y' or option == 'Y':
                command = input('Enter command number ' + str(counter) + ':\n')
                
                if counter > len(text) - 1:
                    text.append(command)
                else:
                    text[counter] = command
                
                counter += 1
                option = input('Would you like to add another command? [Y]es or [N]o?\n')

            textfile.close()

            textfile = open('software.txt', 'w')
            textfile.writelines(text)
            textfile.close()
            print(G + 'Sofware Settings Updated!' + W)
            input(P + '*PRESS ENTER TO CONTINUE*' + W)
        case '3':
            os.system('cls')
            print('Select a setting or action from the list below:')
            print('1. Light Theme. ' + B + '[' + config[0] + ']' + W)
            print('2. Hide file extensions for known files. ' + B + '[' + config[1] + ']' + W)
            print('3. Show hidden files. ' + B + '[' + config[2] + ']' + W)
            print('4. Launch file explorer to "This PC". ' + B + '[' + config[3] + ']' + W)
            print('5. Show frequent files in quick access. ' + B + '[' + config[4] + ']' + W)
            print('6. Show recent files in quick access. ' + B + '[' + config[5] + ']' + W)
            print('7. Power Plan File: ' + B + config[6] + W)
            print('8. File and Printer Sharing. ' + B + '[' + config[7] + ']' + W)
            print('9. Network Discovery. ' + B + '[' + config[8] + ']' + W)
            print('10. Notifications. ' + B + '[' + config[9] + ']' + W)
            print('11. Hide Desktop Icons. ' + B + '[' + config[10] + ']' + W)
            print('12. Taskbar aligned center. ' + P + '(Windows 11 only!) ' + B + '[' + config[11] + ']' + W)
            print('13. Search on Taskbar. ' + B + '[' + config[12] + ']' + W)
            option = input('0. Return to Main Menu\n\n')
            
            while not option == '0':
                try:
                    stnum = int(option) - 1
                    if config[stnum] == 'ENABLED' or config[stnum] == 'DISABLED':
                        config[stnum] = toggle(config[stnum])
                    else:
                        config[stnum] = input('Enter new value:')
                except Exception as e:
                    print(R + str(e) + W)
                
                os.system('cls')
                print('Select a setting or action from the list below:')
                print('1. Light Theme. ' + B + '[' + config[0] + ']' + W)
                print('2. Hide file extensions for known files. ' + B + '[' + config[1] + ']' + W)
                print('3. Show hidden files. ' + B + '[' + config[2] + ']' + W)
                print('4. Launch file explorer to "This PC". ' + B + '[' + config[3] + ']' + W)
                print('5. Show frequent files in quick access. ' + B + '[' + config[4] + ']' + W)
                print('6. Show recent files in quick access. ' + B + '[' + config[5] + ']' + W)
                print('7. Power Plan File: ' + B + config[6] + W)
                print('8. File and Printer Sharing. ' + B + '[' + config[7] + ']' + W)
                print('9. Network Discovery. ' + B + '[' + config[8] + ']' + W)
                print('10. Notifications. ' + B + '[' + config[9] + ']' + W)
                print('11. Hide Desktop Icons. ' + B + '[' + config[10] + ']' + W)
                print('12. Taskbar aligned center. ' + P + '(Windows 11 only!) ' + B + '[' + config[11] + ']' + W)
                print('13. Search on Taskbar. ' + B + '[' + config[12] + ']' + W)
                option = input('0. Return to Main Menu\n\n')
            
            textfile = open('config.txt','w')
            textfile.writelines("\n".join(config))
            textfile.close()
        case '4':
            os.system('cls')
            option = input('Select a setting or action from the list below:\n1. Remove Chocolatey Files After Install. ' + B + '[' + settings[0] + ']' + W + '\n0. Return to Main Menu\n')
            while not option == '0':
                match option:
                    case '1':
                        settings[0] = toggle(settings[0])
                print(settings[0])
                textfile = open('settings.txt', 'w')
                textfile.writelines("\n".join(settings))
                textfile.close()
                os.system('cls')
                option = input('Select a setting or action from the list below:\n1. Remove Chocolatey Files After Install. ' + B + '[' + settings[0] + ']' + W + '\n0. Return to Main Menu\n')
                        
    os.system('cls')
    option = input('\nSelect an action from the list below:\n1. Run Windows Quick Setup.\n2. Change default software list.\n3. Change default windows configuration.\n4. Change tool settings.\n0. Exit\n\n')

os.system('cls')