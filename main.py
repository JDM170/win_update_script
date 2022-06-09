from sys import exit
from os import getcwd, remove, system
from os.path import isfile, exists
from time import time
from ctypes import windll
import subprocess


properties = {
    "wim": {
        "message": "\nВведите путь до wim-файла:\n",
        "path": "",
        "func": isfile
    },
    "mnt": {
        "message": "\nВведите путь до папки монтирования:\n",
        "path": "",
        "func": exists
    },
    "upd": {
        "message": "\nВведите путь до папки с обновлениями:\n",
        "path": "",
        "func": exists
    }
}

ps_script_path = getcwd() + "\\ps1_script.tmp.ps1"
# ps_script_raw = r"""Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process
ps_script_raw = r"""{appslist}

Remove-WindowsImage -ImagePath {wimpath} -Index 3
Remove-WindowsImage -ImagePath {wimpath} -Index 2
Remove-WindowsImage -ImagePath {wimpath} -Index 1

Mount-WindowsImage -ImagePath {wimpath} -Index 1 -Path {mntpath}

reg load HKLM\onedrive_delete {mntpath}\Users\default\ntuser.dat
reg delete HKLM\onedrive_delete\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\ /v OneDriveSetup /f
reg unload HKLM\onedrive_delete

Get-AppxProvisionedPackage -Path {mntpath} | ForEach-Object {{
        if ($apps -contains $_.DisplayName) {{
                Write-Host Removing $_.DisplayName...
                Remove-AppxProvisionedPackage -Path {mntpath} -PackageName $_.PackageName | Out-Null
            }}
        }}

Add-WindowsPackage -Path {mntpath} -PackagePath {updpath}
"""

ps_script_additional = r"""
Dismount-WindowsImage -Path {mntpath} -Save
Clear-WindowsCorruptMountPoint

Mount-WindowsImage -ImagePath {wimpath} -Index 1 -Path {mntpath}

Add-WindowsPackage -Path {mntpath} -PackagePath {sec_updpath}
"""

bat_script_path = getcwd() + "\\bat_script.tmp.bat"
bat_script_raw = r"""@echo off

dism /cleanup-image /image:"{mntpath}" /startcomponentcleanup /resetbase /scratchdir:"C:\Windows\Temp"

dism /unmount-wim /mountdir:"{mntpath}" /commit

dism /cleanup-wim
"""
# dism /export-image /sourceimagefile:"{wimpath}" /sourceindex:1 /destinationimagefile:"{wimpath}.esd" /Compress:recovery

apps = [
    "Microsoft.BingWeather",
    "Microsoft.GetHelp",
    "Microsoft.Getstarted",
    "Microsoft.Messaging",
    "Microsoft.Microsoft3DViewer",
    "Microsoft.MixedReality.Portal",
    "Microsoft.MicrosoftOfficeHub",
    "Microsoft.MicrosoftSolitaireCollection",
    "Microsoft.MicrosoftStickyNotes",
    "Microsoft.MSPaint",#Paint 3D (Windows 10)
    "Microsoft.Office.OneNote",
    "Microsoft.OneConnect",
    "Microsoft.People",
    "Microsoft.ScreenSketch",#Скриншоты (Windows 10 1809+)
    "Microsoft.YourPhone",#Ваш телефон (Windows 10 1809+)
    "Microsoft.Print3D",
    "Microsoft.SkypeApp",
    "Microsoft.ZuneMusic",
    "Microsoft.ZuneVideo",#Кино и ТВ
    "Microsoft.XboxApp"#Xbox (Windows 10)
    # "Microsoft.Windows.Cortana" # ?????
]

w11_apps = [
    "Microsoft.WindowsAlarms",
    "Microsoft.WindowsCamera",
    "microsoft.windowscommunicationsapps",
    "Microsoft.WindowsFeedbackHub",
    "Microsoft.WindowsMaps",
    "Microsoft.WindowsSoundRecorder",
    "Microsoft.GamingApp",#Xbox (Windows 11)
    "Microsoft.PowerAutomateDesktop",#(Windows11)
    "Microsoft.Todos",#(Windows11)
    "Microsoft.BingNews",#Новости (Windows 11)
    "MicrosoftWindows.Client.WebExperience",#Виджеты (Windows 11)
    "Microsoft.Paint"#Paint (Windows11)
]


def is_admin():
    return windll.shell32.IsUserAnAdmin() != 0


def ask_input(message, checking_func):
    user_input = input(message)
    if not user_input or len(user_input) == 0 or not checking_func(user_input):
        ask_input(message, checking_func)
    else:
        return user_input


def convert_applist_to_string(is_w11):
    global apps, w11_apps
    result = "$apps = @(\n"
    for app in apps if not is_w11 else apps + w11_apps:
        result += "\t'{}',\n".format(app)
    return result[:-2] + "\n)"


def do_cleanup():
    if isfile(ps_script_path):
        remove(ps_script_path)
    if isfile(bat_script_path):
        remove(bat_script_path)
    mntpath = properties.get("mnt").get("path")
    if mntpath:
        subprocess.Popen("dism /unmount-wim /mountdir:\"{}\" /discard".format(mntpath)).wait()
        subprocess.Popen("dism /cleanup-wim").wait()


def main():
    # Input values
    for keys, value in properties.items():
        if isinstance(value, dict):
            value["path"] = ask_input(value.get("message"), value.get("func"))
    additional_upd = input("\nВведите путь для второй папки с обновлениями (чтобы пропустить нажмите Enter):\n")
    is_w11 = bool(input("\nWindows 10 (0) или Windows 11 (1) ?\n"))
    system("pause")
    # Start process
    time_start = time()
    appslist = convert_applist_to_string(is_w11)
    with open(ps_script_path, mode="w") as f:  # Create powershell script file
        wim = properties.get("wim").get("path")
        mnt = properties.get("mnt").get("path")
        f.write(ps_script_raw.format(
                appslist=appslist,
                wimpath=wim,
                mntpath=mnt,
                updpath=properties.get("upd").get("path")
            )
        )
        if additional_upd and exists(additional_upd):
            f.write(ps_script_additional.format(
                    wimpath=wim,
                    mntpath=mnt,
                    sec_updpath=additional_upd
                )
            )
    with open(bat_script_path, mode="w") as f:  # Create bat script file
        f.write(bat_script_raw.format(
                mntpath=properties.get("mnt").get("path"),
                wimpath=properties.get("wim").get("path")
            )
        )
    # subprocess.Popen("powershell Unblock-File -Path {}".format(ps_script_path)).wait()
    subprocess.Popen("powershell {}".format(ps_script_path)).wait()
    subprocess.Popen("{}".format(bat_script_path)).wait()
    properties.get("wim")["path"] = ""
    properties.get("mnt")["path"] = ""
    properties.get("upd")["path"] = ""
    do_cleanup()
    print("Затраченное время:", ((time()-time_start) / 60))


if __name__ == '__main__':
    try:
        if not is_admin():
            print("Запустите программу от имени администратора!")
            raise KeyboardInterrupt
        else:
            main()
    except KeyboardInterrupt:
        do_cleanup()
        exit()
