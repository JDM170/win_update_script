from sys import exit
from os import getcwd, remove
from os.path import isfile, exists
import ctypes
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

ps1_script_path = getcwd() + "\\ps1_script.tmp.ps1"
# ps1_script_raw = r"""Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process
ps1_script_raw = r"""{appslist}

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

bat_script_path = getcwd() + "\\bat_script.tmp.bat"
bat_script_raw = r"""@echo off

dism /cleanup-image /image:"{mntpath}" /startcomponentcleanup /resetbase /scratchdir:"C:\Windows\Temp"

dism /unmount-wim /mountdir:"{mntpath}" /commit

dism /export-image /sourceimagefile:"{wimpath}" /sourceindex:1 /destinationimagefile:"{wimpath}.esd" /Compress:recovery

dism /cleanup-wim
"""

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
    return ctypes.windll.shell32.IsUserAnAdmin() != 0


def ask_input(message, checking_func):
    user_input = input(message)
    if not user_input or len(user_input) == 0 or not checking_func(user_input):
        ask_input(message, checking_func)
    else:
        return user_input


def do_cleanup():
    if isfile(ps1_script_path):
        remove(ps1_script_path)
    if isfile(bat_script_path):
        remove(bat_script_path)
    mntpath = properties.get("mnt").get("path")
    if mntpath:
        subprocess.Popen("dism /unmount-wim /mountdir:\"{}\" /discard".format(mntpath)).wait()
        subprocess.Popen("dism /cleanup-wim").wait()


def convert_applist_to_string():
    global apps
    result = "$apps = @(\n"
    for app in apps:
        result += "\t'{}',\n".format(app)
    return result[:-2] + "\n)"


def main():
    global apps
    for keys, value in properties.items():
        if isinstance(value, dict):
            value["path"] = ask_input(value.get("message"), value.get("func"))
    if bool(input("\nWindows 10 (0) или Windows 11 (1) ?\n")):
        apps += w11_apps
    app_list = convert_applist_to_string()
    with open(ps1_script_path, mode="w") as f:
        f.write(ps1_script_raw.format(
                appslist=app_list,
                wimpath=properties.get("wim").get("path"),
                mntpath=properties.get("mnt").get("path"),
                updpath=properties.get("upd").get("path")
            )
        )
    with open(bat_script_path, mode="w") as f:
        f.write(bat_script_raw.format(
                mntpath=properties.get("mnt").get("path"),
                wimpath=properties.get("wim").get("path")
            )
        )
    # subprocess.Popen("powershell Unblock-File -Path {}".format(ps1_script_path)).wait()
    subprocess.Popen("powershell {}".format(ps1_script_path)).wait()
    subprocess.Popen("{}".format(bat_script_path)).wait()
    properties.get("wim")["path"] = ""
    properties.get("mnt")["path"] = ""
    properties.get("upd")["path"] = ""
    do_cleanup()


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
