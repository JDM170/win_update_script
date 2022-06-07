# Скрипт для внесения изменений в WIM-Файл (W10)

# Список приложений для удаления
$apps = @(
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
)


write-host ""
write-host "Введите полный путь до wim-файла:"
$wimpath = read-host

write-host ""
write-host "Введите полный путь до папки монтирования:"
$mntpath = read-host

write-host ""
write-host "Введите полный путь до папки с обновлениями:"
$updpath = read-host

write-host ""
write-host "Нажмите Enter для запуска"
read-host

$start_of_changes = [int](Get-Date -UFormat %s -Millisecond 0)

write-host ""
write-host "Удаляем лишние образы систем..."
Remove-WindowsImage -ImagePath $wimpath -Index 3
Remove-WindowsImage -ImagePath $wimpath -Index 2
Remove-WindowsImage -ImagePath $wimpath -Index 1

write-host ""
write-host "Монтируем образ..."
Mount-WindowsImage -ImagePath $wimpath -Index 1 -Path $mntpath

write-host ""
write-host "Удаляем установщик OneDrive из автозапуска..."
# Remove-Item "$mntpath\Windows\SysWOW64\OneDrive.ico" -Force
# Remove-Item "$mntpath\Windows\SysWOW64\OneDriveSetup.exe" -Force
# Remove-Item "$mntpath\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk" -Force
reg load HKLM\onedrive_delete $mntpath\Users\default\ntuser.dat
reg delete HKLM\onedrive_delete\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\ /v OneDriveSetup /f
reg unload HKLM\onedrive_delete

write-host ""
write-host "Удаляем лишние системные приложения..."
Get-AppxProvisionedPackage -Path $mntpath | ForEach-Object {
        if ($apps -contains $_.DisplayName) {
                Write-Host Removing $_.DisplayName...
                Remove-AppxProvisionedPackage -Path $mntpath -PackageName $_.PackageName | Out-Null
            }
        }

# write-host "Пытаемся удалить Cortana..."
# Remove-AppxProvisionedPackage -Path $mntpath -PackageName Microsoft.549981C3F5F10

write-host ""
write-host "Интегрируем обновления..."
Add-WindowsPackage -Path $mntpath -PackagePath $updpath

# write-host ""
# write-host "Сканируем здоровье образа..."
# Repair-WindowsImage -Path $mntpath -ScanHealth

# write-host ""
# write-host "\nПробуем запустить очистку образа от старых файлов..."
# Dism.exe /image:%mntpath% /cleanup-image /startcomponentcleanup /resetbase

write-host ""
write-host "Сохраняем изменения..."
Dismount-WindowsImage -Path $mntpath -Save
Clear-WindowsCorruptMountPoint

$end_of_changes = [int](Get-Date -UFormat %s -Millisecond 0)
write-host "Затраченное время (в минутах): " (($end_of_changes - $start_of_changes) / 60)

[gc]::Collect()
