# Добавление обновлений в образ

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
write-host "Монтируем образ..."
Mount-WindowsImage -ImagePath $wimpath -Index 1 -Path $mntpath

write-host ""
write-host "Интегрируем обновления..."
Add-WindowsPackage -Path $mntpath -PackagePath $updpath

write-host ""
write-host "Сохраняем изменения..."
Dismount-WindowsImage -Path $mntpath -Save
Clear-WindowsCorruptMountPoint

$end_of_changes = [int](Get-Date -UFormat %s -Millisecond 0)
write-host "Затраченное время (в минутах): " (($end_of_changes - $start_of_changes) / 60)

[gc]::Collect()
