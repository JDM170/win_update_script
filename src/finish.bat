@echo off

set /P wim="Введите путь до wim-файла: "
set /P mnt="Введите путь до папки монтирования: "

echo
echo Mounting...
dism /mount-wim /wimfile:%wim% /index:1 /mountdir:%mnt%

echo
echo Starting cleanup image...
dism /cleanup-image /image:%mnt% /startcomponentcleanup /resetbase /scratchdir:"C:\Windows\Temp"

echo
echo Unmounting...
dism /unmount-wim /mountdir:%mnt% /commit

echo
echo Exporting complete image...
dism /export-image /sourceimagefile:%wim% /sourceindex:1 /destinationimagefile:%wim%+".fin"

echo
echo Cleaning up...
dism /cleanup-wim