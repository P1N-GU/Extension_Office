# Extension_Office
This application alter register Windows for permitted open files Office in Google Workspace.

To add one new register add in "string file_content" the type extension, copy line alow and cute in variable

REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\".extension"\OpenWithProgids /v "GoogleDriveFS.gdoc or gsheets or gslides" /t REG_BINARY /d 00000000 /f
REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\".extension"\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f

Example:

REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.docx\OpenWithProgids /v GoogleDriveFS.gdoc /t REG_BINARY /d 00000000 /f
REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.docx\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
