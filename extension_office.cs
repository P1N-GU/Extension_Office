using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Principal;
using System.Diagnostics;

namespace Extension_Office
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create folder
            string folder_path = @"C:\";

            string folder_name = "extension_office";

            string full_path = Path.Combine(folder_path, folder_name);

            try
            {
                if (!Directory.Exists(full_path))
                {
                    Directory.CreateDirectory(full_path);

                    Console.WriteLine("Folder created successfully: " + full_path);
                }
                else
                {
                    Console.WriteLine("Folder already exists: " + full_path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error created folder: " + ex.Message);
            }

            //Create .BAT
            string file_path = @"C:\extension_office\extension_office.bat";

            //Line low to add register in Windows for permited files in Google Workspace
            string file_content = @"
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.docx\OpenWithProgids /v GoogleDriveFS.gdoc /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.docx\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.doc\OpenWithProgids /v GoogleDriveFS.gdoc /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.doc\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dot\OpenWithProgids /v GoogleDriveFS.gdoc /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dot\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xlsx\OpenWithProgids /v GoogleDriveFS.gsheet /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xlsx\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xls\OpenWithProgids /v GoogleDriveFS.gsheet /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xls\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xlsm\OpenWithProgids /v GoogleDriveFS.gsheet /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xlsm\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.csv\OpenWithProgids /v GoogleDriveFS.gsheet /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.csv\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.csv\OpenWithProgids /v GoogleDriveFS.gsheet /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.csv\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pptx\OpenWithProgids /v GoogleDriveFS.gslides /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pptx\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ppt\OpenWithProgids /v GoogleDriveFS.gslides /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ppt\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.rtf\OpenWithProgids /v GoogleDriveFS.gdoc /t REG_BINARY /d 00000000 /f
            REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.rtf\OpenWithList /v b /t REG_SZ /d GoogleDriveFS.exe /f
            ";
            try
            {
                File.WriteAllText(file_path, file_content);

                Console.WriteLine("File .bat created successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error created file .bat: " + ex.Message);
            }


            //Create .VBS
            string startupPath = Path.Combine(
                "C:\\Users\\All Users\\Microsoft\\Windows\\Start Menu\\Programs\\Startup",
                "start_extension_office.vbs");

            string scriptContent = @"Set WshShell = CreateObject(""WScript.Shell"" )
WshShell.Run chr(34) & ""C:\extension_office\extension_office.bat"" & Chr(34), 0 
Set WshShell = Nothing";

            Directory.CreateDirectory(Path.GetDirectoryName(startupPath));

            File.WriteAllText(startupPath, scriptContent);

            Console.WriteLine("Script VBS created in: " + startupPath);
        }
    }
}