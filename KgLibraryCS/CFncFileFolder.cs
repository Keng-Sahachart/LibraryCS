using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

//namespace KengsLibraryCs
//{
    public class CFncFileFolder
    {
        // #######################################################################################
        // #########  เกี่ยวกับ ไฟล์ โฟลเดอร์      ##########################################################
        // #######################################################################################
        public static bool FolderExists(string sPath)
        {
            return Directory.Exists(sPath);
        }
        public static bool IsFile(string sPath) // /C#
        {
            return File.Exists(sPath);
        }
        public static void CreateFolder(string Path)
        {
            //My.MyProject.Computer.FileSystem.CreateDirectory(Path);
            Directory.CreateDirectory(Path);
        }


        /// <summary>
    /// สร้างทุกๆ Folder และ Sub Folder จาก Path
    /// </summary>
    /// <param name="Path"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool CreateFolderAndSub(string Path)
        {
            try
            {
                var pathParts = Path.Split(Conversions.ToChar(@"\\"));
                // RunPath
                for (int i = 0, loopTo = pathParts.Length - 1; i <= loopTo; i++)
                {
                    if (i > 0)
                        // มีปัญหา ตอน Combine กับ Letter ของ Drive ใช้แบบ เชื่อมสตริงไปก่อน
                        // pathParts(i) = System.IO.Path.Combine(pathParts(i - 1), pathParts(i)) 
                        pathParts[i] = pathParts[i - 1] + @"\" + pathParts[i];

                    if (!Directory.Exists(pathParts[i]))
                        Directory.CreateDirectory(pathParts[i]);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

            return false;
        }
        /// <summary>
        /// เปิด Folder ขึ้นมาด้วย Windows Explorer
        /// </summary>
        /// <param name="PathToOpen"></param>
        public static void OpenFolder(string PathToOpen)
        {
            //var info = My.MyProject.Computer.FileSystem.GetFileInfo(PathToOpen);
            System.IO.FileInfo info = new FileInfo(PathToOpen);
            string PathFolder;

            if (Directory.Exists(PathToOpen))
                PathFolder = PathToOpen;
            else
                PathFolder = info.DirectoryName;

            /* หลัง update windows แล้วใช้ไม่ได้
             Process.Start("explorer.exe ","@"+PathToOpen); 
             Process.Start("explorer.exe", "/select," & PathToOpen)*/
            Process.Start("\"" + PathFolder + "\"");
        }
        public static long GetFileSizeInKB(string FilePath)
        {
            // Imports System.IO 
            try
            {
                var MyFile = new FileInfo(FilePath);
                long FileSize = Conversions.ToLong(Math.Ceiling(MyFile.Length / (double)1024)); // * 1024)))
                return FileSize;
            }
            catch
            {
                return 0;
            }
        }
        /// <summary>เอาชื่อไฟล์ ออกมาจาก Path</summary>
    /// <param name="FilePath" >Path ที่ต้องการแยกชื่อไฟล์ออกมา</param>
    /// <returns>ชื่อไฟล์ </returns>
        public static string GetFileName(string FilePath)
        {
            return Path.GetFileName(FilePath);
        }

        public static string GetFullPathWithoutExtension(string FilePath)
        {
            string FullPathWithoutExtension;
            string Path;
            string FileName;
            Path = System.IO.Path.GetDirectoryName(FilePath);
            FileName = System.IO.Path.GetFileNameWithoutExtension(FilePath);
            FullPathWithoutExtension = System.IO.Path.Combine(Path, FileName);
            return FullPathWithoutExtension;
        }

        /// <summary>สำหรับสร้างชื่อไฟล์ใหม่ไม่ให้ซ้ำกับไฟล์เดิม เพื่อเก็บไฟล์ตัวเก่าไว้ 
        /// \n ส่ง Path ไฟล์เข้าไปแล้ว เช็กดูว่า มีไฟล์แล้วหรือยัง แล้ว เปลี่ยนชื่อไม่ให้ซ้ำ </summary>
        /// <param name="PathFileToSave"> Path ไฟล์ ที่ต้องการเช็ก</param>
        /// <returns>ได้ชื่อไฟล์ใหม่ ที่ไม่ซ้ำกับไฟล์ที่มีอยู่ หรือชื่อเดิมถ้าไม่มีไฟล์</returns>
        public static string NewFileNameUnique(string PathFileToSave)
        {
            int Counter = 0;
            string NewFileName = PathFileToSave;
            string FileNameNoExt = Path.GetFileNameWithoutExtension(PathFileToSave);
            string Extention = Path.GetExtension(PathFileToSave);
            string FilePath = Path.GetDirectoryName(PathFileToSave);
            while (File.Exists(NewFileName))
            {
                Counter = Counter + 1;
                NewFileName = Path.Combine(FilePath, string.Format("{0}_{1}", FileNameNoExt, Counter.ToString()) + Extention);
            }
            return NewFileName;
        }
        public static object NewFolderNameUnique(string PathFolder)
        {
            int Counter = 0;
            string NewPathFolder = PathFolder;
            while (FolderExists(NewPathFolder))
            {
                Counter = Counter + 1;
                NewPathFolder = string.Format("{0}_{1}", PathFolder, Counter.ToString());
            }
            return NewPathFolder;
        }
        /// <summary>เรียกไฟล์</summary>
        public static void RunFile(string PathFile)
        {
            Process.Start(PathFile);
        }

        /// <summary>
    /// กำหนด Path File
    /// ข้อเสียคือ ถ้าไม่เลือกไฟล์ ค่า เก่า ของ Text จะเปลี่ยน ด้วย Nothing
    /// </summary>
    /// <param name="StrFilter">"ตัวอย่าง ไฟล์ Excel 2003(*.xls)|*.xls|ไฟล์ Excel 2007(*.xlsx)|*.xlsx"</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static object GetPathFile(string StrFilter = "All files (*.*)|*.*")
        {
            var dialogOpenFiles = new OpenFileDialog();
            // dialogOpenFiles.Multiselect = True
            dialogOpenFiles.Filter = StrFilter;
            try
            {
                if (dialogOpenFiles.ShowDialog() == DialogResult.OK)
                    return dialogOpenFiles.FileNames[0];
            }
            catch
            {
                Interaction.Beep();
            }
            return null;
        }
        /// <summary>
    /// เปิดหน้าต่างเลือกไฟล์ เพื่อเอา PAth ใส่ Text
    /// ข้อดี เมื่อไม่เลือกไฟล์ แล้ว TextBox จะคงค่าเหมือนเดิม
    /// </summary>
    /// <param name="TxtBoxPath">TextBox</param>
    /// <param name="StrFilter">"ตัวอย่าง ไฟล์ Excel 2003(*.xls)|*.xls|ไฟล์ Excel 2007(*.xlsx)|*.xlsx"</param>
        public static void GetPathFileToTxtBx(ref TextBox TxtBoxPath, string StrFilter = "All files (*.*)|*.*", string DefaultPath = null)
        {
            var dialogOpenFiles = new OpenFileDialog();
            // dialogOpenFiles.Multiselect = True
            dialogOpenFiles.Filter = StrFilter;
            if (DefaultPath != null)
                dialogOpenFiles.InitialDirectory = DefaultPath;
            try
            {
                if (dialogOpenFiles.ShowDialog() == DialogResult.OK)
                    TxtBoxPath.Text = dialogOpenFiles.FileNames[0];
            }
            catch
            {
                Interaction.Beep();
            }
        }
        public static void GetPathFolderToTxtBx(ref TextBox TxtBoxPath)
        {
            var FolderBrowserDialog = new FolderBrowserDialog();

            try
            {
                if (FolderBrowserDialog.ShowDialog() ==DialogResult.OK)
                    TxtBoxPath.Text = FolderBrowserDialog.SelectedPath;
            }
            catch
            {
                Interaction.Beep();
            }
        }

    public static DateTime GetFile_CreationTime(string Path)
    {
        FileInfo fInfo = new FileInfo(Path);
        return fInfo.CreationTime;
    }
    public static DateTime GetFile_LastWriteTime(string Path)
    {
        FileInfo fInfo = new FileInfo(Path);
        return fInfo.LastWriteTime;
    }
    public static DateTime GetFile_LastAccessTime(string Path)
    {
        FileInfo fInfo = new FileInfo(Path);
        return fInfo.LastAccessTime;
    }
}
//}
