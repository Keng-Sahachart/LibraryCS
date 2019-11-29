using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
//using Microsoft.VisualBasic;
using System.Diagnostics;

namespace KengsLibraryCs
{
    public class CFncFileFolder_
    {
        public static bool FolderExists(string sPath)
        {
            return Directory.Exists(sPath);
        }
        public static bool FolderExist(string sPath)
        {
            return File.Exists(sPath);
        }


        /// <summary>สำหรับสร้างชื่อไฟล์ใหม่ไม่ให้ซ้ำกับไฟล์เดิม เพื่อเก็บไฟล์ตัวเก่าไว้ 
        /// \n ส่ง Path ไฟล์เข้าไปแล้ว เช็กดูว่า มีไฟล์แล้วหรือยัง แล้ว เปลี่ยนชื่อไม่ให้ซ้ำ </summary>
        /// <param name="PathFileToSave"> Path ไฟล์ ที่ต้องการเช็ก</param>
        /// <returns>ได้ชื่อไฟล์ใหม่ ที่ไม่ซ้ำกับไฟล์ที่มีอยู่ หรือชื่อเดิมถ้าไม่มีไฟล์</returns>
        public static string NewFileNameUnique(string PathFileToSave)
        {
            int Counter = 0;
            string NewFileName = PathFileToSave;
            string FileNameNoExt = System.IO.Path.GetFileNameWithoutExtension(PathFileToSave);
            string Extention = System.IO.Path.GetExtension(PathFileToSave);
            string FilePath = System.IO.Path.GetDirectoryName(PathFileToSave);
            while (File.Exists(NewFileName))
            {
                Counter = Counter + 1;
                NewFileName = System.IO.Path.Combine(FilePath, string.Format("{0}_{1}", FileNameNoExt, Counter.ToString()) + Extention);
            }
            return NewFileName;
        }

        public static void OpenFolder(string PathToOpen)
        {
            System.IO.FileInfo info = new FileInfo(PathToOpen);
            string PathFolder;

            if (Directory.Exists(PathToOpen))
                PathFolder = PathToOpen;
            else
                PathFolder = info.DirectoryName;

            /*หลัง update windows แล้วใช้ไม่ได้
             Process.Start("explorer.exe ","@"+PathToOpen); 
             Process.Start("explorer.exe", "/select," & PathToOpen)*/
            Process.Start("\"" + PathFolder + "\"");
        }


    }
}
