using System.Data;
using Microsoft.VisualBasic;
using System;
using System.Text;
using System.IO;
//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{
    public static class CFncTextFile
    {
        /// <summary>
    /// อ่านไฟล์ ออกมาเป็น String 
    /// </summary>
    /// <param name="PathTxtFile">ตำแหน่งของไฟล์ Text </param>
    /// <returns>String</returns>
        public static string ReadTextFileToString(string PathTxtFile)
        {
            string RetStr = null;
            foreach (var THisLine in System.IO.File.ReadAllText(PathTxtFile).Split(System.Convert.ToChar(Environment.NewLine)))
                RetStr += THisLine;

            return RetStr;
        }

        /// <summary>
    /// อ่านไฟล์ ออกมาเป็น String (ใช้ Class System.IO.File)
    /// </summary>
    /// <param name="textFilePath">ตำแหน่งของไฟล์ Text </param>
    /// <returns>String</returns>
    /// <remarks>https://msdn.microsoft.com/en-us/library/eked04a7.aspx</remarks>
        public static string ReadTextFileToStringV1(string textFilePath)
        {
            string RetStr = null;
            if (File.Exists(textFilePath) == false)
                Interaction.MsgBox("ไม่พบไฟล์");
            else
            {
                var sr = File.OpenText(textFilePath);
                while (sr.Peek() >= 0)
                    RetStr += sr.ReadLine();
                sr.Close();
            }
            return RetStr;
        }

        /// <summary>
    /// แปลงข้อความ เป็น รหัส UTF-16
    /// </summary>
    /// <param name="str">String</param>
    /// <returns>String</returns>
        public static string ConvertToUTF16(string str)
        {
            var ArrayOFBytes = Encoding.Unicode.GetBytes(str);
            string UTF16 = null;
            int v;
            var loopTo = ArrayOFBytes.Length - 1;
            for (v = 0; v <= loopTo; v++)
            {
                if (v % 2 == 0)
                {
                    int t = ArrayOFBytes[v];
                    ArrayOFBytes[v] = ArrayOFBytes[v + 1];
                    ArrayOFBytes[v + 1] = Convert.ToByte(t);
                }
            }

            var loopTo1 = ArrayOFBytes.Length - 1;
            for (v = 0; v <= loopTo1; v++)
            {
                string c = Conversion.Hex(ArrayOFBytes[v]);
                if (c.Length == 1)
                    c = "0" + c;
                UTF16 = UTF16 + c;
            }

            return UTF16;
        }

        /// <summary>
    /// http://stackoverflow.com/questions/18915633/determine-textfile-encoding
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static Encoding GetFileEncoding(string filePath)
        {
            using (var sr = new StreamReader(filePath, true))
            {
                sr.Read();
                return sr.CurrentEncoding;
            }
        }

        /// <summary>
        /// get ค่า Encoding ของข้อมูล
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    /// <remarks>
    /// เรียกใช้ใน VB
    /// Dim data() As Byte = File.ReadAllBytes("test.txt")
    /// Dim detectedEncoding As Encoding = DetectEncodingFromBom(Data)
    /// http://stackoverflow.com/questions/18915633/determine-textfile-encoding
    /// </remarks>
        public static Encoding DetectEncodingFromBom(byte[] data)
        {
            Encoding detectedEncoding = null;
            foreach (EncodingInfo info in Encoding.GetEncodings())
            {
                var currentEncoding = info.GetEncoding();
                var preamble = currentEncoding.GetPreamble();
                bool match = true;
                if (preamble.Length > 0 & preamble.Length <= data.Length)
                {
                    for (int i = 0, loopTo = preamble.Length - 1; i <= loopTo; i++)
                    {
                        if (preamble[i] != data[i])
                        {
                            match = false;
                            break;
                        }
                    }
                }
                else
                    match = false;
                if (match)
                {
                    detectedEncoding = currentEncoding;
                    break;
                }
            }
            return detectedEncoding;
        }

        public static DataTable ConvertToDataTable(string filePath, int numberOfColumns)
        {
            var tbl = new DataTable();

            for (int col = 0, loopTo = numberOfColumns - 1; col <= loopTo; col++)
                tbl.Columns.Add(new DataColumn("Column" + (col + 1).ToString()));


            var lines = File.ReadAllLines(filePath);

            foreach (string line in lines)
            {
                var cols = line.Split(':');

                var dr = tbl.NewRow();
                for (int cIndex = 0; cIndex <= 2; cIndex++)
                    dr[cIndex] = cols[cIndex];

                tbl.Rows.Add(dr);
            }

            return tbl;
        }
    }
}
