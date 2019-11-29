using System.Data;
using Microsoft.VisualBasic;
using System.Linq;
using System.Collections;
using System;
using System.IO;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{
    public static class CFncDataTable
    {

        /// <summary>
    /// Replace ข้อความ ใน คอลัมน์
    /// </summary>
    /// <param name="DatatableToReplace">DataTable ที่จำทำการ Replace</param>
    /// <param name="TargetString">ข้อความที่จะเปลี่ยน</param>
    /// <param name="ToString">ข้อความที่จะแทนที่</param>
    /// <param name="InColumnNumber">หมายเลขคอลัมน์ เริ่มจาก 0</param>
    /// <returns>DataTable</returns>
    /// <remarks></remarks>
        public static DataTable ReplaceIncolumn(DataTable DatatableToReplace, string TargetString, string ToString, int InColumnNumber)
        {
            // Dim dtRes As DataTable
            for (int nRow = 0, loopTo = DatatableToReplace.Rows.Count - 1; nRow <= loopTo; nRow++)
            {
                string NewString;
                NewString = DatatableToReplace.Rows[nRow][InColumnNumber].ToString().Replace(TargetString, ToString);
                DatatableToReplace.Rows[nRow][InColumnNumber] = NewString;
            }
            return DatatableToReplace;
        }

        /// <summary>
    /// Trim Cell ใน DataaTable
    /// </summary>
    /// <param name="DataTable">DataaTable ที่ต้องการ Trim</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static DataTable DatatableTrimCell(DataTable DataTable)
        {
            foreach (DataRow DRow in DataTable.Rows)
            {
                for (int nCol = 0, loopTo = DataTable.Columns.Count - 1; nCol <= loopTo; nCol++)
                    DRow[nCol] = Strings.Trim((string)DRow[nCol]);
            }
            return DataTable;
        }

        /// <summary>
    /// DataTable บันทึกเป้น txt ไฟล์ โดยระหว่าง คอลัมน์ จะใช้ Delimiter เป็นตัวแยกแยะ ข้อมูล
    /// / ทำงาน ด้วยการรัน เป็น Row - Column ลูปซ้อนลูป ตามลำดับ แบบละเอียด
    /// </summary>
        public static bool DataTableSaveToTxtFile(ref DataTable dTable, string Filename, string Delimiter = "~")
        {
            try
            {
                StreamWriter Stm;
                Stm = new StreamWriter(Filename, false, System.Text.Encoding.UTF8);
                string Str;
                foreach (DataRow DtRow in dTable.Rows)
                {
                    var StrInLine = new ArrayList();
                    for (int nCol = 0, loopTo = dTable.Columns.Count - 1; nCol <= loopTo; nCol++)
                        StrInLine.Add(DtRow[nCol].ToString());
                    Str = string.Join(Delimiter, StrInLine.ToArray(typeof(string)) as string[]);
                    Stm.WriteLine(Str);
                }
                Stm.Close();
                return true;
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.Message);
                return false;
            }
        }
        /// <summary>
    /// แปลง Array To DataTable
    /// </summary>
    /// <param name="ArrayLine"></param>
    /// <param name="Delimiter"></param>
    /// <returns>/C#</returns>
        public static DataTable ArrayToDataTable(string[] ArrayLine, string Delimiter = "~")
        {
            var dt = new DataTable();

            int TopDataCount = 0;
            foreach (string THisLine in ArrayLine) // My.Computer.FileSystem.ReadAllText(Filename).Split(Environment.NewLine)
            {
                var DataAry = Strings.Split(THisLine, Delimiter);
                // StringWatch &= DataAry
                dt.Rows.Add(DataAry);
                if (Information.IsArray(DataAry) == true)
                    TopDataCount = Conversions.ToInteger(Interaction.IIf(DataAry.Count() > TopDataCount, DataAry.Count(), TopDataCount));
            }

            return dt;
        }


        /// <summary>
        /// Convert DataTable To Array/620815
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns>Array Object Type /For Auto Int or string</returns>
        public static Object DataTableToArray(DataTable dt)
        {
            Object[,] DataArray = new Object[dt.Rows.Count,  dt.Columns.Count];
            if (dt.Rows.Count > 0)
            {
                for (int r = 0; r < dt.Rows.Count ; r++)
                {
                    for (int c = 0;c < dt.Columns.Count ; c++)
                    {
                        DataArray[r, c] = dt.Rows[r][c];
                    }
                }
            }
            return  DataArray;
        }


        /// <summary>
        /// เอา Column ที่ต้องการ *จากทุก Row* จาก DataTable ออกมา เป็น ลักษณะ Array
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="columnIndex">คอลัมน์ที่ต้องการ</param>
        /// <returns>string array ของ Column นั้น</returns>
        /// <remarks>/c#</remarks>
        public static string[] ColumnToStringArray(DataTable dataTable, int columnIndex) // 
        {
            var allAutoCompletes = from row in dataTable.AsEnumerable()
                                   let autoComplete = row.Field<string>(columnIndex)
                                   select autoComplete;
            return allAutoCompletes.ToArray(); // แปลงเป็น Array /Linq
        }
    }
}
