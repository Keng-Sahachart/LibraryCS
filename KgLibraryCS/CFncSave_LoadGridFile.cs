using System.Data;
using Microsoft.VisualBasic;
using System.Linq;
using System;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{
    public static class CFncSave_LoadGridFile
    {
        // Dim ThisFilename As String = Application.StartupPath & "\MyData.dat"
        /// <summary>
    /// นำ DataGrid มาบันทึกเป็นไฟล์ ด้วยคำสั่งของ DataGrid เอง ข้อเสีย คือ ไม่มีการแยกแยะ คอลัมน์
    /// </summary>
        public static void SaveGridData(ref DataGridView ThisGrid, string Filename, string Delimited = null)
        {
            ThisGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            ThisGrid.SelectAll();
            File.WriteAllText(Filename, ThisGrid.GetClipboardContent().GetText().TrimEnd());
            ThisGrid.ClearSelection();
        }

        /// <summary>
    /// DataTable บันทึกเป้น txt ไฟล์ โดยระหว่าง คอลัมน์ จะใช้ Delimiter เป็นตัวแยกแยะ ข้อมูล
    /// / ทำงานด้วย การ รัน Row และ Join
    /// </summary>
        public static object DataTableSaveToTxtFile1(ref DataTable DataTable, string FileNameSave, string Delimiter = "~")
        {
            try
            {
                using (var StmW = new StreamWriter(FileNameSave, false, Encoding.UTF8))
                {
                    foreach (DataRow drow in DataTable.Rows)
                    {
                        string lineoftext = string.Join(Delimiter, drow.ItemArray.Select(s => s.ToString()).ToArray());

                        StmW.WriteLine(lineoftext);
                    }
                    StmW.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                // MsgBox(ex.Message)
                return false;
            }
        }




        /// <summary>
    /// โหลด ไฟล์ ใส่ DataGrid
    /// </summary>
        public static void LoadGridData(ref DataGridView ThisGrid, string Filename, string Delimiter = "~")
        {
            var dt = new DataTable();
            // 1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
            for (int nC = 0; nC <= 100; nC++)
                dt.Columns.Add("Col" + Conversions.ToString(nC));
            // 2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
            // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
            int TopDataCount = 0;

            foreach (var THisLine in System.IO.File.ReadAllText(Filename).Split(Conversions.ToChar(Environment.NewLine)))
            {
                var DataAry = Strings.Split(THisLine, Delimiter);
                dt.Rows.Add(DataAry);
                TopDataCount = Conversions.ToInteger(Interaction.IIf(DataAry.Count() > TopDataCount, DataAry.Count(), TopDataCount));
            }
            // 3.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป
            for (int rowBack = dt.Columns.Count - 1, loopTo = TopDataCount; rowBack >= loopTo; rowBack += -1)
                dt.Columns.Remove(dt.Columns[rowBack]);
            ThisGrid.Rows.Clear();
            ThisGrid.DataSource = dt;
        }

        /// <summary>
    /// โหลดไฟล์ ใส่ DataGrid
    /// </summary>
        public static void LoadGridData2(ref DataGridView ThisGrid, string Filename, char SplitChar)
        {
            var file = new StreamReader(Filename);
            var dt = new DataTable();
            // 1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
            for (int nC = 0; nC <= 100; nC++)
                dt.Columns.Add("Col" + Conversions.ToString(nC));
            // 2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
            // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
            int TopDataCount = 0;
            
            foreach (string newline in System.IO.File.ReadAllText(Filename).Split(Conversions.ToChar(Environment.NewLine)))
            {
                var dr = dt.NewRow();
                var values = newline.Split(SplitChar); // (Microsoft.VisualBasic.ChrW(32))
                dr.ItemArray = values;
                dt.Rows.Add(dr);
                TopDataCount = Conversions.ToInteger(Interaction.IIf(values.Count() > TopDataCount, values.Count(), TopDataCount));
            }
            // 3.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป
            for (int rowBack = dt.Columns.Count - 1, loopTo = TopDataCount; rowBack >= loopTo; rowBack += -1)
                dt.Columns.Remove(dt.Columns[rowBack]);
            file.Close();
            ThisGrid.DataSource = dt;
        }
        /// <summary>
    /// โหลดไฟล์ เป็น DataTable
    /// </summary>
        public static DataTable LoadTxtToDataTable(string Filename, string Delimiter = "~")
        {
            var dt = new DataTable();
            // 1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
            for (int nC = 0; nC <= 100; nC++)
                dt.Columns.Add("Col" + Conversions.ToString(nC));
            // 2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
            // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
            //string StringWatch = "";
            int TopDataCount = 0;
            var lines = File.ReadAllLines(Filename);
            foreach (var THisLine in lines)
            {
                var DataAry = Strings.Split(THisLine, Delimiter);
                // StringWatch &= DataAry
                dt.Rows.Add(DataAry);
                if (Information.IsArray(DataAry) == true)
                    TopDataCount = Conversions.ToInteger(Interaction.IIf(DataAry.Count() > TopDataCount, DataAry.Count(), TopDataCount));
            }
            // 3.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป
            for (int rowBack = dt.Columns.Count - 1, loopTo = TopDataCount; rowBack >= loopTo; rowBack += -1)
                dt.Columns.Remove(dt.Columns[rowBack]);
            return dt;
        }


        public static DataTable Trim_DataTable(DataTable dt)
        {
            foreach (DataRow Row in dt.Rows)
            {
                foreach (DataColumn Col in dt.Columns)
                {
                    if (Col.DataType == Type.GetType("System.String"))
                    {
                        if (Information.IsDBNull(Row[Col.ColumnName]))
                            Row[Col.ColumnName] = "";
                        else
                            Row[Col.ColumnName] = Row[Col.ColumnName].ToString().Trim();
                    }
                }
            }
            return dt;
        }


        public static object Trim_DataTableAndCleanSpecialChar_DataTable(DataTable dt, bool WantTrim = true, bool WantClean = true)
        {
            foreach (DataRow Row in dt.Rows)
            {
                foreach (DataColumn Col in dt.Columns)
                {
                    if (Col.DataType == Type.GetType("System.String"))
                    {
                        if (Information.IsDBNull(Row[Col.ColumnName]))
                            Row[Col.ColumnName] = "";
                        else
                        {
                            if (WantTrim == true)
                                Row[Col.ColumnName] = Row[Col.ColumnName].ToString();
                            if (WantClean == true)
                                // Row.Item[Col.ColumnName] = RemoveSpecialCharacters(Row.Item[Col.ColumnName])
                                // ----------------------------
                                // Row.Item[Col.ColumnName] = Regex.Replace(Row.Item(Col.ColumnName], "[^A-Za-z0-9ก-๙\-/]", " ")
                                // Row.Item[Col.ColumnName] = Regex.Replace(Row.Item(Col.ColumnName], "[;\/:*?""<>|&']", "")
                                // ----------------------------
                                // ล้าง ตัว *อักษรพิเศษ*  ที่ไม่ได้อยู่ในแป้นพิมพ์ ปกติ / เอาเฉพาะตัวอักษร ที่อยู่ในตาราง ASCII ทั้งหมด 
                                Row[Col.ColumnName] = Regex.Replace(Row[Col.ColumnName].ToString(), "[^!-๙]", " ");
                        }
                    }
                }
            }
            return dt;
        }

        public static string RemoveSpecialCharacters(string str)
        {
            var sb = new StringBuilder();
            foreach (char c in str)
            {
                if (c >= '0' && c <= '9' || c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' || c >= 'ก' && c <= 'ฮ' || c >= 'ฯ' && c <= '๙' || c == '.' || c == '?' || c == '!' || c == '_')
                    sb.Append(c);
                else
                    sb.Append(" ");
            }
            return sb.ToString();
        }
    }
}
