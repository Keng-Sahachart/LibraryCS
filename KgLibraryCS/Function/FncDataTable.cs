using System.Data;
using Microsoft.VisualBasic;
using System.Linq;
using System.Collections;
using System;
using System.IO;
using Microsoft.VisualBasic.CompilerServices;
using System.Windows.Forms;

namespace kgLibraryCs
{
    public static class FncDataTable
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
            Object[,] DataArray = new Object[dt.Rows.Count, dt.Columns.Count];
            if (dt.Rows.Count > 0)
            {
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        DataArray[r, c] = dt.Rows[r][c];
                    }
                }
            }
            return DataArray;
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


        //#region Save DataTable To Txt

        ///// <summary>

        /////     ''' นำ DataGrid มาบันทึกเป็นไฟล์ ด้วยคำสั่งของ DataGrid เอง ข้อเสีย คือ ไม่มีการแยกแยะ คอลัมน์

        /////     ''' </summary>
        //public void SaveGridData(ref DataGridView ThisGrid, string Filename, string Delimited = null)
        //{
        //    ThisGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        //    ThisGrid.SelectAll();
        //    System.IO.File.WriteAllText(Filename, ThisGrid.GetClipboardContent().GetText().TrimEnd());
        //    ThisGrid.ClearSelection();
        //}

        ///// <summary>
        /////     ''' DataTable บันทึกเป้น txt ไฟล์ โดยระหว่าง คอลัมน์ จะใช้ Delimiter เป็นตัวแยกแยะ ข้อมูล 
        /////     ''' / ทำงานด้วย การ รัน Row และ Join
        /////     ''' </summary>
        //public Boolean DataTableSaveToTxtFile1(ref DataTable DataTable, string FileNameSave, string Delimiter = "~")
        //{
        //    try
        //    {
        //        using (StreamWriter StmW = new StreamWriter(FileNameSave, false, System.Text.Encoding.UTF8))
        //        {
        //            foreach (DataRow drow in DataTable.Rows)
        //            {
        //                var lineoftext = string.Join(Delimiter, drow.ItemArray.Select(s => s.ToString()).ToArray());

        //                StmW.WriteLine(lineoftext);
        //            }
        //            StmW.Close();
        //        }
        //        return true;
        //    }
        //    catch //(Exception ex)
        //    {
        //        // MsgBox(ex.Message)
        //        return false;
        //    }
        //}

        //#endregion


        //#region Load Txt to DataGrid
        ///// <summary>

        /////     ''' โหลด ไฟล์ ใส่ DataGrid

        /////     ''' </summary>
        //public void LoadGridData(ref DataGridView ThisGrid, string Filename, string Delimiter = "~")
        //{
        //    DataTable dt = new DataTable();
        //    // 1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
        //    for (var nC = 0; nC <= 100; nC++)
        //        dt.Columns.Add("Col" + nC);
        //    // 2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
        //    // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
        //    int TopDataCount = 0;

        //    string[] allLine = File.ReadAllText(Filename).Split(new[] { Environment.NewLine },StringSplitOptions.None);
        //    foreach (var THisLine in allLine)
        //    {
        //        var DataAry = Strings.Split(THisLine, Delimiter);
        //        dt.Rows.Add(DataAry);
        //        TopDataCount =(int) Interaction.IIf(DataAry.Count() > TopDataCount, DataAry.Count(), TopDataCount);
        //    }
        //    // 3.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป
        //    for (var rowBack = (dt.Columns.Count - 1); rowBack >= TopDataCount; rowBack += -1)
        //        dt.Columns.Remove(dt.Columns[rowBack]);

        //    ThisGrid.Rows.Clear();
        //    ThisGrid.DataSource = dt;
        //}

        ///// <summary>
        /////     ''' โหลดไฟล์ ใส่ DataGrid
        /////     ''' </summary>
        //public void LoadGridData2(ref DataGridView ThisGrid, string Filename, char SplitChar)
        //{
        //    System.IO.StreamReader file = new System.IO.StreamReader(Filename);
        //    DataTable dt = new DataTable();
        //    // 1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
        //    for (var nC = 0; nC <= 100; nC++)
        //        dt.Columns.Add("Col" + nC);
        //    // 2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
        //    // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
        //    int TopDataCount = 0;
        //    //string newline = "";
        //    string[] allLine = File.ReadAllText(Filename).Split(new[] { Environment.NewLine }, StringSplitOptions.None);
        //    foreach (var newline in allLine)
        //    {
        //        DataRow dr = dt.NewRow();
        //        string[] values = newline.Split(SplitChar); // (Microsoft.VisualBasic.ChrW(32))
        //        dr.ItemArray = values;
        //        dt.Rows.Add(dr);
        //        TopDataCount =(int) Interaction.IIf(values.Count() > TopDataCount, values.Count(), TopDataCount);
        //    }
        //    // 3.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป
        //    for (var rowBack = (dt.Columns.Count - 1); rowBack >= TopDataCount; rowBack += -1)
        //        dt.Columns.Remove(dt.Columns[rowBack]);
        //    file.Close();
        //    ThisGrid.DataSource = dt;
        //}
        ///// <summary>
        /////     ''' โหลดไฟล์ เป็น DataTable
        /////     ''' </summary>
        //public DataTable LoadTxtToDataTable(string Filename, string Delimiter = "~")
        //{
        //    DataTable dt = new DataTable();
        //    // 1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
        //    for (var nC = 0; nC <= 100; nC++)
        //        dt.Columns.Add("Col" + nC);
        //    // DataTable_TextFile.DataTableSaveToTxtFile1(dt, "D:\LoadTxtToDataTable_ColCreate.txt", "")
        //    // 2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
        //    // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
        //    //string StringWatch = "";
        //    int TopDataCount = 0;
        //    string[] lines = System.IO.File.ReadAllLines(Filename);
        //    foreach (var THisLine in lines) // My.Computer.FileSystem.ReadAllText(Filename).Split(Environment.NewLine)
        //    {
        //        var DataAry = Strings.Split(THisLine, Delimiter);
        //        // StringWatch &= DataAry
        //        dt.Rows.Add(DataAry);
        //        if (Information.IsArray(DataAry) == true)
        //            TopDataCount = (int) Interaction.IIf(DataAry.Count() > TopDataCount, DataAry.Count(), TopDataCount);
        //    }
        //    // DataTable_TextFile.DataTableSaveToTxtFile1(dt, "D:\LoadTxtToDataTable_Convtd.txt", "")
        //    // 3.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป
        //    for (var rowBack = (dt.Columns.Count - 1); rowBack >= TopDataCount; rowBack += -1)
        //        dt.Columns.Remove(dt.Columns[rowBack]);
        //    return dt;
        //}
        //#endregion

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
    }
}
