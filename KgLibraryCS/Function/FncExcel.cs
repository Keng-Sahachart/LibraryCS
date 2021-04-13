using System.Data;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel; // add reference
using System.Data.OleDb;
using VBIDE = Microsoft.Vbe.Interop; // สำหรับ ใช้ทำงานกับ VB script หรือ Macro ใน Excel 
using System.IO;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.Win32;
using Microsoft.Vbe;//.VBIDE;
using Microsoft.Vbe.Interop;


namespace kgLibraryCs
{
    //  ExcelTools = Microsoft.Office.Tools.Excel
    //  ExcelTools9 = Microsoft.Office.Tools.Excel.v9.0
    // //////////////////////////////////////////////////////////////////////////
    // Excel.XlLineStyle
    // wSheet.Range(strRangeTableBody).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
    // //////////////////////////////////////////////////////////////////////////
    // System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
    // ** จะกระทำการใดๆ ใน Sheet ต้อง Active ก่อน ถ้าเป็น Sheet ใหม่
    // ** ไม่ยอมให้ทำงานกับ Sheet ที่ไม่ได้ Active จะฟ้อง Error ทันที
    // #############################################################
    // การเรียก Range หรือ  Cell
    // wSheet.Cells(10, "A")

    public static class FncExcel
    {

        // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        /// <summary>ทำให้ font เป็นปรับตามขนาด Cell</summary>
    /// <param name="wSheet">ตัวแปร WorkSheet ที่ต้องการจะแก้ไข</param>
    /// <param name="Range">กำหนด Range A1:B2</param>
    /// <param name="TrueOrFalse">กำหนด สถานะ ใช่หรือไม่ใช่</param>
        public static void Set_FontSizeAutoFit(ref Excel.Worksheet wSheet, object Range, bool TrueOrFalse = true)
        {
            wSheet.get_Range(Range).ShrinkToFit = TrueOrFalse;
        }

        public static void Do_Merge(ref Excel._Worksheet wSheet, object Range)
        {
            wSheet.Activate();

            wSheet.get_Range(Range).Merge();
        }

        public static void Set_Alignment_Vertical(ref Excel._Worksheet wsheet, object Range, Excel.XlVAlign Align)
        {
            wsheet.Activate();
            wsheet.get_Range(Range).VerticalAlignment = Align;
        }

        public static void Set_Alignment_Horizontal(ref Excel._Worksheet wsheet, object Range, Excel.XlVAlign Align)
        {
            wsheet.Activate();
            wsheet.get_Range(Range).HorizontalAlignment = Align;
        }
        public static void Sheet_SetBorderCellAround(ref Excel._Worksheet wSheet, object Range, Excel.XlLineStyle XlLineStyle)
        {
            wSheet.Activate();
            wSheet.get_Range(Range).BorderAround(XlLineStyle);
        }

        public static void Sheet_SetBorderCell(ref Excel._Worksheet wSheet, object Range, Excel.XlLineStyle XlLineStyle
                                , Excel.XlBordersIndex XlBorderIndex = default(Excel.XlBordersIndex))
        {
            wSheet.Activate();
            if (XlBorderIndex == default(int))
                wSheet.get_Range(Range).Borders.LineStyle = XlLineStyle;
            else
                wSheet.get_Range(Range).Borders[XlBorderIndex].LineStyle = XlLineStyle;
        }

        /// <summary> ใส่สีพื้นหลังให้ Cell ใน Range ที่ต้องการ </summary>
    /// <param name="wSheet">ตัวแปร WorkSheet ที่ต้องการจะนำมาใส่สีพื้นหลังให้ Cell</param>
    /// <param name="Range">Range ของ Cell ที่ต้องการใส่สีพื้นหลัง</param>
    /// <param name="Red">ค่า  0-255 </param>
    /// <param name="Green">ค่า  0-255 </param>
    /// <param name="Blue">ค่า  0-255 </param>
        public static void Sheet_BackGroundColor(ref Excel._Worksheet wSheet, string Range, int Red, int Green, int Blue)
        {
            wSheet.Activate();
            wSheet.get_Range(Range).Interior.Color = Information.RGB(Red, Green, Blue);
        }


        /// <summary>แปลงหมายเลข Column ให้กลายชื่อ Column A-Z </summary>
    /// <param name="index">หมายเลข Column</param>
        public static string NumberToColumnName(int index)
        {
            string NumberToColumnNameRet = default(string);

            char[] chars = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H' , 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'};

            index -= 1; // '//adjust so it matches 0-indexed array rather than 1-indexed column
            int quotient = index / 26; // '//normal / operator rounds. \ does integer division, which truncates
            if (quotient > 0)
                NumberToColumnNameRet = NumberToColumnName(quotient) + Conversions.ToString(chars[index % 26]);
            else
                NumberToColumnNameRet = Conversions.ToString(chars[index % 26]);
            return NumberToColumnNameRet;
        }

        public static string ConvertToLetter(int iCol)
        {
            string RetVal = "";
            int iAlpha;
            int iRemainder;
            iAlpha = Conversions.ToInteger(Conversion.Int(iCol / (double)27));
            iRemainder = iCol - iAlpha * 26;
            if (iAlpha > 0)
                RetVal = Conversions.ToString((char)(iAlpha + 64));
            if (iRemainder > 0)
                RetVal = RetVal + Conversions.ToString((char)(iRemainder + 64));
            return RetVal;
        }

        /// <summary>หาหมายเลขของ Row สุดท้ายที่ถูกใช้งาน</summary>
    /// <param name="wSheet"> WorkSheet ที่ต้องการทราบ Row สุดท้ายที่ถูกใช้งาน</param>
        public static int SheetLastRow(Excel._Worksheet wSheet)
        {
            wSheet.Activate();
           
            var rng = wSheet.UsedRange;
            return wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; // xlCellTypeLastCell).Row()
        }

        public static void FreezePane(ref Excel._Worksheet wSheet, string RangeSelect)
        {
            wSheet.get_Range(RangeSelect).Application.ActiveWindow.FreezePanes = true;
        }

        /// <summary>หาหมายเลขของ Column สุดท้ายที่ถูกใช้งาน</summary>
    /// <param name="wSheet"> WorkSheet ที่ต้องการทราบ Column สุดท้ายที่ถูกใช้งาน</param>
        public static int SheetLastColumn(Excel._Worksheet wSheet)
        {
            wSheet.Activate();
            // Return wSheet.UsedRange.Columns.Count()
            return wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
        }

        public static void SheetAutoFill_From1Cell(ref Excel._Worksheet wSheet, string ValueFirstCell, string Range)
        {
            // create formula
            wSheet.Activate();
            string FirstCell = Range.Split(':')[0];
            wSheet.get_Range(FirstCell).Formula = ValueFirstCell;

            // copy this formula down using autofill
            // wSheet.Range(RangeName).AutoFill(wSheet.Range(FormulaFirstCell))
            if ((FirstCell ?? "") != (Range.Split(':')[1] ?? ""))
                wSheet.get_Range(FirstCell).AutoFill(wSheet.get_Range(Range));
        }

        public static void DeleteRow(ref Excel.Worksheet WorkSheet, int RowIndex)
        {
            WorkSheet.Rows[RowIndex].Delete();
        }


        /// <summary>
        /// บันทึกไฟล์ Excel เฉพาะ .XLS และ .XLSX เท่านั้น
        /// </summary>
        /// <param name="xlApp">Excel.Application</param>
        /// <param name="BoolSaveToXlsx">ใช่คือ XLSX ไม่ใช่คือ XLS</param>
        /// <param name="FilePathToSaveWithOutExt">ตำแหน่งเซฟไฟล์ ไม่ต้องการ นามสกุลไฟล์</param>
        /// <returns>ส่งออกมาเป็นตำแหน่งไฟล์ เต็ม</returns>
        /// <remarks></remarks>
        /// จำได้ว่า เวอร์ชั่นที่ต่างกัน ซึ่งจะมีค่าพารามิเตอร์ที่ ไม่เหมือนกันในบางเวอร์ชั่น ตอนนี้ เอาออกก่อน หาเว็บ อ้างอิงไม่เจอ
        public static string SaveAsWorkBookByVersion(Excel.Application XlApp, Boolean BoolSaveToXlsX, string FilePathToSaveWithOutExt)
        {
            bool DefaultDisplayAlerts = XlApp.DisplayAlerts;

            Excel.XlFileFormat EnumFileFormat = Excel.XlFileFormat.xlExcel8;//int EnumFileFormat = 56;
            string FileFormat = "xls";
            string XlsVer = XlApp.Version;

            if (BoolSaveToXlsX == true)
            {
                EnumFileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook; // 51
                FileFormat = "xlsx";//
            }
            else
            {
                EnumFileFormat = Excel.XlFileFormat.xlExcel8;
                FileFormat = "xls";//
            }

            // Check File Extension
            String Path = "", FileWithOutExt = "", FileExtension = "", FileNameSaveFull = "";
            FileExtension = System.IO.Path.GetExtension(FilePathToSaveWithOutExt);
            if (FileExtension != "")
            { // ป้องกัน ถ้ามี ใส่ Extension ติดเข้ามาด้วย
                Path = System.IO.Path.GetDirectoryName(FilePathToSaveWithOutExt);
                FileWithOutExt = System.IO.Path.GetFileNameWithoutExtension(FilePathToSaveWithOutExt);
                FileNameSaveFull = System.IO.Path.ChangeExtension(FilePathToSaveWithOutExt, FileFormat);
            }
            else
            {
                FileNameSaveFull = FilePathToSaveWithOutExt + "." + FileFormat;
            }

            FileNameSaveFull = FncFileFolder.NewFileNameUnique(FileNameSaveFull);//สร้างชื่อไฟล์ใหม่ หาก มีไฟล์ชื่อเดิมอยู่แล้ว

            XlApp.DisplayAlerts = false;
            XlApp.ActiveWorkbook.SaveAs(FileNameSaveFull, EnumFileFormat);
            XlApp.DisplayAlerts = DefaultDisplayAlerts;

            return FileNameSaveFull;
        }

        public static string WbookSaveAsXls(ref Excel.Application xlApp, string FilePathToSaveWithOutExt)
        {
            // เลือกไฟล์ ฟอร์แมต สำหรับ ใช้ 2003 กับ 2007 
            string FileFormat = "xls";
            int XlsVer = Conversions.ToInteger(xlApp.Version);

            string FileNameSaveFull;
            FileNameSaveFull = FilePathToSaveWithOutExt + "." + FileFormat;
            FileNameSaveFull = FncFileFolder.NewFileNameUnique(FileNameSaveFull); // สร้างชื่อไฟล์ใหม่ หาก มีไฟล์ชื่อเดิมอยู่แล้ว

            if (XlsVer == 11)
                xlApp.ActiveWorkbook.SaveAs(FileNameSaveFull, 43);
            else
                xlApp.ActiveWorkbook.SaveAs(FileNameSaveFull, 56);

            return FileNameSaveFull;
        }

        /// <summary>
    /// เซ็ตทุกชีท ใน WorkBook ให้มาอยู่ที่ Range ที่กำหนด
    /// </summary>
    /// <param name="wBook"></param>
    /// <param name="StrRange"></param>
    /// <remarks></remarks>
        public static void SetCursurAllSheetInWBook(ref Excel.Workbook wBook, string StrRange)
        {
            foreach (Excel._Worksheet wSht in wBook.Worksheets)
            {
                if ((int)wSht.Visible == (int)Excel.XlSheetVisibility.xlSheetVisible)
                {
                    wSht.Activate();
                    wSht.get_Range(StrRange).Select();
                }
            }
        }
        // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        // @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

        /// <summary>Copy และ Paste Column ภายใน WorkSheet โดย Column ที่ถูกวางจากถูกแทนที่ด้วยข้อมูลที่ Copy มา</summary>
    /// <param name="wSheet"> WorkSheet</param>
    /// <param name="SourceNumberToColumnName"> ชื่อ Column ที่ต้องการ Copy</param>
    /// <param name="DestinationNumberToColumnName"> ชื่อ Column ที่ต้องการ Paste ข้อมูลทับ</param>
        public static void ExcelCopyColumn(ref Excel._Worksheet wSheet, string SourceNumberToColumnName, string DestinationNumberToColumnName)
        {
            wSheet.Activate();
            wSheet.Columns[SourceNumberToColumnName].copy();
            wSheet.Columns[DestinationNumberToColumnName].Select();
            wSheet.Paste();
        }
        public static void ExcelCopyColumnSheetToSheet(ref Excel._Worksheet wSheetCopy, string SourceNumberToColumnName, ref Excel._Worksheet wSheetPast, string DestinationNumberToColumnName)
        {
            wSheetCopy.Activate();
            wSheetCopy.Columns[SourceNumberToColumnName].copy();
            wSheetPast.Activate();
            wSheetPast.Columns[DestinationNumberToColumnName].Select();
            wSheetPast.Paste();
        }
        public static void ExcelCopyRow(ref Excel._Worksheet wSheet, int SourceRowNumber, string TargetrRowNumber)
        {
            wSheet.Activate();
            wSheet.Rows[SourceRowNumber].copy();
            wSheet.Rows[TargetrRowNumber].Select();
            wSheet.Rows[TargetrRowNumber].Insert(Shift: Excel.XlDirection.xlDown);    // Insert multiple copied rows
            wSheet.Paste();
        }
        public static Excel.Worksheet CopySheet(ref Excel.Workbook wBook, int numSheetToCopy, int numSheetDestination)
        {
            var wShtRet = new Excel.Worksheet();
            try
            {
                wBook.Sheets[numSheetToCopy].Copy(Before: wBook.Sheets[numSheetDestination]);
            }
            // Return wShtRet 'wBook.Worksheets(numSheetDestination)
            catch //(Exception ex)
            {
                wBook.Sheets[numSheetToCopy].Copy(after: wBook.Sheets[wBook.Sheets.Count]);
            }
            wShtRet = wBook.ActiveSheet;
            return wShtRet;
        }

        public static Excel.Worksheet CopySheet(ref Excel.Workbook wBook, String shtName, int numSheetDestination = 0)
        {
            var wShtRet = new Excel.Worksheet();
            try
            {
                wBook.Sheets[shtName].Copy(Before: wBook.Sheets[numSheetDestination]);
                
            }
            // Return wShtRet 'wBook.Worksheets(numSheetDestination)
            catch //(Exception ex)
            {
                wBook.Sheets[shtName].Copy(after: wBook.Sheets[wBook.Sheets.Count]);
            }
            wShtRet = wBook.ActiveSheet;
            return wShtRet;
        }
        /// <summary>
    /// Copy ทุกชีทจาก WorkBook แรก ไปอยู่อีก WorkBook ที่สอง
    /// </summary>
        public static void CopyWorkBookToWorkBook(ref Excel._Workbook wBookSource, ref Excel.Workbook wBookDestination)
        {
            wBookSource.Activate();
            for (int nSht = 1, loopTo = wBookSource.Worksheets.Count; nSht <= loopTo; nSht++)
                wBookSource.Sheets[nSht].Copy(Before: wBookDestination.Sheets[nSht]);
        }
        public static void Sort2Column(ref Excel._Worksheet wSheet, string NumberToColumnName1, string NumberToColumnName2)
        {
            wSheet.Activate();

            wSheet.Columns.Sort(Key1: wSheet.Columns[NumberToColumnName1], Order1: Excel.XlSortOrder.xlAscending, Key2: wSheet.Columns[NumberToColumnName2], Order2: Excel.XlSortOrder.xlAscending, Orientation: Excel.XlSortOrientation.xlSortColumns, Header: Excel.XlYesNoGuess.xlNo, SortMethod: Excel.XlSortMethod.xlPinYin, DataOption1: Excel.XlSortDataOption.xlSortNormal, DataOption2: Excel.XlSortDataOption.xlSortNormal, DataOption3: Excel.XlSortDataOption.xlSortNormal);
        }
        // 
        /// <summary> ConvertTextToColumn(SheetOpen, "A", "~") </summary>
        public static void ConvertTextToColumn(ref Excel._Worksheet wSheet, string DataColumn, string SymbolExplode)
        {
            wSheet.Activate();
           
            wSheet.Columns[DataColumn].TextToColumns(DataType: Excel.XlTextParsingType.xlDelimited, TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone, ConsecutiveDelimiter: true, Other: true, OtherChar: SymbolExplode, TrailingMinusNumbers: true);
        }

    //    public static void ConvertTextToColumnFormatTest(ref Excel.Worksheet wSheet, string DataColumn, string SymbolExplode)
    //    {
    //        wSheet.Activate();
    //        // wSheet.Columns(DataColumn).TextToColumns(DataType:=Excel.XlTextParsingType.xlDelimited, _
    //        // ConsecutiveDelimiter:=True, _
    //        // Other:=True, _
    //        // OtherChar:=SymbolExplode, _
    //        // TrailingMinusNumbers:=True)
    //        wSheet.Columns[DataColumn].TextToColumns(DataType: Excel.XlTextParsingType.xlDelimited, TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone, ConsecutiveDelimiter: true, Other: true, OtherChar: SymbolExplode, TrailingMinusNumbers: true
    //, FieldInfo: new[] { { 1, 2 }, { 2, 2 } });
    //    }

        public static void wSheetInsertRow(ref Excel._Worksheet wSheet, int nRow, int ManyRows = 1)
        {
            wSheet.Activate();
            for (int i = 1, loopTo = ManyRows; i <= loopTo; i++)
                wSheet.Rows[nRow].Insert(Shift: Excel.XlDirection.xlDown);
        }
        public static void Set_FormatCell(ref Excel._Worksheet wSheet, string Range, string StringFormat)
        {
            wSheet.Activate();
            wSheet.get_Range(Range).NumberFormat = StringFormat;
        }
        public static void wSheetNumberFormatInComma(ref Excel._Worksheet wSheet, string Range)
        {
            wSheet.Activate();
            wSheet.get_Range(Range).NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";
        }


        public static void wSheetReplaceAll(ref Excel._Worksheet wSheet, string ToFind, string ToReplace)
        {
            wSheet.Activate();
            wSheet.Cells.Replace(ToFind, ToReplace, 2, 1, false, false, false, false);
        }
        /// <summary>Replace คำ กำหนด Range ได้</summary>
        public static void wSheetReplaceInRANGe(ref Excel._Worksheet wSheet, string ToFind, string ToReplace, string Range = null)
        {
            wSheet.Activate();
            if (Range != null)
                wSheet.get_Range(Range).Replace(ToFind, ToReplace);
            else
                wSheet.UsedRange.Replace(ToFind, ToReplace, 2, 1, false, false, false, false);
        }
        public static void SheetDestroySpaceBarCell(ref Excel._Worksheet wSheet) // เร็ว 1 จากการเทียบ กับ ฟังชั่นอื่น รู้สึกว่า การ Count จะทำให้นาน
        {
            wSheet.Activate(); // แต่ฟังชั่นนี้ เร็วเพราะ ใช้วิธีการ Replace ช่องว่าง ให้หมด แล้วนับ ว่าเป็น 0 หรือไม่ ต่างกับฟังชั้นที่ใช้การ นับช่องว่าง
            int LastRow = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int LastColumn = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            string CellVal;
            string CellLen;
            for (int y = 1, loopTo = LastRow - 1; y <= loopTo; y++)
            {
                for (int x = 1, loopTo1 = LastColumn - 1; x <= loopTo1; x++)
                {
                    CellVal = wSheet.Cells[y, x].value;
                    CellLen = Conversions.ToString(Strings.Len(CellVal));
                    if (CellVal != null)
                    {
                        CellVal = CellVal.Replace(" ", ""); // ปัญหา: ถ้าเจอ Nothing แล้ว จะ Error ทันที่ แก้ไข ให้อยู่ใน if
                        if (Strings.Len(CellVal) == 0)
                            wSheet.Cells[y, x].value = null;
                    }
                }
            }
        }

        /// <summary> ทำให้ Cell ที่มีแต่ค่า OneChar1,2,3,4,5 เท่านั้น เป็นช่องว่างไปเลย  รับได้ 5 อักษร คำ ที่ไม่ต้องการ   </summary>
    /// <param name="wSheet">ตัวแปร WorkSheet ที่ต้องการจะนำมาลบตัวอักษรที่ต้องการลบ</param>
    /// <param name="OneChar1">ตัวอักษรทั้ง 5 ตัว ถ้ามีเฉพาะ 5 ตัวนี้ใน Cell จะโดนลบออกไปจาก Sheet</param>
    /// <remarks>อะรูไม่ไร้</remarks>
        public static void SheetDestroyCharCell(ref Excel._Worksheet wSheet, string OneChar1, string OneChar2 = null, string OneChar3 = null, string OneChar4 = null, string OneChar5 = null) // เร็ว 1 จากการเทียบ กับ ฟังชั่นอื่น รู้สึกว่า การ Count จะทำให้นาน
        {
            wSheet.Activate(); // แต่ฟังชั่นนี้ เร็วเพราะ ใช้วิธีการ Replace ช่องว่าง ให้หมด แล้วนับ ว่าเป็น 0 หรือไม่ ต่างกับฟังชั้นที่ใช้การ นับช่องว่าง
            int LastRow = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int LastColumn = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            for (int y = 1, loopTo = LastRow - 1; y <= loopTo; y++)
            {
                string CellVal;
                string CellLen;
                for (int x = 1, loopTo1 = LastColumn - 1; x <= loopTo1; x++)
                {
                    // If y = 83 Then
                    // y = y  'debug
                    // End If
                    CellVal = wSheet.Cells[y, x].value;
                    CellLen = Conversions.ToString(Strings.Len(CellVal));
                    if (CellVal != null)
                    {
                        CellVal = CellVal.Replace(OneChar1, "");
                        if (OneChar2 != null)
                        {
                            CellVal = CellVal.Replace(OneChar2, "");
                            if (OneChar3 != null)
                            {
                                CellVal = CellVal.Replace(OneChar3, ""); // ปัญหา: ถ้าเจอ Nothing แล้ว จะ Error ทันที่ แก้ไข ให้อยู่ใน if
                                if (OneChar4 != null)
                                {
                                    CellVal = CellVal.Replace(OneChar4, ""); // ปัญหา: ถ้าเจอ Nothing แล้ว จะ Error ทันที่ แก้ไข ให้อยู่ใน if
                                    if (OneChar5 != null)
                                        CellVal = CellVal.Replace(OneChar5, "");// ปัญหา: ถ้าเจอ Nothing แล้ว จะ Error ทันที่ แก้ไข ให้อยู่ใน if
                                }
                            }
                        }
                        if (Strings.Len(CellVal) == 0)
                            wSheet.Cells[y, x].value = null;
                    }
                }
            }
        }

        /// <summary> ลบ ช่องว่าง ทางซ้ายและทางขวา ของข้อความออก </summary>
        public static void SheetTrimCell(ref Excel._Worksheet wSheet)
        {
            wSheet.Activate();
            int LastRow = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int LastColumn = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            string CellVal;
            for (int y = 1, loopTo = LastRow - 1; y <= loopTo; y++)
            {
                for (int x = 1, loopTo1 = LastColumn - 1; x <= loopTo1; x++)
                {
                    CellVal = wSheet.Cells[y, x].value;
                    if (CellVal != null)
                        wSheet.Cells[y, x].value = Strings.Trim(CellVal);
                }
            }
        }

        /// <summary> ลบ ช่องว่าง ทางซ้ายและทางขวา ของข้อความออก </summary>
        public static void SheetTrimCellQuick(ref Excel._Worksheet wSheet, string Range = null)
        {
            wSheet.Activate();
            Excel.Range SheetRang;
            if (Range == null)
                SheetRang = wSheet.UsedRange;
            else
                SheetRang = wSheet.get_Range(Range);
            foreach (Excel.Range cell in SheetRang)
                cell.Value = Strings.Trim(cell.Value);
        }

        /// <summary> Copy ค่าที่อยู่ ใน Cell เหนือ Cell ที่มีช่องว่าง มาใส่ในช่องว่าง </summary>
        public static void SheetColumnCopyFall(ref Excel._Worksheet wSheet, int ColumnNumber)
        {
            // ก่อนใช้ฟังชั่นนี้อย่าลืม ใช้คำสั่ง Trim ก่อน ไม่ใช่นั้น อาจะไม่ได้ผล
            wSheet.Activate();
            string CurrentWord = "";
            string CurrentCell = "";
            int LastRow = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            for (int nRowRun = 1, loopTo = LastRow; nRowRun <= loopTo; nRowRun++) // - 1
            {
                // CurrentCell = Trim(wSheet.Cells(nRowRun, nCol).value()) ' อาจต้องใช้ ตัวนี้ใน IF , ถ้า Trim ต้นฉบับอาจไม่เหมือน Copy
                CurrentCell = wSheet.Cells[nRowRun, ColumnNumber].value();
                if (CurrentCell != null)
                    CurrentWord = CurrentCell;
                else if (!string.IsNullOrEmpty(CurrentWord))
                    wSheet.Cells[nRowRun, ColumnNumber].Value = CurrentWord;
            }
        }



        /// <summary>  เช็ก SheetName ว่ามีอยู่ใน WorkBook หรือป่าว </summary>
    /// <param name="WB">Excel.Workbook</param> <param name="SheetName">ชื่อของ Sheet</param>
    /// <returns>มีหรือ ไม่มี</returns>
        public static bool SheetExists(Excel.Workbook WB, string SheetName)
        {
            bool SheetExistsRet = default(bool);
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            // SheetExists
            // This tests whether SheetName exists in a workbook. If R is
            // present, the workbook containing R is used. If R is omitted,
            // Application.Caller.Worksheet.Parent is used.
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Excel.Worksheet WS;

            try
            {
                Information.Err().Clear();
                WS = WB.Worksheets[SheetName];
                if (Information.Err().Number == 0)
                    SheetExistsRet = true;
                else
                    SheetExistsRet = false;

            }
            catch
            {
                SheetExistsRet = false;
            }
            return SheetExistsRet;
        }

        public static int GetSheetNumber(Excel.Workbook WB, string SheetName)
        {
            int ShtNum;
            if (SheetExists(WB, SheetName) == true)
                ShtNum = Conversions.ToInteger(WB.Sheets[SheetName].Index);
            else
                ShtNum = 0;
            return ShtNum;
        }

        /// <summary> ทำการ ไฮไลท์ ให้กับ Cell ที่มีคำที่กำหนด ตาม Column
    /// #สนใจทุกตัวอักษรใน Cell
    /// ,ไม่ได้ Trim ก่อนเปรียบเทียบ
    /// #ใช้การวนลูปดูทีละ Cell </summary>
    /// <param name="wSheet">Excel.Worksheet</param>
    /// <param name="strSearch">คำที่ต้องการค้นหา</param>
    /// <param name="ColumnNumberToFind">หมายเลข  Column ที่ต้องการหา</param>
    /// <param name="ColorIndex">หมายเลขของ Excel Interior Color  : 1-56</param>
    /// <param name="Req_Trim">ต้องการ Trim ก่อน เปรียบเทียบคำหรือไม่</param>
        public static void HiLightRowByWordInColumn(ref Excel._Worksheet wSheet, string strSearch, int ColumnNumberToFind, int ColorIndex, bool Req_Trim = true)
        {
            wSheet.Activate();
            int nLastRow = SheetLastRow(wSheet);
            for (int nRow = 1, loopTo = nLastRow; nRow <= loopTo; nRow++)
            {
                string CellStr = wSheet.Cells[nRow, ColumnNumberToFind].value;
                if (Operators.ConditionalCompareObjectEqual(Interaction.IIf(Req_Trim, Strings.Trim(CellStr), CellStr), strSearch, false))
                    wSheet.Rows[nRow].Interior.ColorIndex = ColorIndex;
            }
        }
        /// <summary> ทำการ ไฮไลท์ ให้กับ Cell ที่มีคำที่กำหนด ตาม Column
    /// #สนใจทุกตัวอักษรใน Cell
    /// ,ไม่ได้ Trim ก่อนเปรียบเทียบ
    /// #ใช้การ Find</summary>
    /// <param name="wSheet">Excel.Worksheet</param>
    /// <param name="strSearch">คำที่ต้องการค้นหา</param>
    /// <param name="RangeToFind">หมายเลข  Column ที่ต้องการหา</param>
    /// <param name="ColorIndex">หมายเลขของ Excel Interior Color  : 1-56</param>
        public static void HiLightRowByWordInRange(ref Excel._Worksheet wSheet, string RangeToFind, int ColorIndex
                                           , string strSearch
                                           , Excel.XlFindLookIn LookIn_Inp = Excel.XlFindLookIn.xlValues
                                           , Excel.XlLookAt LookAt_Inp = Excel.XlLookAt.xlWhole
                                           , Excel.XlSearchOrder SearchOrder_Inp = Excel.XlSearchOrder.xlByRows
                                           , Excel.XlSearchDirection SearchDirection_Inp = Excel.XlSearchDirection.xlNext
                                           , bool MatchCase_Inp = false)
        {
            wSheet.Activate();
            {
                var withBlock = wSheet.get_Range(RangeToFind);
                var Rng = withBlock.Find(What: strSearch
                              , After: withBlock.Cells[withBlock.Cells.Count]
                               , LookIn: LookIn_Inp
                               , LookAt: LookAt_Inp
                               , SearchOrder: SearchOrder_Inp
                               , SearchDirection: SearchDirection_Inp
                               , MatchCase: MatchCase_Inp);
                if (!(Rng == null))
                {
                    string firstAddress = Rng.Address;
                    do
                    {
                        Rng.EntireRow.Interior.ColorIndex = ColorIndex;
                        Rng = withBlock.FindNext(Rng);
                    }
                    while (!(Rng == null) & (Rng.Address ?? "") != (firstAddress ?? ""));
                }
            }
        }


        
        /// <summary>
        /// Copy Module ข้าม Workbook
        /// </summary>
        /// <param name="SourceWB"></param>
        /// <param name="strModuleName"></param>
        /// <param name="TargetWB"></param>
        /// <param name="n"></param>
        /// <remarks></remarks>
        public static void CopyModule(Excel.Workbook SourceWB, VBIDE.CodeModule strModuleName, Excel.Workbook TargetWB, int n)
        {
            string strFolder;
            // copies a module from one workbook to another
            // example: 
            // CopyModule Workbooks("Book1.xls"), "Module1", Workbooks("Book2.xls")
            // Dim cmpComponent As VBIDE.VBComponent
            string strTempFile;
            strFolder = SourceWB.Path;
            if (Strings.Len(strFolder) == 0)
                strFolder = FileSystem.CurDir();
            strFolder = strFolder + @"\";
            strTempFile = strFolder + "tmpexport.bas";
            // On Error Resume Next
            // Dim n As Integer = 1

            string moduleCode;
            moduleCode = readVbaToString(SourceWB, );
            SourceWB.VBProject.VBComponents.VBE.ActiveCodePane.Export(strTempFile);
            
            //Microsoft.Vbe.Interop.VBComponents cmpComponent;
            //cmpComponent = SourceWB.VBProject.VBComponents;
            //cmpComponent.

            TargetWB.VBProject.VBComponents.Import(strTempFile);
            FileSystem.Kill(strTempFile);
        }
        

    /// <summary>
    /// Copy Module ข้าม Workbook ทั้งหมด
    /// </summary>
    /// <param name="SourceWB"></param>
    /// <param name="TargetWB"></param>
    /// <remarks></remarks>
        public static void CopyAllVBACode(ref Excel.Workbook SourceWB, Excel.Workbook TargetWB)
        {
            VBIDE.VBProject VBProj;
            //VBIDE.VBComponent VBComp;
            VBIDE.CodeModule CodeMod;

            VBProj = SourceWB.VBProject;

            int n = 0;
            foreach (VBIDE.VBComponent VBComp in VBProj.VBComponents)
            {
                if ((int)VBComp.Type == (int)VBIDE.vbext_ComponentType.vbext_ct_StdModule)
                {
                    CodeMod = VBComp.CodeModule;
                    {
                        var withBlock = CodeMod;
                    }
                    // CopyModule(SourceWB, CodeMod.Name, TargetWB)
                    CopyModule(SourceWB, CodeMod, TargetWB, n);
                }
                else
                {
                }
                n += 1;
            }
        }

        public static string readVbaToString(Excel.Workbook SourceWB, String ModuleName)
        {
            String strCode = ""; 

            VBIDE.VBProject VBProj;
            //VBIDE.VBComponent VBComp;
            VBIDE.CodeModule CodeMod;

            VBProj = SourceWB.VBProject;

            int n = 0;
            foreach (VBIDE.VBComponent VBComp in VBProj.VBComponents)
            {
                if ((int)VBComp.Type == (int)VBIDE.vbext_ComponentType.vbext_ct_StdModule)
                {
                    CodeMod = VBComp.CodeModule;
                    {
                        var withBlock = CodeMod;
                    }
                    // CopyModule(SourceWB, CodeMod.Name, TargetWB)
                    // CopyModule(SourceWB, CodeMod, TargetWB, n);
                    if (CodeMod.Name == ModuleName)
                    {
                        int i  ;
                        
                        for( i = 1;i <= CodeMod.CountOfLines;i++){
                            strCode = strCode + CodeMod.Lines[i, 1].ToString();//& vbNewLine
                        }
                           
                    }
                }
                else
                {
                }
                n += 1;
            }
            return strCode;
        }
        /// <summary>
    /// Insert New Module ด้วย String ที่เป็น Code Macro
    /// </summary>
    /// <param name="xlApp"></param>
    /// <param name="wBook"></param>
    /// <param name="MacroCode"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static bool InsertMacro(ref Excel.Application xlApp, ref Excel.Workbook wBook, string MacroCode,string moduleName ="")
        {
            VBIDE.VBComponent oModule;
            try
            {
                // Create a new VBA code module.
                oModule = wBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
                if (moduleName != "")
                {
                    oModule.Name = moduleName;
                }
                // Add the VBA macro to the new code module.
                oModule.CodeModule.AddFromString(MacroCode);
                return true;
            }
            catch //(Exception ex)
            {
                return false;
            }
        }

        /// <summary>ทำการ Trim Cell โดย สร้าง และ เรียกใช้ Macro เร็วกว่าแบบวนลูปมากๆ</summary>
        public static bool ExcelMarcroTrimBColumn(ref Excel.Application xlApp, ref Excel.Workbook wBook, string ColumnName)
        {
            // Dim xlApp As Excel.Application
            // Dim wBook As Excel.Workbook
            VBIDE.VBComponent oModule;
            string sCode;

            try
            {
                // wBook = xlApp.Workbooks.Add

                // Create a new VBA code module.
                oModule = wBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

                sCode = "Sub TrimBText()" + Constants.vbCr + "   Dim MyRange As String" + Constants.vbCr + "   MyRange = \"" + ColumnName + "\"" + Constants.vbCr + "   Range(MyRange).Select" + Constants.vbCr + "   Dim MyCell As Range" + Constants.vbCr + "   On Error Resume Next" + Constants.vbCr + "       Selection.Cells.SpecialCells(xlCellTypeConstants, 23).Select" + Constants.vbCr + "       For Each MyCell In Selection.Cells" + Constants.vbCr + "           MyCell.Value = Trim(MyCell.Value)" + Constants.vbCr + "       Next" + Constants.vbCr + "   On Error GoTo 0" + Constants.vbCr + "End Sub" + Constants.vbCr;


                // Add the VBA macro to the new code module.
                oModule.CodeModule.AddFromString(sCode);
                {
                    var withBlock = xlApp;
                    withBlock.ScreenUpdating = false;
                    withBlock.Calculation = Excel.XlCalculation.xlCalculationManual; // xlManual
                    withBlock.EnableEvents = false;
                }

                xlApp.Run("TrimBText");

                {
                    var withBlock1 = xlApp;
                    withBlock1.ScreenUpdating = true;
                    withBlock1.Calculation = Excel.XlCalculation.xlCalculationAutomatic; // xlAutomatic
                    withBlock1.EnableEvents = true;
                }

                return true;
            }
            catch //(Exception ex)
            {
                return false;
            }
        }

        ///// <summary>สำหรับใช้ตอนที่ทำการ Query Table ด้วย Excel หรือ DataRecord ที่ตัวเลขจะกลายเป็น Text ทำให้ใช้คำนวนไม่ได้</summary>
        //public static bool ExcelMarcroCleanTextValue(ref Excel.Application xlApp, ref Excel.Workbook wBook)
        //{
        //    // Dim xlApp As Excel.Application
        //    // Dim wBook As Excel.Workbook
        //    VBIDE.VBComponent oModule;
        //    string sCode;

        //    try
        //    {
        //        // wBook = xlApp.Workbooks.Add
        //        // Create a new VBA code module.
        //        oModule = wBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

        //        sCode = "Sub CleanTextValue()" + Constants.vbCr + "   For Each Cell In ActiveSheet.UsedRange" + Constants.vbCr + "       Cell.value = Cell.value" + Constants.vbCr + "   Next Cell" + Constants.vbCr + "End sub" + Constants.vbCr;


        //        // Add the VBA macro to the new code module.
        //        oModule.CodeModule.AddFromString(sCode);
        //        {
        //            var withBlock = xlApp;
        //            withBlock.ScreenUpdating = false;
        //            withBlock.Calculation = Excel.XlCalculation.xlCalculationManual; // xlManual
        //            withBlock.EnableEvents = false;
        //        }

        //        xlApp.Run("CleanTextValue");

        //        {
        //            var withBlock1 = xlApp;
        //            withBlock1.ScreenUpdating = true;
        //            withBlock1.Calculation = Excel.XlCalculation.xlCalculationAutomatic; // xlAutomatic
        //            withBlock1.EnableEvents = true;
        //        }

        //        DeleteVBAMacroCode(ref wBook, "CleanTextValue");



        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}

    //    /// <summary>
    ///// คำสั่ง ลบ Macro จาก ไฟล์ Excel
    ///// </summary>
    ///// <param name="ActiveWorkbook"></param>
    ///// <param name="MacroName"></param>
    ///// <param name="ModuleName"></param>
    ///// <remarks></remarks>
        //public static void DeleteVBAMacroCode(ref Excel.Workbook ActiveWorkbook, string MacroName, string ModuleName = "Module1")
        //{
        //    object activeIDE; // VBProject
        //    activeIDE = ActiveWorkbook.VBProject;
        //    VBIDE.VBComponent Element; // VBComponent
        //                               // Dim LineCount As Integer
        //    foreach (var Element in activeIDE.VBComponents)
        //    {
        //        if ((int)Element.Type == VBIDE.vbext_ComponentType.vbext_ct_StdModule & (Element.Name ?? "") == (ModuleName ?? ""))
        //        {
        //            // LineCount = Element.CodeModule.CountOfLines
        //            int MacroStartLine = Element.CodeModule.get_ProcStartLine(MacroName, VBIDE.vbext_ProcKind.vbext_pk_Proc);
        //            int MacroCountLine = Element.CodeModule.get_ProcCountLines(MacroName, VBIDE.vbext_ProcKind.vbext_pk_Proc);

        //            Element.CodeModule.DeleteLines(MacroStartLine, MacroCountLine);
        //        }
        //    }
        //}

        /*public static void DeleteAllVBACode(ref Excel.Workbook ActiveWorkbook)
        {
            VBIDE.VBProject VBProj;
            VBIDE.VBComponent VBComp;
            VBIDE.CodeModule CodeMod;

            VBProj = ActiveWorkbook.VBProject;


            foreach (var VBComp in VBProj.VBComponents)
            {
                if ((int)VBComp.Type == (int)Microsoft.Vbe.VBIDE.vbext_ComponentType.vbext_ct_Document)
                {
                    CodeMod = VBComp.CodeModule;
                    {
                        var withBlock = CodeMod;
                        withBlock.DeleteLines(1, withBlock.CountOfLines);
                    }
                }
                else
                    VBProj.VBComponents.Remove(VBComp);
            }
        }*/


        // #########################################################################################
        // #########################################################################################    
        // #########################################################################################


        /// <summary> เซฟไฟล์ Excel เป็น CSV </summary>
    /// <param name="wSheet"> wSheet ที่ จะ เซฟ </param>
    /// <param name="PathSave"> ที่อยู่ตำแหน่งไฟล์ ที่ต้องการเซฟ </param>
        public static void ExcelSaveAsCsv(Excel._Worksheet wSheet, string PathSave)
        {
            wSheet.Activate();
            wSheet.SaveAs(PathSave, Excel.XlFileFormat.xlCSVWindows);
        }

        /// <summary> นำข้อมูลที่ Select ได้มา Copy ลง Excel Sheet ที่กำลัง Active **แต่ จะมี Column ติดมาด้วย ตัวเลขยังเป็น Text: อย่างเร็ว</summary>
    /// <param name="HostDB">IP database หรือ ชื่อ Host</param>
    /// <param name="DatabaseName">ชื่อฐานข้อมูล</param>
    /// <param name="UserDB"> ชื่อ ผู้ใช้ฐานข้อมูล</param>
    /// <param name="PassDB"> รหัส ผู้ใช้ฐานข้อมูล</param>
    /// <param name="wSheet">ตัวแปร อ็อบเจ็ก ชนิด Excel.WorkSheet</param>
    /// <param name="RangeStart">กำหนด เซลล์แรก ที่จะนำข้อมูลใส่ลงไป</param>
    /// <param name="sqlSelect">คำสั่ง Select ที่ดึงข้อมูลออกมา</param>
        public static void QueryTableToExcel(string HostDB, string DatabaseName, string UserDB, string PassDB, ref Excel._Worksheet wSheet, string RangeStart, string sqlSelect)
        {
            wSheet.Activate();
            string ConnString = string.Format("OLEDB;Provider=SQLOLEDB.1;Data Source={0};Initial Catalog={1};Persist Security Info=False;User ID={2} ;password={3}", HostDB, DatabaseName, UserDB, PassDB);

            wSheet.QueryTables.Add(ConnString, wSheet.get_Range(RangeStart), sqlSelect).Refresh(); // (BackgroundQuery:=False)
        }

        /// <summary>Query ตารางออกมาเป็น ไฟล์ Excel </summary>
    /// <param name="FileName">กำหนด FileName ใน AS400</param>
    /// <param name="MemberName">กำหนด MemberName ใน AS400</param>
    /// <param name="PathToSaveXls">ไฟล์ Output ที่ต้องการ</param>
        public static void QryAS400ToExcelFile(string FileName, string MemberName, string PathToSaveXls)
        {
            var xlApp = new Excel.Application();
            Excel.Workbook wBook;
            var wSheet = new Excel.Worksheet();
            wBook = xlApp.Workbooks.Add();
            wSheet = (Excel.Worksheet)wBook.Worksheets[1];

            var cn = new ADODB.Connection();
            var rs = new ADODB.Recordset();
            cn.Open("Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=pcs;Password=pcu8;Data Source=192.10.10.10;Force Translate=0;Catalog Library List=QS36F;SSL=DEFAULT;");
            string sqlSelect = string.Format("SELECT * FROM {0}({1})", FileName, MemberName);
            rs.Open(sqlSelect, cn);

            xlApp.ActiveCell.CopyFromRecordset(rs);
            xlApp.Visible = true;
            // wBook.SaveAs(PathToSaveXls, Excel.XlFileFormat.xlExcel8) 'xlExcel8 =97-2003  'ตอนใช้คำสั่งนี้ใช้ MS Office2013 แต่พอมาใช้ 2003 จะใช้ไม่ได้
            wBook.SaveAs(PathToSaveXls); // xlExcel8 =97-2003
        }

        public static void ExportDataTableToCsv(DataTable target, string outputFile)
        {
            var writer = new StreamWriter(outputFile, false);
            for (int rowIndex = 0, loopTo = target.Rows.Count - 1; rowIndex <= loopTo; rowIndex++)
            {
                for (int columnIndex = 0, loopTo1 = target.Columns.Count - 1; columnIndex <= loopTo1; columnIndex++)
                {
                    var fieldValue = target.Rows[rowIndex][columnIndex];
                    if (!Information.IsDBNull(fieldValue))
                    {
                        string output = fieldValue.ToString();
                        if (output.Contains(Conversions.ToString(',')))
                        {
                            writer.Write(ControlChars.Quote);
                            writer.Write(output);
                            writer.Write(ControlChars.Quote);
                        }
                        else
                            writer.Write(output);
                    }
                    if (columnIndex < target.Columns.Count - 1)
                        writer.Write(",");
                }
                if (rowIndex < target.Rows.Count - 1)
                    writer.WriteLine();
            }
            writer.Close();
        }

        public static void DataTableToCSV(ref DataTable DataTable, string PathSaveCSV)
        {
            // append = true คือ Text จะต่อจากข้อมูลเดิม  false คือ เริ่มข้อมูลใหม่
            using (var sw = new StreamWriter(PathSaveCSV, false, System.Text.Encoding.UTF8))
            {
                string DataCell = null;
                for (int i = 0, loopTo = DataTable.Rows.Count - 1; i <= loopTo; i++)
                {
                    for (int j = 0, loopTo1 = DataTable.Columns.Count - 1; j <= loopTo1; j++)
                    {
                        if (Information.IsDBNull(DataTable.Rows[i][j]))
                            DataCell = "";
                        else
                            DataCell = DataTable.Rows[i][j].ToString();
                        sw.Write("\"" + DataCell + "\"");
                        sw.Write(",");
                    }
                    sw.Write(sw.NewLine);
                }
            }
        }
        public static void DataTableToCSV_V1(DataTable table, string filename, string sepChar)
        {
            var writer = new StreamWriter(filename);
            try
            {

                // first write a line with the columns name
                string sep = "";
                var builder = new System.Text.StringBuilder();
                foreach (DataColumn col in table.Columns)
                {
                    builder.Append(sep).Append(col.ColumnName);
                    sep = sepChar;
                }
                writer.WriteLine(builder.ToString());

                // then write all the rows
                foreach (DataRow row in table.Rows)
                {
                    sep = "";
                    builder = new System.Text.StringBuilder();

                    foreach (DataColumn col in table.Columns)
                    {
                        builder.Append(sep).Append(row[col.ColumnName]);
                        sep = sepChar;
                    }
                    writer.WriteLine(builder.ToString());
                }
            }
            finally
            {
                if (!(writer == null))
                    writer.Close();
            }
        }

        public static DataTable ConvertExcelFileToDataTable(string PathFile, string SheetName, bool WantFirstRow = false)
        {
            // #### !!!! ฟังชั่นนี้มีปัญหา เปิดไฟล์ Excel ขึ้นมาเอง ตรง .Open -> ถ้า xlApp เปิดอยู่ ถึงจะเป็นไฟล์อื่นก็ตาม ถ้าเรียกใช้ จะใช้ Xlapp ตัวเดียวกันเปิด
            // --> แก้ไขแล้ว เป็นเพราะ ทำการเปิดไฟล์ ระหว่าง XlApp ยังทำงานอยู่ -> ให้แก้ไขโดย เรียกใช้หลัง xlApp Close เสร็จเรียบร้อยแล้ว
            var objConn = new OleDbConnection();
            var dtAdapter = new OleDbDataAdapter();
            var dt = new DataTable();

            string Row1IsHeader;
            if (WantFirstRow == true)
                Row1IsHeader = "No";
            else
                Row1IsHeader = "Yes";

            string strConnString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=" + PathFile + ";Extended Properties=\"Excel 8.0;HDR=" + Row1IsHeader + ";IMEX=1\"; ";
            // HDR=Yes ไม่ต้องการ Row แรก ถือว่าเป็น Column

            objConn = new OleDbConnection(strConnString);
            objConn.Open(); // !!!! มีปัญหา  เปิดไฟลื Excel ขึ้นมาเอง 

            string strSQL;
            strSQL = "SELECT * FROM [" + SheetName + "$]";

            dtAdapter = new OleDbDataAdapter(strSQL, objConn);
            dtAdapter.Fill(dt);

            dtAdapter = null;

            objConn.Close();
            objConn = null;
            return dt;
        }
        public static DataTable ConvertExcelFileToDataTableV2(string PathFile, string SheetName, bool WantFirstRow = false)
        {
            var objConn = new OleDbConnection();
            var dtAdapter = new OleDbDataAdapter();
            var dt = new DataTable();

            string Row1IsHeader;
            if (WantFirstRow == true)
                Row1IsHeader = "No";
            else
                Row1IsHeader = "Yes";

            string strConnString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=" + PathFile + ";Extended Properties=\"Excel 8.0;HDR=" + Row1IsHeader + ";IMEX=1\"; ";
            // HDR=Yes ไม่ต้องการ Row แรก ถือว่าเป็น Column

            objConn = new OleDbConnection(strConnString);
            objConn.Open();
            // Error Code :  External table is not in the expected format. =  ไฟล์ที่จะนำข้อมูลออก ไม่ใช่ Format .Xls  โดยแท้จริง

            string strSQL;
            strSQL = "SELECT * FROM [" + SheetName + "$]";

            dtAdapter = new OleDbDataAdapter(strSQL, objConn);
            dtAdapter.Fill(dt);

            dtAdapter = null;

            objConn.Close();
            objConn = null;

            return dt;
        }
        public static DataTable ConvertExcelFileToDataTableV3(string StrFilePath, string SheetName, bool WantFirstRow = false)
        {
            var objdt = new DataTable();
            var ExcelCon = new OleDbConnection();
            OleDbDataAdapter ExcelAdp;
            OleDbCommand ExcelComm;
            // Dim Col1 As DataColumn
            try
            {
                string Row1IsHeader;
                if (WantFirstRow == true)
                    Row1IsHeader = "No";
                else
                    Row1IsHeader = "Yes";
                // ExcelCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                // "Data Source= " & StrFilePath & _
                // ";Extended Properties=""Excel 8.0;"""
                ExcelCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=" + StrFilePath + ";Extended Properties=\"Excel 8.0;HDR=" + Row1IsHeader + ";IMEX=1\"; ";
                ExcelCon.Open();

                string StrSql = "SELECT * FROM [" + SheetName + "$]";
                ExcelComm = new OleDbCommand(StrSql, ExcelCon);
                ExcelAdp = new OleDbDataAdapter(ExcelComm);
                objdt = new DataTable();
                ExcelAdp.Fill(objdt);

                // --- Create Column With SRNo.
                // Col1 = New DataColumn
                // Col1.DefaultValue = 0
                // Col1.DataType = System.Type.GetType("System.Decimal")
                // Col1.Caption = "Sr No."
                // Col1.ColumnName = "SrNo"
                // objdt.Columns.Add(Col1)
                // Col1.SetOrdinal(1)

                ExcelCon.Close();
            }
            catch //(Exception ex)
            {
            }

            finally
            {
                ExcelCon = null;
                ExcelAdp = null;
                ExcelComm = null;
            }
            return objdt;
        }
        /// <summary>
    /// แปลง ด้วย การ เอามาทีละเซล ป้องกัน ว่า คอลลัมน์ ชอบมี ตัวเลข นำหน้า และตัวอักษร ที่ Row อื่น มันจะคิดว่า เป็นคอลัมน์ ชนิด Double
    /// ช้าไปหน่อย
    /// </summary>
        public static DataTable ConvertExcelFileToDataTableV4(string StrFilePath, string SheetName, bool WantFirstRow = false)
        {
            //DataTable dtRet = null;
            var dtWk = new DataTable();
            var xlApp = new Excel.Application();
            xlApp.Visible = false; // *
            xlApp.DisplayAlerts = false;
            {
                var withBlock = xlApp;
                withBlock.ScreenUpdating = false;
                withBlock.EnableEvents = false;
            }
            Excel.Workbook wBookOpen = null;
            var SheetOpen = new Excel.Worksheet();
            string SheetNameOpen = null;

            wBookOpen = xlApp.Workbooks.Open(StrFilePath);
            SheetOpen = (Excel.Worksheet)wBookOpen.Sheets[SheetName];
            SheetNameOpen = SheetOpen.Name;

            int nLastRow = SheetLastRow(SheetOpen);
            int nLastCol = SheetLastColumn(SheetOpen);

            // Dim dtRow As New DataRow
            for (int nCol = 1, loopTo = nLastCol; nCol <= loopTo; nCol++)
                dtWk.Columns.Add("F" + nCol.ToString("000"));


            for (int nRow = 1, loopTo1 = nLastRow; nRow <= loopTo1; nRow++)
            {
                var DataAry = new string[nLastCol - 1 + 1];
                for (int nCol = 1, loopTo2 = nLastCol; nCol <= loopTo2; nCol++)
                    DataAry[nCol - 1] = SheetOpen.Cells[nRow, nCol].value;
                dtWk.Rows.Add(DataAry);
            }

            wBookOpen.Close();

            // xlApp.Quit()
            xlApp = null;

            return dtWk;
        }

        /// <summary>
    /// วิธีนี้ แก้ปัญหา ใน Col มี ตัวเลข ที่ Row แรก และ Row อื่นเป็น Char ทำให้ Char หาย เพราะคิดว่า เป็น Col ชนิด Double
    /// ใช้วิธี Excel App เปิดไฟล์ออกมา แล้ว Insert Row แรก ให้เป็น F1 F2 F3...Fx แล้ว Query ออกมา เหมือนฐานข้อมูล
    /// วิธีนี้ ต้องอาศัยเครื่อง User ลงโปรแกรม Excel
    /// </summary>
    /// <param name="StrFilePath">ตำแหน่งไฟล์ที่ต้องการแปลง</param> <param name="SheetNumber">หมายเลขตำแหน่งชีท</param>
    /// <param name="ImportedOrginalFormat">
    /// 0=แปลงข้อมูลเป็น General ก่อนแล้วค่อยเอาเข้า
    /// ,1=แปลงข้อมูลเป็น Text ก่อนแล้วค่อยเอาเข้า
    /// ,2=นำไฟล์เข้าแบบใช้ข้อมูลตาม Format เลย
    /// </param>
    /// <returns>DataTable</returns>
        public static DataTable ConvertExcelFileToDataTableV5(string StrFilePath, int SheetNumber, int ImportedOrginalFormat = 0)
        {
            var dt = new DataTable();

            // สร้างตัวแปร เพื่อ การปิด Process เมื่อใช้เสร็จ เพราะ ปิดแบบ ปกติไม่ได้
            //object xlAppObj;
            //xlAppObj = Interaction.CreateObject("Excel.Application");
            Excel.Application xlAppObj = new Excel.Application();
            // ตรวจจับ Process ID ของ App เพื่อ จะได้ปิด Process ได้ถูก เมื่อ ฟังก์ชั่นนี้ ทำงานเสร็จ
            int xlHWND = Conversions.ToInteger(xlAppObj.Hwnd);
            int ProcIdXL = 0;
            fncProcessManager.GetWindowThreadProcessId((IntPtr)xlHWND, ref ProcIdXL);
            var xproc = Process.GetProcessById(ProcIdXL);

            // ###########################################################
            // ทำการ เตรียมไฟล์ Excel ที่จะ Qry ออกมาเป็น varchar ให้ทุก Col
            Excel.Application xlApp;
            xlApp = xlAppObj; // ถ่ายทอดให้ Excel App ของจริง เพื่อการใช้งาน ที่ง่ายขึ้น 
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false; // *'True '
            {
                var withBlock = xlApp;
                withBlock.ScreenUpdating = false;
                withBlock.EnableEvents = false;
            }
            Excel.Workbook wBookOpen = null;
            var SheetOpen = new Excel.Worksheet();
            string SheetNameOpen = null;
            try
            {
                wBookOpen = xlApp.Workbooks.Open(StrFilePath);

                // SheetOpen = CType(wBookOpen.Worksheets(SheetNumber), Excel.Worksheet)
                SheetOpen = (Excel.Worksheet)wBookOpen.Worksheets[SheetNumber];
                SheetOpen.Name = Strings.Trim(SheetOpen.Name); // ป้องกัน Error เนื่องจาก มีช่องว่างในชื่อ
                SheetNameOpen = SheetOpen.Name;

                // ปรับฟอร์แมตเซล ก่อน ไม่งั้น จะเข้าไปแบบที่ format กำหนดไว้
                switch (ImportedOrginalFormat)
                {
                    case 0:
                        {
                            SheetOpen.UsedRange.NumberFormat = "General";
                            break;
                        }

                    case 1:
                        {
                            SheetOpen.UsedRange.NumberFormat = "@"; // Text
                            break;
                        }

                    default:
                        {
                            break;
                        }
                }

                int nLastRow = SheetLastRow(SheetOpen);
                int nLastCol = SheetLastColumn(SheetOpen);

                SheetOpen.Rows[1].Insert(Shift: Excel.XlDirection.xlDown);
                for (int nCol = 1, loopTo = nLastCol; nCol <= loopTo; nCol++)
                    SheetOpen.Cells[1, nCol].value = "F" + nCol.ToString("000");

                // เก็บไฟล์ไว้ที่ตำแหน่ง Temp Folder
                string TempPath = Path.GetTempPath();
                string TempFileName = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                string tempXls = Path.Combine(TempPath, TempFileName); // + ".xls")

                // wBookOpen.SaveAs(tempXls, FileFormat:=Excel.XlFileFormat.xlWorkbookNormal)
                SaveAsWorkBookByVersion( xlApp, true, tempXls);
                tempXls = wBookOpen.FullName;
                wBookOpen.Close();
                xlAppObj.Quit();
                xlApp = null;

                // ###########################################################
                var objConn = new OleDbConnection();
                var dtAdapter = new OleDbDataAdapter();


                bool WantFirstRow = true;  // ต้องใช้ Row แรก เป็นตัวกำหนด DataType เป็น Varchar
                string Row1IsHeader; // 'HDR=Yes ไม่ต้องการ Row แรก ถือว่าเป็น Column
                if (WantFirstRow == true)
                    Row1IsHeader = "No";
                else
                    Row1IsHeader = "Yes";

                string strConnString; // = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
                                      // "Data Source=" & tempXls & ";Extended Properties=""Excel 8.0;HDR=" & Row1IsHeader & ";IMEX=1""; "

                strConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tempXls + ";Extended Properties=\"Excel 12.0;MaxScanRows=1;HDR=" + Row1IsHeader + ";IMEX=1\";";

                // strConnString = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
                // "Data Source=" & tempXls & ";Extended Properties=""Excel 8.0;HDR=" & Row1IsHeader & ";IMEX=1""; "

                objConn = new OleDbConnection(strConnString);
                objConn.Open();
                // Error Code :  External table is not in the expected format. =  ไฟล์ที่จะนำข้อมูลออก ไม่ใช่ Format .Xls  โดยแท้จริง

                string strSQL;
                strSQL = "SELECT * FROM [" + SheetNameOpen + "$]";

                dtAdapter = new OleDbDataAdapter(strSQL, objConn);
                dtAdapter.Fill(dt);

                dtAdapter = null;

                objConn.Close();
                objConn = null;
                // ###########################################################
                System.IO.File.Delete(tempXls);// ลบไฟล์ Copy ออก หลังจากเสร็จแล้ว
                dt.Rows[0].Delete(); // ก่อน ส่งกลับ ให้ ลบ Row F1 F2 F3 ....Fx ออกก่อน
                dt.AcceptChanges(); // แก้อาการฟ้อง Error => deleted row information cannot be accessed through the row. เมื่อเรียกใช้งาน
            }
            catch //(Exception ex)
            {
            }

            // '## นำตัวแปรจาก ข้างต้นมา ทำการปิด Process
            // System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppObj)
            if (!xproc.HasExited)
                xproc.Kill();

            return dt;
        }

        public static DataTable ConvertExcelFileToDataTableV5(string StrFilePath, string SheetName, int ImportedOrginalFormat = 0)
        {
            var dt = new DataTable();

            // สร้างตัวแปร เพื่อ การปิด Process เมื่อใช้เสร็จ เพราะ ปิดแบบ ปกติไม่ได้
         
            Excel.Application xlAppObj = new Excel.Application();
            // ตรวจจับ Process ID ของ App เพื่อ จะได้ปิด Process ได้ถูก เมื่อ ฟังก์ชั่นนี้ ทำงานเสร็จ
            int xlHWND = Conversions.ToInteger(xlAppObj.Hwnd);
            int ProcIdXL = 0;
            fncProcessManager.GetWindowThreadProcessId((IntPtr)xlHWND, ref ProcIdXL);
            var xproc = Process.GetProcessById(ProcIdXL);

            // ###########################################################
            // ทำการ เตรียมไฟล์ Excel ที่จะ Qry ออกมาเป็น varchar ให้ทุก Col
            Excel.Application xlApp;
            xlApp = xlAppObj; // ถ่ายทอดให้ Excel App ของจริง เพื่อการใช้งาน ที่ง่ายขึ้น 
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false; // *'True '
            {
                var withBlock = xlApp;
                withBlock.ScreenUpdating = false;
                withBlock.EnableEvents = false;
            }
            Excel.Workbook wBookOpen = null;
            var SheetOpen = new Excel.Worksheet();
            string SheetNameOpen = null;
            try
            {
                wBookOpen = xlApp.Workbooks.Open(StrFilePath);

                // SheetOpen = CType(wBookOpen.Worksheets(SheetNumber), Excel.Worksheet)
                SheetOpen = (Excel.Worksheet)wBookOpen.Worksheets[SheetName];
                SheetNameOpen = SheetOpen.Name;

                // ปรับฟอร์แมตเซล ก่อน ไม่งั้น จะเข้าไปแบบที่ format กำหนดไว้
                switch (ImportedOrginalFormat)
                {
                    case 0:
                        {
                            SheetOpen.UsedRange.NumberFormat = "General";
                            break;
                        }

                    case 1:
                        {
                            SheetOpen.UsedRange.NumberFormat = "@"; // Text
                            break;
                        }

                    default:
                        {
                            break;
                        }
                }

                int nLastRow = SheetLastRow(SheetOpen);
                int nLastCol = SheetLastColumn(SheetOpen);

                SheetOpen.Rows[1].Insert(Shift: Excel.XlDirection.xlDown);
                for (int nCol = 1, loopTo = nLastCol; nCol <= loopTo; nCol++)
                    SheetOpen.Cells[1, nCol].value = "F" + nCol.ToString("000");

                // เก็บไฟล์ไว้ที่ตำแหน่ง Temp Folder
                string TempPath = Path.GetTempPath();
                string TempFileName = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                string tempXls = Path.Combine(TempPath, TempFileName); // + ".xls")

                // wBookOpen.SaveAs(tempXls, FileFormat:=Excel.XlFileFormat.xlWorkbookNormal)
                SaveAsWorkBookByVersion( xlApp, true, tempXls);
                tempXls = wBookOpen.FullName;
                wBookOpen.Close();
                xlAppObj.Quit();
                xlApp = null;

                // ###########################################################
                var objConn = new OleDbConnection();
                var dtAdapter = new OleDbDataAdapter();


                bool WantFirstRow = true;  // ต้องใช้ Row แรก เป็นตัวกำหนด DataType เป็น Varchar
                string Row1IsHeader; // 'HDR=Yes ไม่ต้องการ Row แรก ถือว่าเป็น Column
                if (WantFirstRow == true)
                    Row1IsHeader = "No";
                else
                    Row1IsHeader = "Yes";

                string strConnString; // = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
                                      // "Data Source=" & tempXls & ";Extended Properties=""Excel 8.0;HDR=" & Row1IsHeader & ";IMEX=1""; "


                strConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tempXls + ";Extended Properties=\"Excel 12.0;MaxScanRows=1;HDR=" + Row1IsHeader + ";IMEX=1\";";

                // strConnString = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
                // "Data Source=" & tempXls & ";Extended Properties=""Excel 8.0;HDR=" & Row1IsHeader & ";IMEX=1""; "

                objConn = new OleDbConnection(strConnString);
                objConn.Open();
                // Error Code :  External table is not in the expected format. =  ไฟล์ที่จะนำข้อมูลออก ไม่ใช่ Format .Xls  โดยแท้จริง

                string strSQL;
                strSQL = "SELECT * FROM [" + SheetNameOpen + "$]";

                dtAdapter = new OleDbDataAdapter(strSQL, objConn);
                dtAdapter.Fill(dt);

                dtAdapter = null;

                objConn.Close();
                objConn = null;
                // ###########################################################
                // ลบไฟล์ Copy ออก หลังจากเสร็จแล้ว
                System.IO.File.Delete(tempXls); //My.MyProject.Computer.FileSystem.DeleteFile(tempXls); 

                dt.Rows[0].Delete(); // ก่อน ส่งกลับ ให้ ลบ Row F1 F2 F3 ....Fx ออกก่อน
                dt.AcceptChanges(); // แก้อาการฟ้อง Error => deleted row information cannot be accessed through the row. เมื่อเรียกใช้งาน
            }
            catch //(Exception ex)
            {
            }

            // '## นำตัวแปรจาก ข้างต้นมา ทำการปิด Process
            // System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppObj)
            if (!xproc.HasExited)
                xproc.Kill();

            return dt;
        }


        public static void ImportDataTableToDataBase(DataTable DataTableToImport, string connetionString, string ImportTableInDataBase)
        {
            // ใช้กับตารางอะไรก็ได้  ที่มี ชื่อ Column เป็น Col001 ขี้นไป -> แล้วไปทำ View แสดงอีกทีหลัง
            // Dim connetionString As String
            // OleDB==> "Provider=SQLOLEDB.1;Data Source=CO-Sahachart;Initial Catalog=FIFO;Persist Security Info=False;User ID=sa ;password=medline"
            // SQLCLIENT==> Data Source=192.168.4.6;Initial Catalog=Sahachart;Persist Security Info=True;User ID=sa;Password=ufida;Network Library=DBMSSOCN
            var connection = new OleDbConnection();

            var adapter = new OleDbDataAdapter();
            int i;
            string sql;
            string sql_TagColumn = null;
            string sql_TagValues = null;
            if (Strings.InStr(connetionString, "Provider") == 0)
                connetionString = "Provider=SQLOLEDB;" + connetionString;
            connection = new OleDbConnection(connetionString);
            connection.Open();

            // //////////////////////////////////////////////////////////////////
            // ///////////////////ลบข้อมูลในตาราง ก่อน ที่จะ Insert เข้าไปใหม่ \\\\\\\\\\\\\\\\\\\\\
            // //////////////////////////////////////////////////////////////////
            adapter.InsertCommand = new OleDbCommand("DELETE FROM " + ImportTableInDataBase, connection);
            adapter.InsertCommand.ExecuteNonQuery();
            adapter.InsertCommand = new OleDbCommand("DBCC CHECKIDENT ( " + ImportTableInDataBase + ",RESEED,0) ", connection);
            adapter.InsertCommand.ExecuteNonQuery();
            // '//////////////////////////////////////////////////////////////////

            // สร้าง Tag ของ Column ที่ใช้จริงใน Excel
            for (int nCol = 0, loopTo = DataTableToImport.Columns.Count - 1; nCol <= loopTo; nCol++)
            {
                if (nCol == 0)
                    sql_TagColumn = "Col" + (nCol + 1).ToString("000");
                else
                    sql_TagColumn += ",Col" + (nCol + 1).ToString("000");
            }

            var loopTo1 = DataTableToImport.Rows.Count - 1;
            for (i = 0; i <= loopTo1; i++)
            {
                var DataTableToImportCol = new DataColumn();
               
                for (int nCol = 0, loopTo2 = DataTableToImport.Columns.Count - 1; nCol <= loopTo2; nCol++)
                {
                    if (nCol == 0)
                        sql_TagValues = "'" + Strings.Replace(DataTableToImport.Rows[i][nCol].ToString(), "'", "''") + "'";
                    else
                        sql_TagValues += ",'" + Strings.Replace(DataTableToImport.Rows[i][nCol].ToString(), "'", "''") + "'";
                }
                // End If

                sql = "insert into " + ImportTableInDataBase + "(" + sql_TagColumn + ") values(" + sql_TagValues + ")";

                adapter.InsertCommand = new OleDbCommand(sql, connection);
                adapter.InsertCommand.ExecuteNonQuery();
            }
            connection.Close();
        }

        public static void ImportExcelToDataBase_Direct(Excel.Worksheet wSheet, string OleDbConnStr, string NewImprtTbInDb)
        {
            // ใช้กับตารางอะไรก็ได้  ที่มี ชื่อ Column เป็น Col001 ขี้นไป -> แล้วไปทำ View แสดงอีกทีหลัง
           
            var connection = new OleDbConnection();

            var adapter = new OleDbDataAdapter();
            //int i = 0;
            string sql;
            string sql_TagColumn = null;
            string sql_TagValues = "";

            connection = new OleDbConnection(OleDbConnStr);
            connection.Open();

            // //////////////////////////////////////////////////////////////////
            // ///////////////////ลบข้อมูลในตาราง ก่อน ที่จะ Insert เข้าไปใหม่ \\\\\\\\\\\\\\\\\\\\\
            // //////////////////////////////////////////////////////////////////
            adapter.InsertCommand = new OleDbCommand("DELETE FROM " + NewImprtTbInDb, connection);
            adapter.InsertCommand.ExecuteNonQuery();
            adapter.InsertCommand = new OleDbCommand("DBCC CHECKIDENT ( " + NewImprtTbInDb + ",RESEED,0) ", connection);
            adapter.InsertCommand.ExecuteNonQuery();
            // '//////////////////////////////////////////////////////////////////

            int nLastRow = SheetLastRow(wSheet);
            int nLastCol = SheetLastColumn(wSheet);


            // '''''''''''
            // สร้าง Tag ของ Column ที่ใช้จริงใน Excel
            for (int nCol = 0, loopTo = nLastCol - 1; nCol <= loopTo; nCol++)
            {
                if (nCol == 0)
                    sql_TagColumn = "Col" + (nCol + 1).ToString("000");
                else
                    sql_TagColumn += ",Col" + (nCol + 1).ToString("000");
            }

            foreach (Excel.Range cell in wSheet.UsedRange)
            {
                // cell.Value = Trim(cell.value)
                if (Operators.ConditionalCompareObjectEqual(cell.Column, 1, false))
                    sql_TagValues = "'" + Strings.Replace(cell.Value, "'", "''") + "'"; // ทำการแปลง  single qoute ในคำ ให้เป็น  single qoute ที่ VB ใช้ได้ ' <--> ''
                else
                {
                    sql_TagValues += ",'" + Strings.Replace(cell.Value, "'", "''") + "'";
                    if (Conversions.ToBoolean(!string.IsNullOrEmpty(sql_TagValues) & Operators.ConditionalCompareObjectEqual(cell.Column, nLastCol, false)))
                    {
                        sql = "insert into " + NewImprtTbInDb + "(" + sql_TagColumn + ") values(" + sql_TagValues + ")";

                        adapter.InsertCommand = new OleDbCommand(sql, connection);
                        adapter.InsertCommand.ExecuteNonQuery();
                        sql_TagValues = "";
                    }
                }
            }

            connection.Close();
        }

        public static void ImportExcelToDataBase(string PathExcel, string SheetName, string connetionString, string ImportTableInDataBase)
        {
            // หมายเหตุ: ฟังชั่นนี้ ไม่ได้ป้องกัน ผลจาก จาก text to Column แล้ว Cel ที่มีเครื่องหมาย - นำหน้า และ ชิดขวา หรือที่มีผลกระทบจากการคิดว่าเป็นสูตร
            // ควรใส่ ฟังชั่น Trim ให้ลบ ช่องว่าง " " ทางซ้าย,ขวา หรือ DestroySpace ให้ล้าง Cell ที่มีแต่ช่องว่าง " " ออกก่อน
            var DataTableFromExcel = new DataTable();
            // ///////////////////////////////////////////////////////////////////////////
            // @@@ ขั้นตอนการ Import ลง DataBase @@@
            DataTableFromExcel = ConvertExcelFileToDataTable(PathExcel, SheetName);
            // ///////////////////////////////////////////////////////////////////////////
            string ConStringImport = connetionString; // @@@@
            ImportDataTableToDataBase(DataTableFromExcel, ConStringImport, ImportTableInDataBase);
        }

        // ######################################################################################################
        // ######################################################################################################
        // ######################################################################################################
        /// <summary>
    /// ทำการนำ DataTable มา Copy Insert ในตำแหน่งที่กำหนด สร้างมาเพื่อ นำ DataTable ที่มีข้อมูลสำเร็จแล้ว มาวางใน WorkSheet พร้อมคำนวน
    /// : ใช้การ Loop ทีละ Row
    /// </summary>
    /// <param name="wSheet"></param>   <param name="DataTable"></param>
    /// <param name="OrdinatesCellToPaste">A1</param>
    /// <param name="nColumnPlusToDown">ใช้ในกรณีที่ต้องการ ให้ คอลัมน์ หลังคอลัมน์ที่มีอยู่ ถูกดึงลงไปด้วย เผื่อกรณี ต้องการ ให้มีสูตรด้วย Excel</param>
    /// <returns></returns>
    /// <param name="ShiftDownByCopy">กำหนดการ เลื่อน Row ลง ด้วยการ Copy หรือไม่ เผื่อกรณี Excel มีสูตรคำนวน</param>
        public static int PasteTableOnWorkSheet(Excel.Worksheet wSheet, DataTable DataTable, string OrdinatesCellToPaste = "A1"
                                       , int nColumnPlusToDown = 0, bool ShiftDownByCopy = false)
        {
            int NumCol = DataTable.Columns.Count;

            // หาตำแหน่งที่จะเป็นจุดเริ่มต้นที่จะวาง ตาราง จาก พิกัดที่ Input เข้ามา
            int nRowOrigin = wSheet.get_Range(OrdinatesCellToPaste).Row;
            int nColOrigin = wSheet.get_Range(OrdinatesCellToPaste).Column;

            int nColFinalData = nColOrigin + (NumCol - 1); // ระยะจริงของข้อมูล ที่มีข้อมูลจริง
            int nColFinalInsert = nColOrigin + (NumCol + nColumnPlusToDown - 1); // ระยะที่ต้อง Inser Shift Down จะมีผลในกรณี ที่มีการ เพิ่มระยะการ ShiftDown

            var OrdinatesFinalRngData = wSheet.Cells[nRowOrigin, nColFinalData].Address(RowAbsolute: false, ColumnAbsolute: false);
            var OrdinatesFinalRngInsert = wSheet.Cells[nRowOrigin, nColFinalInsert].Address(RowAbsolute: false, ColumnAbsolute: false);

            var RngInsertData = OrdinatesCellToPaste + ":" + OrdinatesFinalRngData;
            var RngInsert = OrdinatesCellToPaste + ":" + OrdinatesFinalRngInsert;

            // wSheet.Rows(SourceRowNumber).copy()
            // wSheet.Rows(TargetrRowNumber).Select()
            // wSheet.Rows(TargetrRowNumber).Insert(Shift:=Excel.XlDirection.xlDown)    'Insert multiple copied rows
            // wSheet.Paste()
            if (ShiftDownByCopy == false)
            {
                wSheet.get_Range(RngInsert).Select();
               // ReverseIterator RowList = new ReverseIterator(
                foreach (DataRow dRow in new ReverseIterator(DataTable.Rows))
                {
                    wSheet.get_Range(RngInsert).Insert(Shift: Excel.XlDirection.xlDown);
                    wSheet.get_Range(RngInsert).Value = dRow.ItemArray;
                }
            }
            else
                foreach (DataRow dRow in new ReverseIterator(DataTable.Rows))
                {
                    wSheet.get_Range(RngInsert).Select();
                    wSheet.get_Range(RngInsert).Copy();
                    wSheet.get_Range(RngInsert).Insert(Shift: Excel.XlDirection.xlDown);
                    // wSheet.Paste()
                    wSheet.get_Range(RngInsertData).Value = dRow.ItemArray;
                }

            return 0;
        }

        /// <summary>
    /// DataTable Save เป็นไฟล์ Excel
    /// </summary>
    /// <param name="Table">DataTable ที่ต้องการแปลง</param>
    /// <param name="PathToSaveXls">ตำแหน่งที่ต้องการเซฟ</param>
    /// <returns>ผล ได้หรือไม่ได้</returns>
        public static bool TableSaveAsExcelFile(DataTable Table, string PathToSaveXls)
        {
            try
            {
                var AppXl = new Excel.Application();
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                wBook = AppXl.Workbooks.Add();
                wSheet = wBook.Worksheets[1];
                DataTableToExcelSheet(ref Table, ref wSheet); // *
                wBook.SaveAs(PathToSaveXls);
                wBook.Close();
                AppXl.Quit();
                return true;
            }
            catch //(Exception ex)
            {
                return false;
            }
        }

        /// <summary>  แปลง DataTable เป็น Excel ชีท ใช้การ Loop เก็บค่าลง Cell </summary>
    /// <param name="DataTable">DataTable ที่ต้องการ</param>
    /// <param name="wSheet">Input WorkSheet ที่ต้องการเอา DataTable ลง</param>
        public static void DataTableToExcelSheet(ref DataTable DataTable, ref Excel.Worksheet wSheet)
        {
            // Dim wSheet As New Excel.Worksheet
           ((Excel._Worksheet) wSheet).Activate();
            // Dim s As String
            int nRow = DataTable.Rows.Count - 1;
            int nCol = DataTable.Columns.Count - 1;
            for (int i = 0, loopTo = nRow; i <= loopTo; i++)
            {
                for (int j = 0, loopTo1 = nCol; j <= loopTo1; j++)
                {
                    if (Information.IsDBNull(DataTable.Rows[i][j]) == true)
                        continue;
                    else
                        wSheet.Cells[i + 1, j + 1].value = DataTable.Rows[i][j].ToString();
                }
            }
        }

        /// <summary> นำข้อมูลจากการ Query เพื่อ Export ออกมาเป็น WorkSheet ของ Excel  ส่งค่ากลับมา เป็น DataTable</summary>
    /// <param name="wSheet">ตัวแปร WorkSheet ที่ต้องการจะนำมาเก็บข้อมูลตารางจากการ Query </param>
    /// <param name="OleDbCon">Connection ของ OleDbConnection ที่ทำการเชื่อมต่อแล้ว</param>
    /// <param name="sqlSelect">คำสั่ง Query ที่ต้องการนำข้อมูลออกมาเป็นตาราง</param>
    /// <remarks>อะรูไม่ไร้</remarks>
        public static DataTable QryTableToExcelSheet(ref Excel.Worksheet wSheet, ref OleDbConnection OleDbCon, string sqlSelect) // ฟังชั้นนี้อาจมีการเปลี่ยนแปลง เพราะ ใช้ oledb
        {
           ((Excel._Worksheet) wSheet).Activate();
            var myDataAdapter = new OleDbDataAdapter(sqlSelect, OleDbCon);
            var myDataTable = new DataTable();
            myDataTable.Clear();
            myDataAdapter.Fill(myDataTable);
            DataTableToExcelSheet(ref myDataTable, ref wSheet);
            return myDataTable;
        }

        /// <summary>นำข้อมูลที่ Select ได้มา Copy ลง Excel Sheet ที่กำลัง Active ตัวเลขยังเป็น Text
    /// # ต้อง Clean ตัวเลขเพราะยังเป็น Text
    /// # อย่างเร็ว เพราะ ใช้คำสั่งของ Excel Copy Table Recordset โดยตรง</summary>
    /// <param name="OLEDBConnectionString">Connection String สำหรับเชื่อมต่อฐานข้อมูล ของ Oledb
    /// "Provider=SQLOLEDB.1;Data Source=(IPเครื่อง);Initial Catalog=(ชื่อฐานข้อมูล);Persist Security Info=False;User ID=(ชื่อUSER) ;password=(รหัสUSER)"</param>
    /// <param name="sqlSelect">คำสั่ง SELECT ข้อมูลที่ต้องการ </param>
    /// <param name="xlApp">ตัวแปร Excel.Application ที่ กำหนดชีทแล้ว Active ที่ Cell แรก  หรือตามต้องการ  โดย xlApp.Cells(nRow, nCol).select() ก่อน</param>
        public static void ExportTableToExcel(string OLEDBConnectionString, string sqlSelect, ref Excel.Application xlApp, string TargetCell = null, bool WantInsertRow = false)
        {
            var cn = new ADODB.Connection();
            var rs = new ADODB.Recordset();

            cn.ConnectionTimeout = 1600;
            cn.CommandTimeout = 1600;
            cn.Open(OLEDBConnectionString);

            rs.Open(sqlSelect, cn, ADODB.CursorTypeEnum.adOpenStatic);

            if (TargetCell != null)
                xlApp.get_Range(TargetCell).Select();

            if (WantInsertRow == true)
            {
                long OriginalRow;
                Excel.Range TargetRng;
                OriginalRow = xlApp.get_Range(TargetCell).Row;
                TargetRng = xlApp.get_Range(TargetCell);
                xlApp.Rows[Conversions.ToString(TargetRng.Row) + ":" + Conversions.ToString(TargetRng.Row + rs.RecordCount - 1)].Select();
                xlApp.Selection.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            xlApp.ActiveCell.CopyFromRecordset(rs);

            cn.Close();
            // rs.Close()
            cn = null;
            rs = null;
        }




        public enum XLColor
        {
            Aqua = 42,
            Black = 1,
            Blue = 5,
            BlueGray = 47,
            BrightGreen = 4,
            Brown = 53,
            Cream = 19,
            DarkBlue = 11,
            DarkGreen = 51,
            DarkPurple = 21,
            DarkRed = 9,
            DarkTeal = 49,
            DarkYellow = 12,
            Gold = 44,
            Gray25 = 15,
            Gray40 = 48,
            Gray50 = 16,
            Gray80 = 56,
            Green = 10,
            Indigo = 55,
            Lavender = 39,
            LightBlue = 41,
            LightGreen = 35,
            LightLavender = 24,
            LightOrange = 45,
            LightTurquoise = 20,
            LightYellow = 36,
            Lime = 43,
            NavyBlue = 23,
            OliveGreen = 52,
            Orange = 46,
            PaleBlue = 37,
            Pink = 7,
            Plum = 18,
            PowderBlue = 17,
            Red = 3,
            Rose = 38,
            Salmon = 22,
            SeaGreen = 50,
            SkyBlue = 33,
            Tan = 40,
            Teal = 14,
            Turquoise = 8,
            Violet = 13,
            White = 2,
            Yellow = 6
        }

        /// <summary>
    /// Refresh ทีละเซล เหมาะกับการใช้ กับ Sheet ที่ เพิ่งวาง Table ลงใหม่ๆ
    /// => แล้วมีปัญหากับ ฟิลด์ที่เป็น Number แต่ "Format เป็น Text" ทำให้ใช้คำนวนไม่ได้
    /// </summary>
    /// <param name="wsht">Worksheet</param>
    /// <param name="Range">Range</param>
    /// <remarks></remarks>
        public static void RefreshCellByRange(Excel._Worksheet wsht, Excel.Range Range)
        {
            wsht.Activate();
            //Excel.Range Cell;
            foreach (Excel.Range Cell in Range)
                Cell.Value = Cell.Value;
        }

        /// <summary>
    /// เซ็ตเปิดปิด Macro ใน Registry
    /// </summary>
    /// <param name="Choice">ตัวเลือก ตามความเหมาะส ==>
    /// <para>/ 1 = Enable(All)</para>
    /// <para>/ 2 = Disable All with Notification</para>
    /// <para>/ 3 = Disable All Except Digitally Signed</para>
    /// <para>/ 4 = Disable All without Notification</para>
    /// </param>
    /// <remarks>610524 กำหนดให้ใช้กับ Excel 2007 เท่านั่น</remarks>
        public static void SetEnableMacro(int Choice = 1)
        {
            // Version Number =>  2007 = 12, 2010 = 14 ,2013 = 15 ,2016 = 16
            string RegMacro = @"HKEY_CURRENT_USER\Software\Microsoft\Office\{0}\Excel\Security";
            foreach (string ExcelVersion in new[] { "12.0", "14.0", "15.0", "16.0" })
            {
                string RegMcr = string.Format(RegMacro, ExcelVersion);
                Registry.SetValue(RegMcr, "VBAWarnings", Choice);
            }
            string strVBASecuritySetting = (string) Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Excel\Security", "VBAWarnings", string.Empty);
        }

        /// <summary>
    /// เซ็ตเปิดปิด Protect View ใน Registry
    /// เปิดปิด Protect File จาก Network
    /// </summary>
    /// <param name="Choice">ตัวเลือก ตามความเหมาะสม ==>
    /// 1 = เปิด
    /// / 0 = ปิด
    /// </param>
    /// <remarks>610524 กำหนดให้ใช้กับ Excel 2016 เท่านั่น</remarks>
        public static void SetOpenProtectViewFormNet(int Choice = 1)
        {
            // 1 เปิด
            // 0 ปิด
            // Version Number => 2007 = 12, 2010 = 14 ,2013 = 15 ,2016 = 16
            string RegProtectedView = @"HKEY_CURRENT_USER\Software\Microsoft\Office\{0}\Excel\Security\ProtectedView";
            foreach (string ExcelVersion in new[] { "12.0", "14.0", "15.0", "16.0" }) // ใส่ให้หมดเลย
            {
                string RegPrtVw = string.Format(RegProtectedView, ExcelVersion);
                Registry.SetValue(RegPrtVw, "DisableInternetFilesInPV", Choice);
                Registry.SetValue(RegPrtVw, "DisableAttachmentsInPV", Choice);
                Registry.SetValue(RegPrtVw, "DisableUnsafeLocationsInPV", Choice);
            }
        }


        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59)
                SerialDate -= 1; // '// Excel/Lotus 2/29/1900 bug
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }

        public static String sheetNameUnique(Excel.Workbook wb,string shtName){

            int Counter = 1;
            String strName = shtName;

            while (SheetExists(wb, strName))
            {
                strName = shtName + '_' + Counter;
                Counter = Counter + 1;
            }

            return strName;
        }
    }
}
