﻿using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using EXCEL = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;
using System.Collections;
using System.Collections.Generic;
namespace kgLibraryCs
{
    /// <summary>
    /// สร้างตัวแปร เพื่อ การปิด Process เมื่อใช้เสร็จ เพราะ ปิดแบบปกติไม่ได้
    /// </summary>
    public class ExcelApp
    {
        public EXCEL.Application App = new EXCEL.Application();

        // Dim xlAppObj As Object 'ลองปิดดูก่อน แล้วไป เปิด ใน new เนื่องจาก มีการสร้าง Excel App ขึ้น 3
        private int xlHWND; // = xlAppObj.Hwnd
        private int ProcIdXL = 0;
        private Process xproc;

        public ExcelApp()
        {

            // XlsApp.Workbooks.Add()
            // System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            // ตรวจจับ Process ID ของ App เพื่อ จะได้ปิด Process ได้ถูก เมื่อ ฟังก์ชั่นนี้ ทำงานเสร็จ
            xlHWND = App.Hwnd; // xlAppObj.Hwnd

            fncProcessManager.GetWindowThreadProcessId((IntPtr)xlHWND, ref ProcIdXL);
            // Dim xproc As Process = Process.GetProcessById(ProcIdXL)
            xproc = Process.GetProcessById(ProcIdXL);
        }

        public ExcelApp(ref EXCEL.Application AppXls) // สำหรับ อ้างอิง App ที่มีอยู่แล้ว
        {
            // XlsApp.Workbooks.Add()
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            App = AppXls;
        }

        public EXCEL.Workbook Add_NewWorkBook()
        {
            return App.Workbooks.Add();
        }
        public EXCEL.Worksheet Add_NewWorkSheet(ref EXCEL.Workbook wBook, int ShtAddressInsert = 1)
        {
            try
            {
                return wBook.Worksheets.Add(wBook.Worksheets[ShtAddressInsert]);
            }
            catch //(Exception ex)
            {
                return wBook.Worksheets.Add();
            }
        }
        public EXCEL.Workbook wBookOpenFile(string PathFile)
        {
            return App.Workbooks.Open(PathFile);
        }

        public EXCEL.Worksheet wSheet(ref EXCEL.Workbook wBook, int index)
        {
            return (EXCEL.Worksheet)wBook.Worksheets[index];
        }
        public EXCEL.Workbook wBook(int index)
        {
            return App.Workbooks[index];
        }

        public bool SheetExists(EXCEL.Workbook WB, string SheetName)
        {
            bool SheetExistsRet = default(bool);
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            // SheetExists
            // This tests whether SheetName exists in a workbook. If R is
            // present, the workbook containing R is used. If R is omitted,
            // Application.Caller.Worksheet.Parent is used.
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            EXCEL.Worksheet WS;

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

        public void CloseExcel()
        {
            // '## นำตัวแปรจาก ข้างต้นมา ทำการปิด Process
            // System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppObj)
            App = null;
            // App.Quit() 'error
            if (!xproc.HasExited)
                xproc.Kill();
        }

        /// <summary>
        /// บันทึกไฟล์ Excel ด้วยเวอร์ชั่นที่ต่างกัน ซึ่งจะมีค่าพารามิเตอร์ที่ ไม่เหมือนกันในบางเวอร์ชั่น เฉพาะ .XLS และ .XLSX เท่านั้น
        /// </summary>
        /// <param name="BoolSaveToXlsx">ใช่คือ XLSX ไม่ใช่คือ XLS</param>
        /// <param name="FilePathToSaveWithOutExt">ตำแหน่งเซฟไฟล์ ไม่ต้องการ นามสกุลไฟล์</param>
        /// <returns>ส่งออกมาเป็นตำแหน่งไฟล์ เต็ม</returns>
        /// <remarks></remarks>
        public string WbookSaveAsByVersion(bool BoolSaveToXlsX, string FilePathToSaveWithOutExt)
        {

            // เลือกไฟล์ ฟอร์แมต สำหรับ ใช้ 2003 กับ 2007 
            int EnumFileFormat = 56;
            string FileFormat = "xls";
            int XlsVer = Conversions.ToInteger(App.Version);
            if (BoolSaveToXlsX == true)
            {
                switch (XlsVer)
                {
                    case 11:// Office 2003 ถ้าเป็น 2003 ให้ Save เป็น xls ธรรมดา
                        EnumFileFormat = 51;//43; // Excel.XlFileFormat.xlExcel9795
                        FileFormat = "xls";
                        break;
                    default:
                        EnumFileFormat = 51;
                        FileFormat = "xlsx";
                        break;
                }
            }
            else
                switch (XlsVer)
                {
                    case 11: // Office 2003
                        EnumFileFormat = 43; // Excel.XlFileFormat.xlExcel9795
                        FileFormat = "xls";
                        break;
                    default:
                        EnumFileFormat = 56;
                        FileFormat = "xls";
                        break;
                }

            string FileNameSaveFull;
            FileNameSaveFull = FilePathToSaveWithOutExt + "." + FileFormat;
            FileNameSaveFull = FncFileFolder.NewFileNameUnique(FileNameSaveFull); // สร้างชื่อไฟล์ใหม่ หาก มีไฟล์ชื่อเดิมอยู่แล้ว

            App.ActiveWorkbook.SaveAs(FileNameSaveFull, EnumFileFormat);
            return FileNameSaveFull;
        }

        //    ''' <summary>
        //''' Move หลายๆ ชีท เพื่อ สร้าง WorkBook ใหม่  (OverLoad Function)
        //''' </summary>
        //''' <param name="XlsApp">ClassExcelApp</param>
        //''' <param name="wBook">Workbook ที่เก็บ Worksheet อยู่</param>
        //''' <param name="wSht_ArrayList">ตัวแปร Array List ที่เก็บ WorkSheet เข้ามา</param>
        //''' <param name="PathFile">ตำแหน่ง ชื่อไฟล์ ที่จะเซฟเก็บไว้</param>
        //''' <remarks>590426 OverLoad Function</remarks>
        //Sub SaveSheetToNewWorkBook(XlsApp As ClassExcelApp, wBook As EXCEL.Workbook, wSht_ArrayList As ArrayList, PathFile As String)
        //    wBook.Activate()
        //    ' wSht.Activate()

        //    Dim wShtName As New ArrayList
        //    For Each wSht In wSht_ArrayList
        //        wShtName.Add(wSht.name)
        //    Next
        //    Dim StrwShtNm = wShtName.ToArray()

        //    wBook.Sheets(StrwShtNm).Move() '* New workBook และ Move ชีทนี้ไปที่ New WorkBook พร้อมๆกัน
        //    Dim wBkRpt = XlsApp.App.ActiveWorkbook 'workBook ที่เพิ่ง new 
        //    wBkRpt.SaveAs(PathFile, EXCEL.XlFileFormat.xlWorkbookNormal) '.XLS
        //    wBkRpt.Close() ' ปิด workbook ก่อน ไปทำงาน อื่นต่อ * เอาออก เมื่อต้องใช้ชีทนี้ทำต่อ

        //    wBook.Activate() 'กลับไปที่ WorkBook เก่า

        //End Sub

        public void SaveSheetToNewWorkBook(EXCEL.Workbook wBook,  ArrayList wSht_ArrayList ,string PathFile)//)//  
        {

            ((EXCEL._Workbook)wBook).Activate();
            //ArrayList wShtName = new ArrayList()

            List<string> LstwShtName = new List<string>();

            foreach (EXCEL.Worksheet wSht in wSht_ArrayList)
            {
                LstwShtName.Add(wSht.Name);
            }

            string[] StrwShtNm = LstwShtName.ToArray();

            //    wBook.Sheets(StrwShtNm).Move() '* New workBook และ Move ชีทนี้ไปที่ New WorkBook พร้อมๆกัน
            wBook.Sheets[StrwShtNm].Move();

            //    Dim wBkRpt = XlsApp.App.ActiveWorkbook 'workBook ที่เพิ่ง new 
            EXCEL.Workbook wBkRpt = App.ActiveWorkbook;

            wBkRpt.SaveAs(PathFile, EXCEL.XlFileFormat.xlWorkbookNormal);
            wBkRpt.Close();// ปิด workbook ก่อน ไปทำงาน อื่นต่อ * เอาออก เมื่อต้องใช้ชีทนี้ทำต่อ

            ((EXCEL._Workbook)wBook).Activate();//กลับไปที่ WorkBook เก่า

        }





    }
}
