using System.Data;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Devices;
using System.Linq;
using System.Collections;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; // add reference
using Colors = System.Drawing.Color;
using System.Data.OleDb;
using System.IO;
using System.Threading;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{

    // ###########################################################################
    // ###########     อย่าลืม    Add Reference ADODB  ด้วย
    // ###########################################################################
    // 

    public static class CFncAS400
    {
        private const string AS400_ServerIP = "192.10.10.10";
        private const string AS400_Library = "QS36F";
        private const string AS400_User = "ALL"; // ใช้ ALL ปลอยภัยกว่า '"pcs"
        private const string AS400_Password = "A550555"; // "pcssap"
                                                         // Const HostCheckMember = "172.28.2.125" 'เครื่อง คุณ Turbo

        public const string AS400_P7_ServerIP = "172.29.9.100";


        /// <summary>Qry ออกมาเป็น DataTable * ข้อเสีย คือ ไม่สามารถ Qry ชื่อไฟล์ที่มีจุดได้</summary>
        public static DataTable QryAS400ToDatatable(string FileName, string MemberName)
        {
            // Dim DataSet As New DataSet
            var dt = new DataTable();

            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            string Sql = string.Format("SELECT * FROM QS36F." + FileName + "(" + MemberName + ")", FileName, MemberName);
            // Dim Sql As String = String.Format("SELECT * FROM {0}({1})", FileName, MemberName)
            var AS400Connection = new OleDbConnection(ConnStr);
            // Try
            AS400Connection.Open();
            var Adapter = new OleDbDataAdapter(Sql, AS400Connection);
            Adapter.Fill(dt); // (DataSet)
                              // Catch ex As Exception
                              // MsgBox(ex.Message)
                              // Finally
            AS400Connection.Close();
            // End Try
            return dt;
        }



        /// <summary> Qry เป็น Datatable รองรับ FileName ที่มีจุดในชื่อ โดยจะกำหนด MemberName หรือไม่ก็ได้</summary>
        public static DataTable QryAS400ToDatatableV2(string FileName, string MemberName = null, string IPAs400 = AS400_ServerIP)
        {
            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + IPAs400 + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            var dt = new DataTable();
            var AS400Connection = new OleDbConnection(ConnStr);
            AS400Connection.Open();

            // String aliasCommand = "CREATE ALIAS Qtemp.getMember " +"FOR "+libName+"."+fileName+"("+memberName+")";
            // CREATE ALIAS Qtemp.getMember FOR QS36F."wf.h17"(m561018)
            // ### ชื่อไฟล์ ตัวพิมพ์ใหญ่
            string SqlSelect;
            if (MemberName == null | string.IsNullOrEmpty(MemberName))
            {
                SqlSelect = string.Format("SELECT * FROM QS36F.\"{0}\"", FileName);
                var Adapter = new OleDbDataAdapter(SqlSelect, AS400Connection);
                Adapter.Fill(dt);
            }
            else if (FileName.IndexOf(".") > 0)
            {
                string sqlDo;
                sqlDo = "CREATE ALIAS Qtemp.getMember FOR " + AS400_Library + ".\"" + FileName + "\"(" + MemberName + ")";
                var myCommand = new OleDbCommand(sqlDo, AS400Connection);
                int IntResult = myCommand.ExecuteNonQuery();
                SqlSelect = string.Format("SELECT * FROM Qtemp.getMember");
                var Adapter = new OleDbDataAdapter(SqlSelect, AS400Connection);
                Adapter.Fill(dt);
                sqlDo = "DROP ALIAS Qtemp.getMember";
                myCommand = new OleDbCommand(sqlDo, AS400Connection);
                IntResult = myCommand.ExecuteNonQuery();
            }
            else
            {
                string Sql = string.Format("SELECT * FROM QS36F." + FileName + "(" + MemberName + ")", FileName, MemberName);
                var Adapter = new OleDbDataAdapter(Sql, AS400Connection);
                Adapter.Fill(dt);
            }

            AS400Connection.Dispose();
            AS400Connection.Close();
            AS400Connection = null;

            return dt;
        }

        /// <summary> Qry เป็น Datatable ด้วย ADO RecordSet เร็วกว่า </summary>
        public static DataTable QryAS400ToDatatableByADO(string FileName, string MemberName)
        {
            // Dim ServerIP = "192.10.10.10", Library = "QS36F", User = "pcs", Password = "pcu8"
            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            var cn = new ADODB.Connection();
            cn.Open(ConnStr);

            if (MemberName != null)
                MemberName = string.Format("({0})", MemberName);
            string SqlSelect = string.Format("SELECT * FROM QS36F.\"{0}\"{1}", FileName, MemberName);

            var rs1 = new ADODB.Recordset();
            var Adapter = new OleDbDataAdapter(); // #
            var dt = new DataTable(); // #
            rs1.Open(SqlSelect, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1);
            Adapter.Fill(dt, rs1);
            return dt;
        }


        /// <summary> เช็คว่ามี File Name นี้อยู่หรือป่าว ใช้ คำสั่ง ชุดเดียวกับ  QryAS400ToDatatableV2 </summary>
        public static bool AS400FileExits(string FileName, string MemberName = null)
        {
            bool ReturnValue;
            Network netw = new Network();
            if (netw.Ping(AS400_ServerIP) == false)
                return false;
            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            var dt = new DataTable();
            var AS400Connection = new OleDbConnection(ConnStr);


            AS400Connection.Open();


            string sqlDo;
            OleDbCommand myCommand;
            int IntResult;

            // String aliasCommand = "CREATE ALIAS Qtemp.getMember " +"FOR "+libName+"."+fileName+"("+memberName+")";
            // CREATE ALIAS Qtemp.getMember FOR QS36F."wf.h17"(m561018)
            // ### ชื่อไฟล์ ตัวพิมพ์ใหญ่
            try
            {
                FileName = FileName.ToUpper();
                string SqlSelect;
                if (MemberName == null | string.IsNullOrEmpty(MemberName))
                {
                    SqlSelect = string.Format("SELECT count(*) FROM QS36F.\"{0}\"", FileName);
                    var Adapter = new OleDbDataAdapter(SqlSelect, AS400Connection);
                    Adapter.Fill(dt);
                }
                else if (FileName.IndexOf(".") > 0)
                {

                    sqlDo = "CREATE ALIAS Qtemp.getMember FOR " + AS400_Library + ".\"" + FileName + "\"(" + MemberName + ")";
                    myCommand = new OleDbCommand(sqlDo, AS400Connection);
                    IntResult = myCommand.ExecuteNonQuery();
                    SqlSelect = string.Format("SELECT count(*) FROM Qtemp.getMember");
                    var Adapter = new OleDbDataAdapter(SqlSelect, AS400Connection);
                    Adapter.Fill(dt);


                    sqlDo = "DROP ALIAS Qtemp.getMember";
                    myCommand = new OleDbCommand(sqlDo, AS400Connection);
                    IntResult = myCommand.ExecuteNonQuery();
                }
                else
                {
                    string Sql = string.Format("SELECT count(*) FROM QS36F." + FileName + "(" + MemberName + ")", FileName, MemberName);
                    var Adapter = new OleDbDataAdapter(Sql, AS400Connection);
                    Adapter.Fill(dt);
                }

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                try // ดักไว้อีกชั้น กรณี ไม่ได้กำหนด Member จึงไม่ได้สร้าง Alias ก็จะ Drop Alias ไม่ได้
                {
                    sqlDo = "DROP ALIAS Qtemp.getMember";
                    myCommand = new OleDbCommand(sqlDo, AS400Connection);
                    IntResult = myCommand.ExecuteNonQuery();
                }
                catch (Exception ex02)
                {
                }

                ReturnValue = false;
            }
            finally
            {
                AS400Connection.Dispose();
                AS400Connection.Close();
                AS400Connection = null;
            }

            return ReturnValue;
        }

        /// <summary> เช็คว่ามี File Name นี้อยู่หรือป่าว ใช้ คำสั่ง ชุดเดียวกับ  QryAS400ToDatatableV2 </summary>
        public static bool AS400_P7_FileExits(string FileName, string MemberName = null)
        {
            bool ReturnValue;
            Network netw = new Network();
            if (netw.Ping(AS400_P7_ServerIP) == false)
                return false;
            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_P7_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            var dt = new DataTable();
            var AS400Connection = new OleDbConnection(ConnStr);


            AS400Connection.Open();


            string sqlDo;
            OleDbCommand myCommand;
            int IntResult;

            // String aliasCommand = "CREATE ALIAS Qtemp.getMember " +"FOR "+libName+"."+fileName+"("+memberName+")";
            // CREATE ALIAS Qtemp.getMember FOR QS36F."wf.h17"(m561018)
            // ### ชื่อไฟล์ ตัวพิมพ์ใหญ่
            try
            {
                FileName = FileName.ToUpper();
                string SqlSelect;
                if (MemberName == null | string.IsNullOrEmpty(MemberName))
                {
                    SqlSelect = string.Format("SELECT count(*) FROM QS36F.\"{0}\"", FileName);
                    var Adapter = new OleDbDataAdapter(SqlSelect, AS400Connection);
                    Adapter.Fill(dt);
                }
                else if (FileName.IndexOf(".") > 0)
                {
                    // Dim Rnd2Digit As String = GetRandom(0, 99) 'ใช้ตั้ง Alias ไม่ให้ซ้ำกัน *ตอนนี้ยังไม่ใช้

                    sqlDo = "CREATE ALIAS Qtemp.getMember FOR " + AS400_Library + ".\"" + FileName + "\"(" + MemberName + ")";
                    myCommand = new OleDbCommand(sqlDo, AS400Connection);
                    IntResult = myCommand.ExecuteNonQuery();
                    SqlSelect = string.Format("SELECT count(*) FROM Qtemp.getMember");
                    var Adapter = new OleDbDataAdapter(SqlSelect, AS400Connection);
                    Adapter.Fill(dt);


                    sqlDo = "DROP ALIAS Qtemp.getMember";
                    myCommand = new OleDbCommand(sqlDo, AS400Connection);
                    IntResult = myCommand.ExecuteNonQuery();
                }
                else
                {
                    string Sql = string.Format("SELECT count(*) FROM QS36F." + FileName + "(" + MemberName + ")", FileName, MemberName);
                    var Adapter = new OleDbDataAdapter(Sql, AS400Connection);
                    Adapter.Fill(dt);
                }

                ReturnValue = true;
            }
            catch (Exception ex)
            {
                try // ดักไว้อีกชั้น กรณี ไม่ได้กำหนด Member จึงไม่ได้สร้าง Alias ก็จะ Drop Alias ไม่ได้
                {
                    sqlDo = "DROP ALIAS Qtemp.getMember";
                    myCommand = new OleDbCommand(sqlDo, AS400Connection);
                    IntResult = myCommand.ExecuteNonQuery();
                }
                catch (Exception ex02)
                {
                }

                ReturnValue = false;
            }
            finally
            {
                AS400Connection.Dispose();
                AS400Connection.Close();
                AS400Connection = null;
            }

            return ReturnValue;
        }

        /// <summary>เช็คไฟล์จาก AS400 ว่ามีหรือไม่ โดยควบคุม Label ที่ใช้แสดงผล และแสดงสถานะ "พร้อม" หรือ "ไม่พร้อม" เป็นสี</summary>
    /// <param name="FileName">ชื่อไฟล์</param>
    /// <param name="Member">Member ของไฟล์</param>
    /// <param name="StatusLabel">ลาเบลที่จะให้แสดงผล</param>
        public static void CheckFileExistToLabel(string FileName, string Member, ref Label StatusLabel)
        {
            try
            {
                StatusLabel.Text = "ตรวจเช็ค";
                StatusLabel.ForeColor = Colors.Orange;
                StatusLabel.Refresh();
                Thread.Sleep(100);
                if (AS400FileExits(FileName, Member) == true)
                {
                    StatusLabel.Text = "พร้อม";
                    StatusLabel.ForeColor = Colors.Green;
                }
                else
                {
                    StatusLabel.Text = "ไม่พร้อม";
                    StatusLabel.ForeColor = Colors.Red;
                }
            }
            catch (Exception ex)
            {
                StatusLabel.Text = "ไม่พร้อม";
                StatusLabel.ForeColor = Colors.Red;
            }
        }

        public static string GetMemberLastWorkDate(DateTime DateInput = default(DateTime))
        {
            DateTime DateRet;
            if (DateInput == default(DateTime))
                DateInput = DateAndTime.Now;


            switch (DateInput.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    {
                        DateRet = DateInput.AddDays(-2);
                        break;
                    }

                case DayOfWeek.Monday:
                    {
                        DateRet = DateInput.AddDays(-3);
                        break;
                    }

                default:
                    {
                        DateRet = DateInput.AddDays(-1);
                        break;
                    }
            }

            string nDD;
            string nMM;
            string nThYY;
            nDD = DateRet.Day.ToString("00");
            nMM = DateRet.Month.ToString("00");
            nThYY = Conversions.ToString(CFncDateTime.ThaiYear(DateRet.Year, true));

            return "M" + nThYY + nMM + nDD;
        }
        // ########################################################################################################################
        // ####       Qry ออกมาเป็น ไฟล์            ####################
        // ########################################################################################################################

        /// <summary>ดึง FileName จาก AS400 ส่งข้อมูลออกมาเป็น Datatable และสามารเซฟไฟล์เป็น Xls ได้ด้วย (ใช้ OleDB)</summary>
    /// <param name="FileName">FileName จาก AS400</param>
    /// <param name="Member">*ออพชั่น Member ที่อยู่ใน FileName *ค่าเดิมเป็น Nothing</param>
    /// <param name="PathSaveFile">*ออพชั่น กำหนดตำแหน่งไฟล์และชื่อสกุลไฟล์  *ค่าเดิมเป็น Nothing</param>
        public static DataTable QryAS400ToXls(string FileName, string Member = null, string PathSaveFile = null)
        {
            var dt = new DataTable();
            // '## ตัวจับเวลา
            TimeSpan aDifference ;
            // Try
            dt = QryAS400ToDatatable(FileName, Member);
            // ###################################
            if (PathSaveFile != null)
            {
                var xlApp = new Excel.Application();
                Excel.Workbook wBook;
                var wSheet = new Excel.Worksheet();
                xlApp.DisplayAlerts = false;
                wBook = xlApp.Workbooks.Add();
                wSheet = (Excel.Worksheet)wBook.Worksheets[1];
                // xlApp.Visible = True
                CFncExcel.DataTableToExcelSheet(ref dt, ref wSheet);
                wBook.SaveAs(PathSaveFile, Excel.XlFileFormat.xlExcel9795);
                wBook.Close();
                xlApp.Quit();
            }
            // Catch ex As Exception
            // MsgBox(ex.Message)
            // Finally
            Thread.Sleep(500); // ถ้าไม่มี sleep มันจะใช้ EXCEL ตัวเดิมเปิด ต่อให้ปิดหน้าต่างใหม่  Process ก็ไม่ปิด
                           
            return dt;
        }

        /// <summary>ดึง FileName จาก AS400 เซฟไฟล์เป็น Xls (ใช้ ADODB.Recordset)</summary>
    /// <param name="FileName">FileName จาก AS400</param>
    /// <param name="MemberName">*ออพชั่น Member ที่อยู่ใน FileName *ค่าเดิมเป็น Nothing</param>
    /// <param name="PathSaveFileXls">*ออพชั่น กำหนดตำแหน่งไฟล์และชื่อสกุลไฟล์  *ค่าเดิมเป็น Nothing</param>
        public static object QryAS400ToXlsByADO(string FileName, string MemberName = null, string PathSaveFileXls = null)
        {
            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            var cn = new ADODB.Connection();
            cn.Open(ConnStr);
            string SqlSelect = string.Format("SELECT * FROM QS36F.\"{0}\"{1}", FileName, MemberName);

            // ###############################################
            // ####       ส่วนนี้ เป็นการ เซฟไฟล์ Xls             ####
            // ###############################################
            if (PathSaveFileXls != null)
            {
                var xlApp = new Excel.Application();
                Excel.Workbook wBook;
                var wSheet = new Excel.Worksheet();
                wBook = xlApp.Workbooks.Add();
                wSheet = (Excel.Worksheet)wBook.Worksheets[1];
                var rs2 = new ADODB.Recordset();
                rs2.Open(SqlSelect, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1);
                xlApp.ActiveCell.CopyFromRecordset(rs2);
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                wBook.SaveAs(PathSaveFileXls);
                wBook.Close(); // Excel จะปิดตัวเองพร้อมโปรแกรม
                xlApp.Quit(); // Excel จะปิดตัวเองพร้อมโปรแกรม
            }
            // Return dt 'ConvertRecordsetToDataSet(rs)
            // ' Return rs
            return 1;
        }
        /// <summary>ดึง FileName จาก AS400 เซฟไฟล์เป็น Xls (ใช้ ADODB.Recordset)</summary>
    /// <param name="FileName">FileName จาก AS400</param>
    /// <param name="MemberName">*ออพชั่น Member ที่อยู่ใน FileName *ค่าเดิมเป็น Nothing</param>
    /// <param name="PathSaveFileXls">*ออพชั่น กำหนดตำแหน่งไฟล์และชื่อสกุลไฟล์  *ค่าเดิมเป็น Nothing *ระบุนามสกุลมาด้วย xls หรือ xlsx</param>
        public static object QryAS400ToXlsByADOV2(string FileName, string MemberName = null, string PathSaveFileXls = null)
        {
            // Dim ServerIP = "192.10.10.10", Library = "QS36F", User = "pcs", Password = "pcu8"
            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

            var AS400Connection = new ADODB.Connection();
            AS400Connection.Open(ConnStr);
            object RecordsAffected;
            string SqlSelect;
            if (MemberName == null | string.IsNullOrEmpty(MemberName))
                SqlSelect = string.Format("SELECT * FROM QS36F.\"{0}\"{1}", FileName, MemberName);
            else
            {
                try
                {
                    // เคยเกิดเคส ว่า Already Exsist Alias เลยใส่กันไว้ 620528
                    string sqlDrop = "DROP ALIAS Qtemp.getMember";
                    AS400Connection.Execute(sqlDrop,out RecordsAffected);
                }
                catch (Exception ex)
                {
                }
                string sqlDo;
                sqlDo = "CREATE ALIAS Qtemp.getMember FOR " + AS400_Library + ".\"" + FileName + "\"(" + MemberName + ")";
                AS400Connection.Execute(sqlDo, out RecordsAffected);
                SqlSelect = string.Format("SELECT * FROM Qtemp.getMember");
            }


            // ###############################################
            // ####       ส่วนนี้ เป็นการ เซฟไฟล์ Xls เพื่อสำรอง    ####
            // ###############################################
            if (PathSaveFileXls != null)
            {
                var xlApp = new Excel.Application();
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                wBook = xlApp.Workbooks.Add();
                wSheet = (Excel.Worksheet)wBook.Worksheets[1];
                var rs2 = new ADODB.Recordset();
                rs2.Open(SqlSelect, AS400Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1);
                xlApp.ActiveCell.CopyFromRecordset(rs2);
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                // wBook.SaveAs(PathToSaveXls, Excel.XlFileFormat.xlExcel8)  คำสั่งจากตอนใช้ V.2013
                // wBook.SaveAs(PathSaveFileXls)
                bool XlsRoXlsx;
                if ((Path.GetExtension(PathSaveFileXls) ?? "") == ".xlsx")
                    XlsRoXlsx = true;
                else
                    XlsRoXlsx = false;
                CFncExcel.SaveAsWorkBookByVersion(xlApp, XlsRoXlsx, CFncFileFolder.GetFullPathWithoutExtension(PathSaveFileXls));
                wBook.Close(); // Excel จะปิดตัวเองพร้อมโปรแกรม
                xlApp.Quit(); // Excel จะปิดตัวเองพร้อมโปรแกรม
            }
            // Return dt 'ConvertRecordsetToDataSet(rs)
            // ' Return rs
            if (!(MemberName == null | string.IsNullOrEmpty(MemberName)))
            {
                string sqlDo = "DROP ALIAS Qtemp.getMember";

                AS400Connection.Execute(sqlDo, out RecordsAffected);
            }

            return 1;
        }



        /// <summary>
    /// นำไฟล์ (DataTable) ที่ดึงออกจาก AS400 ซึ่งมี คอลัมน์ เดียว มาแยกคอลัมน์ และ ทำการ กลับเครื่องหมายติดลบ ที่อยู่ทางขวามือ
    /// พร้อมกับ บันทึกไฟล์ โดยทำงาน บน Stream
    /// ทั้งหมด ไม่เพิ่งพา Excel หรือ ดาต้าเบส
    /// 620514 พบว่า record เป็นแสน ทำให้ช้า ได้พัฒนาเป็น V2 แล้ว
    /// </summary>
    /// <param name="DataTableInput">นำเข้าเป็น DataTable</param>
    /// <param name="Separator">กำหนด ตัวแยกคอลัมน์ ปกติใช้ ~</param>
    /// <param name="DoTrim">กำหนด การ Trim ซ้าย ขวา</param>
    /// <param name="DoMinusNeg">ต้องการ ทำการ กลับ -(ลบ) หรือไม่</param>
    /// <param name="PathSaveFile">กำหนดตำแหน่งบันทึกไฟล์ (ถ้าต้องการ ถ้าไม่ต้องการไม่ต้องกำหนด)</param>
    /// <returns>ออกเป็น DataTable ถ้าผิดพลาดจะออกเป็นตารางเปล่ว</returns>
        public static DataTable As400DataTableTextToColumn(ref DataTable DataTableInput, string Separator = "~", bool DoTrim = false, bool DoMinusNeg = false, string PathSaveFile = null)
        {
            var mStrm = new MemoryStream();
            // mStrm.Capacity = 1024
            long l;
            try
            {
                // 1. น้ำ ดาต้าเทเบิล เขียน เป็นไฟล์ (บน Memory Stream) โดยแยก แต่ละคอลัมน์ ด้วย "~"  
                // * ในกรณืที่เป็นไฟล์ จาก AS400 จะมีแค่ คอลัมน์ เดียว เท่านั้น และ ในคอลัมน์ นั้น ก็มี "~" อยู่แล้ว จะถูกแยกในชุดคำสั่งถัดไป
                //MemoryStream mStrm = new MemoryStream();
                string str = "";
                // Using StmW As New StreamWriter(mStrm, System.Text.Encoding.UTF8) '(FileNameSave, False, System.Text.Encoding.UTF8 )
                var StmW = new StreamWriter(mStrm, System.Text.Encoding.UTF8);
                foreach (DataRow drow in DataTableInput.Rows)
                {
                    // จริงๆแล้ว ไฟล์จาก AS400 มีแค่ คอลัมน์ เดียว แต่ใส่่เครื่องหมาย | ไว้เผื่อ
                    string lineoftext = string.Join("|", drow.ItemArray.Select(s => s.ToString()).ToArray());
                    StmW.WriteLine(lineoftext);
                    l = mStrm.Length;
                    str += lineoftext;
                }
                StmW.Flush();  // ทำการ Flush เพื่อ ทำการปิดไฟล์ ถ้าไม่ Flush ประกฏว่า  WriteLine ได้แค่ 400 Line
                               // StmW.Close()
                               // End Using

                // ###############################################################
                // 2. กำหนด จุดเริ่มต้น ในการอ่าน Memory Stream 
                // อ่านด้วย(StreamReader)
                mStrm.Seek(0, SeekOrigin.Begin);
                var StmRd = new StreamReader(mStrm, System.Text.Encoding.UTF8);
                // StmRd.Close()

                // ###############################################################
                // 3. ทำการ อ่าน นำ แต่ละบรรทัด มาแยกเป็น คอลัมน์ ด้วย Separator
                var DtRet = new DataTable();
                // 3.1. ทำการ สร้าง Column ก่อน สำรองไว้ 100
                for (int nC = 0; nC <= 500; nC++)
                    DtRet.Columns.Add("Col" + Conversions.ToString(nC));

                // 3.2. ทำการ อ่านไฟล์ ทีละบรรทัดแล้ว แยก เป็น Array ก่อนที่จะ นำบรรจุใส่ Row ให้ DataTable
                // พร้อมกับ ทำการเก็บค่า ดูว่า ใช้ คอลัมน์ ทั้งหมด กี่คอลัมน์ เพื่อที่จะลบ คอลัมน์ ที่ไม่ได้ใช้ ออก
                int TopDataCount = 0; // เก็บ จำนวนสูงสุด ของ คอลัมน์ เพื่อลบคอลัมน์ ส่วนเกินออก

                //string newThisLine = "";
                while (!StmRd.EndOfStream)
                {
                    string THisLine = StmRd.ReadLine();
                    var DataAry = Strings.Split(THisLine, Separator);

                    if (DoMinusNeg == true | DoTrim == true)
                    {
                        for (int i = 0, loopTo = DataAry.Count() - 1; i <= loopTo; i++)
                        {
                            // If Trim(DataAry(i)) = "24ศ025" Then  'for debug
                            // DataAry(i) = DataAry(i)
                            // End If
                            if (DoTrim == true)
                                DataAry[i] = Strings.Trim(DataAry[i]);
                            if (DoMinusNeg == true & (Strings.Right(DataAry[i], 1) ?? "") == "-" & Information.IsNumeric(DataAry[i]))
                                // เงื่อนไขนี้ ทำงานได้ โดยไม่ต้องเพิ่งข้างบน ไม่สนใจว่าจะ - ซ้านสุดหรือไม่
                                DataAry[i] = Conversions.ToDouble(DataAry[i]).ToString();
                        }
                    }

                    DtRet.Rows.Add(DataAry);
                    TopDataCount = Conversions.ToInteger(Interaction.IIf(DataAry.Count() > TopDataCount, DataAry.Count(), TopDataCount));
                }
                // ###############################################################
                // 4.หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป เนื่องจากเราสร้างคอลัมน์ ไว้สำรอง 100 คอลัมน์
                for (int rowBack = DtRet.Columns.Count - 1, loopTo1 = TopDataCount; rowBack >= loopTo1; rowBack += -1)
                    DtRet.Columns.Remove(DtRet.Columns[rowBack]);
                // ###############################################################
                // ###############################################################
                // 4.001 ทำการ ใส่ '' "ค่าว่าง" ไปแทน Cell ที่เป็น Null เพราะ ใน DataBase จะเป็น Null ทำให้ระบบ ดูข้อมูลยาก 
                // เทียบกับ table ที่ได้จาก EXCEL จะเป็นค่าว่าง  ฉนั่นต้องปรับให้มีค่าเหมือนกัน
                foreach (DataRow drow in DtRet.Rows)
                {
                    for (int nCol = 0, loopTo2 = DtRet.Columns.Count - 1; nCol <= loopTo2; nCol++)
                    {
                        var CellVal = drow[nCol];
                        if (Information.IsDBNull(CellVal))
                            drow[nCol] = "";
                    }
                }
                // Dim foundRows() As Data.DataRow
                // foundRows = DtRet.Select("CompanyName Like 'A%'")
                // ###############################################################
                // 5.'ทำการ บันทึกไฟล์ ด้วย ถ้าต้องการ
                if (PathSaveFile != null)
                {
                    using (var StWFile = new StreamWriter(PathSaveFile, false, System.Text.Encoding.UTF8))
                    {
                        foreach (DataRow drow in DtRet.Rows)
                        {
                            string lineoftext = string.Join(Separator, drow.ItemArray.Select(s => s.ToString()).ToArray());
                            StWFile.WriteLine(lineoftext);
                        }
                    }
                }

                // ทำการปิด Sesion ต่างๆ
                // mStrm.Close() '(ปิดไปแล้ว)
                StmW.Close();
                StmRd.Close();

                return DtRet;
            }
            // ###############################################################
            // Return (1) 'File(mStrm, "text/plain", "CompetitionEntries.csv")
            catch (Exception ex)
            {
                // MsgBox(ex.Message)
                Interaction.Beep();
                var dtr = new DataTable();
                dtr.Columns.Add("A"); dtr.Rows.Add("ผิดพลาด : " + ex.Message);
                return dtr;
            }// New DataTable
        }

      

        /// <summary>
    /// นำไฟล์ (DataTable) ที่ดึงออกจาก AS400 ซึ่งมี คอลัมน์ เดียว มาแยกคอลัมน์ และ ทำการ กลับเครื่องหมายติดลบ ที่อยู่ทางขวามือ
    /// พร้อมกับ บันทึกไฟล์ โดยทำงาน บน Stream
    /// ทั้งหมด ไม่เพิ่งพา Excel หรือ ดาต้าเบส
    /// 620514 พัฒนาเป็น V2 เพิ่มความเร็ว ไม่ใช้ Memory Stream แล้ว เนื่องจากเมื่อ record มี เป็น แสน Row จะทำให้ช้าอย่างเห็นได้ชัด
    /// </summary>
    /// <param name="DataTableInput">นำเข้าเป็น DataTable</param>
    /// <param name="Separator">กำหนด ตัวแยกคอลัมน์ ปกติใช้ ~</param>
    /// <param name="DoTrim">กำหนด การ Trim ซ้าย ขวา</param>
    /// <param name="DoMinusNeg">ต้องการ ทำการ กลับ -(ลบ) หรือไม่</param>
    /// <param name="PathSaveFile">กำหนดตำแหน่งบันทึกไฟล์ (ถ้าต้องการ ถ้าไม่ต้องการไม่ต้องกำหนด)</param>
    /// <returns>ออกเป็น DataTable ถ้าผิดพลาดจะออกเป็นตารางเปล่ว</returns>
        public static DataTable As400DataTableTextToColumnV2(DataTable DataTableInput, string Separator = "~", bool DoTrim = false, bool DoMinusNeg = false, string PathSaveFile = null)
        {

            // ###############################################################
            // ทำการ สร้าง Column ก่อน สำรองไว้ 500
            var DtTxt2Col = new DataTable();
            for (int nC = 0; nC <= 500; nC++)
                DtTxt2Col.Columns.Add("Col" + Conversions.ToString(nC));
            int TopCountCol = 0; // เก็บ จำนวนสูงสุด ของ คอลัมน์ เพื่อลบคอลัมน์ ส่วนเกินออก
                                 // ###############################################################
                                 // Text To Col พร้อม Negative minus
            //string newThisLine = "";
            foreach (DataRow drow in DataTableInput.Rows)
            {
                string THisLine = drow[0] as string;
                string[] DataAry = Strings.Split(THisLine, Separator);

                if (DoMinusNeg == true | DoTrim == true)
                {
                    for (int iCol = 0, loopTo = DataAry.Count() - 1; iCol <= loopTo; iCol++)
                    {
                        if (DoTrim == true)
                            DataAry[iCol] = Strings.Trim(DataAry[iCol]);
                        // เงื่อนไขนี้ ทำให้เหมือน ของ Excel เพราะ จะทำเฉพาะ เครื่องหมาย - อยู่ขวาสุด
                        if (DoMinusNeg == true)
                        {
                            if ((Strings.Right(DataAry[iCol].ToString(), 1) ?? "") == "-")
                            {
                                if (Information.IsNumeric(DataAry[iCol].ToString()))
                                    // เงื่อนไขนี้ ทำงานได้ โดยไม่ต้องเพิ่งข้างบน ไม่สนใจว่าจะ - ซ้ายสุดหรือไม่
                                    DataAry[iCol] = Conversions.ToDouble(DataAry[iCol]).ToString();
                            }
                        }
                    }
                }

                DtTxt2Col.Rows.Add(DataAry);
                TopCountCol = Conversions.ToInteger(Interaction.IIf(DataAry.Count() > TopCountCol, DataAry.Count(), TopCountCol));
            }
            // ###############################################################
            // หลังจากใส่ข้อมูลเสร็จและได้ ทราบว่าใช้ทั้งหมด กี่ คอลัมน์ ก็จะ ลบ คอลัมน์ ที่ไม่มีข้อมูลออก ทิ้งไป เนื่องจากเราสร้างคอลัมน์ ไว้สำรอง 500 คอลัมน์
            for (int rowBack = DtTxt2Col.Columns.Count - 1, loopTo1 = TopCountCol; rowBack >= loopTo1; rowBack += -1)
                DtTxt2Col.Columns.Remove(DtTxt2Col.Columns[rowBack]);

            // ###############################################################
            // ทำการ ใส่ '' "ค่าว่าง" ไปแทน Cell ที่เป็น Null เพราะ ใน DataBase จะเป็น Null ทำให้ระบบ ดูข้อมูลยาก 
            // เทียบกับ table ที่ได้จาก EXCEL จะเป็นค่าว่าง  ฉนั่นต้องปรับให้มีค่าเหมือนกัน
            foreach (DataRow drow in DtTxt2Col.Rows)
            {
                for (int nCol = 0, loopTo2 = DtTxt2Col.Columns.Count - 1; nCol <= loopTo2; nCol++)
                {
                    var CellVal = drow[nCol];
                    if (Information.IsDBNull(CellVal))
                        drow[nCol] = "";
                }
            }

            // ###############################################################
            // ทำการ บันทึกไฟล์ ด้วย ถ้าต้องการ
            if (PathSaveFile != null)
            {
                using (var StWFile = new StreamWriter(PathSaveFile, false, System.Text.Encoding.UTF8))
                {
                    foreach (DataRow drow in DtTxt2Col.Rows)
                    {
                        string lineoftext = string.Join(Separator, drow.ItemArray.Select(s => s.ToString()).ToArray());
                        StWFile.WriteLine(lineoftext);
                    }
                }
            }

            // ###############################################################
            return DtTxt2Col;
        }



        public static object TestQryAS400() // ByVal FileName As String, ByVal MemberName As String) As DataTable
        {
            object ObjResult;
            var dt = new DataTable();

            string ConnStr = "Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=" + AS400_User + ";Password=" + AS400_Password + ";Data Source=" + AS400_ServerIP + ";Force Translate=0;Catalog Library List=" + AS400_Library + ";SSL=DEFAULT;";

 
            string sql = "CREATE ALIAS QTEMP.MYMBRB FOR QS36F.WZBMJUM (MBRB)";

            // Dim Sql As String = String.Format("SELECT * FROM {0}({1})", FileName, MemberName)
            var AS400Connection = new OleDbConnection(ConnStr);
            // Try
            AS400Connection.Open();

            var OleCmd = new OleDbCommand(sql, AS400Connection);
            OleCmd.CommandTimeout = 600;
            ObjResult = OleCmd.ExecuteScalar();

            sql = "SELECT * FROM QTEMP.MYMBRB";
            var Adapter = new OleDbDataAdapter(sql, AS400Connection);
            Adapter.Fill(dt); // (DataSet)
            ObjResult = dt;
            
            AS400Connection.Close();
            // End Try
            return ObjResult;
        }



        /// <summary>  Qry ข้อมูลออกมาจาก AS400 => TextToColumn => SaveFile => Import To SQL Server 's Table
    /// / พร้อม เสริม ฟิลด์ ข้อมูลที่จะใส่เพิ่มไปใน Table
    /// </summary>
    /// <param name="FileName">ชื่อไฟล์ใน AS400 ที่ต้องการ</param>
    /// <param name="Member">Member ของ File AS400</param>
    /// <param name="Separator">ตัวขั่นคอลัมน์</param>
    /// <param name="ConnString">Connection String ของ SQL Server ปลายทาง</param>
    /// <param name="ToTableName">ชื่อตาราง ปลายทาง</param>
    /// <param name="PathSaveFile">ตำแหน่งเซฟไฟล์ ถ้าต้องการ Backup ให้ใส่ Parameter</param>
    /// <param name="ArL_AttrNameAndVal">ชื่อ Attribute และ Value ควรประกาศเป็น myArrayList.Add({"AttrName", "AttrVal"}</param>
    /// <remarks>เรียกใช้ ฟังก์ชั่น สำเร็จรูปอื่นๆ</remarks>
        public static void ImportQryAS400ToTableDataBase_WithFileAttribute(string FileName, string Member, string Separator
                                                            , string ConnString, string ToTableName
                                                            , string PathSaveFile = null
                                                            , ArrayList ArL_AttrNameAndVal = null
                                                            , bool WantColumnNameTableByDataTable = false
    )
        {
            DataTable dtToImport = null;
            string AS400_FileName = FileName;

            dtToImport = QryAS400ToDatatableV2(AS400_FileName, Member);
            // ###############################################



            if (dtToImport.Columns.Count == 1 & Separator != null)
            {
                // มี Col เดียว เพราะเพิ่ง Qry มาจาก AS400 ใหม่ๆ หรือ เป็น DT มาจากไฟล์ W Excel ที่ยังไม่ TextToCol 
                dtToImport = As400DataTableTextToColumn(ref dtToImport, Separator, true, true);
                dtToImport = CFncDataTable.DatatableTrimCell(dtToImport);
            }
            else
                // ไม่ต้อง TxtToCol เซฟไฟล์เก็บไว้อย่างเดียว
                dtToImport = CFncSave_LoadGridFile.Trim_DataTable(dtToImport);
            if (PathSaveFile != null)
            {
                PathSaveFile = CFncFileFolder.NewFileNameUnique(PathSaveFile);
                CFncSave_LoadGridFile.DataTableSaveToTxtFile1(ref dtToImport, PathSaveFile, Separator);
            }
            // ###############################################
            // วนสร้าง Column ที่เป็นของ Attribute
            if (ArL_AttrNameAndVal != null)
            {
                for (int nAttr = 0, loopTo = ArL_AttrNameAndVal.Count - 1; nAttr <= loopTo; nAttr++)
                {
                    string AttributeName = (string)ArL_AttrNameAndVal[nAttr];
                    string AttributeVal = (string)ArL_AttrNameAndVal[nAttr];

                    // Dim ColName As String = AttributeName

                    var dtCol_Att = new DataColumn() { ColumnName = AttributeName, DataType = typeof(string), DefaultValue = AttributeVal };

                    dtCol_Att.ReadOnly = false;
                    dtToImport.Columns.Add(dtCol_Att);
                    dtToImport.Columns[dtCol_Att.ColumnName].SetOrdinal(nAttr);
                }
            }



            // ###############################################
            var SqlServer = new ClsMsSql(ConnString); // (txt_Host.Text, txt_User.Text, txt_Passw.Text, txt_Database.Text)
                                                                  // ## นำ ข้อมูล As400 เข้า DataBase
                                                                  // สร้างคำสั่งสร้างตารางสำหรับ นำ ไฟล์ As400 เข้า DataBase
            string Imp_TableName = ToTableName;

            string SqlCrtImptTable;
            if (WantColumnNameTableByDataTable == true)
                SqlCrtImptTable = FncDataBaseTool.GenCreateTableByDataTableImport(Imp_TableName, dtToImport);
            else
                SqlCrtImptTable = FncDataBaseTool.GenCreateTableImport(Imp_TableName, dtToImport.Columns.Count);

            // สร้างตาราง ตรวจสอบว่ามีตารางหรือไม่แล้ว ถ้ามีให้ลบก่อนสร้าง
            if (SqlServer.TableExists(Imp_TableName))
                SqlServer.DeleteTable(Imp_TableName);
            SqlServer.ExecuteNonQuery(SqlCrtImptTable);

            // ###################
            SqlServer.CopyDatatableToDatabaseTable(dtToImport, Imp_TableName);

            // MsgBox("บันทึกไฟล์เข้าฐานข้อมูลแล้ว")
            Interaction.Beep();
            SqlServer.CloseConnection();
        }
    }
}
