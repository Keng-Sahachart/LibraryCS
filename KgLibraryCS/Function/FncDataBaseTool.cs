using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
using System;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace kgLibraryCs
{
    public static class FncDataBaseTool
    {
        /// <summary> เป็นการ เรียกใช้งาน Store Procedure สามารถใส่ Parameter ได้ 10 ตัว ParaName1-10 และ Val1-10</summary>
    /// <param name="DataBaseConnection">Connection ของ OleDbConnection ที่ทำการเชื่อมต่อแล้ว</param>
    /// <param name="StoreName">ชื่อของ Store Procedure ที่ต้องการเรียกใช้</param>
    /// <remarks>อะรูไม่ไร้</remarks>
        public static OleDbDataReader ExcuteStoreProcedure(ref OleDbConnection DataBaseConnection
                                      , string StoreName
                                      , string ParaName1 = null, string Val1 = null
                                      , string ParaName2 = null, string Val2 = null
                                      , string ParaName3 = null, string Val3 = null
                                      , string ParaName4 = null, string Val4 = null
                                      , string ParaName5 = null, string Val5 = null
                                      , string ParaName6 = null, string Val6 = null
                                      , string ParaName7 = null, string Val7 = null
                                      , string ParaName8 = null, string Val8 = null
                                      , string ParaName9 = null, string Val9 = null
                                      , string ParaName10 = null, string Val10 = null)
        {
            var cmd = new OleDbCommand(StoreName);

            if (ParaName1 != null)
                cmd.Parameters.AddWithValue(ParaName1, Val1);
            else
                goto Exit_AddParameter; // ป้องกันการ If หลายที เกินความจำเป็น อาจทำให้โปรแกรมช้า
            if (ParaName2 != null)
                cmd.Parameters.AddWithValue(ParaName2, Val2);
            else
                goto Exit_AddParameter;
            if (ParaName3 != null)
                cmd.Parameters.AddWithValue(ParaName3, Val3);
            else
                goto Exit_AddParameter;
            if (ParaName4 != null)
                cmd.Parameters.AddWithValue(ParaName4, Val4);
            else
                goto Exit_AddParameter;
            if (ParaName5 != null)
                cmd.Parameters.AddWithValue(ParaName5, Val5);
            else
                goto Exit_AddParameter;
            if (ParaName6 != null)
                cmd.Parameters.AddWithValue(ParaName6, Val6);
            else
                goto Exit_AddParameter;
            if (ParaName7 != null)
                cmd.Parameters.AddWithValue(ParaName7, Val7);
            else
                goto Exit_AddParameter;
            if (ParaName8 != null)
                cmd.Parameters.AddWithValue(ParaName8, Val8);
            else
                goto Exit_AddParameter;
            if (ParaName9 != null)
                cmd.Parameters.AddWithValue(ParaName9, Val9);
            else
                goto Exit_AddParameter;
            if (ParaName10 != null)
                cmd.Parameters.AddWithValue(ParaName10, Val10);
            Exit_AddParameter:
            ;


            // Dim reader As OleDbDataReader
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = DataBaseConnection;

            // reader = cmd.ExecuteReader()
            return cmd.ExecuteReader();
        }
        public static DataTable QryTableToDataTable(ref OleDbConnection OleDbCon, string sqlSelect) // ฟังชั้นนี้อาจมีการเปลี่ยนแปลง เพราะ ใช้ oledb
        {
            var myDataAdapter = new OleDbDataAdapter(sqlSelect, OleDbCon);
            var myDataTable = new DataTable();
            myDataTable.Clear();
            myDataAdapter.Fill(myDataTable);
            return myDataTable;
        }

        public static DataTable QryTableToDataTable(string OleDbConnectionString, string sqlSelect) // ฟังชั้นนี้อาจมีการเปลี่ยนแปลง เพราะ ใช้ oledb
        {
            var Connection = new OleDbConnection();
            var adapter = new OleDbDataAdapter();
            Connection = new OleDbConnection(OleDbConnectionString);
            Connection.Open();
            var myDataAdapter = new OleDbDataAdapter(sqlSelect, Connection);
            var myDataTable = new DataTable();
            myDataTable.Clear();
            myDataAdapter.Fill(myDataTable);
            return myDataTable;
        }
        /// <summary>
    /// สร้าง คำสั่ง สำหรับ ค้นหา จากตารางได้ เพียงแค่ ใส่ ฟิลด์ที่ต้องการโชว์ ฟิลด์ที่ต้องการค้นหาคำ คำที่ต้องการ
    /// </summary>
    /// <param name="FieldShow">ฟิลด์ที่ต้องการดึงออกมา</param>
    /// <param name="FieldSearch">ฟิลด์ที่ต้องการค้นหาคำ</param>
    /// <param name="WordsSearch">คำที่ต้องการค้นหาในฟิลด์ สามารถกำหนด Option ได้</param>
    /// <param name="TableName">ตาราง</param>
    /// <param name="WantAllWord">กำหนดว่า ต้องการค้นหาให้ตรงกับทุกตัวอักษรที่ค้นหา หรือ แบ่งหาเป็นคำ โดยแบ่งตามช่องว่าง</param>
    /// <returns>ได้คำสั่ง SELECT สำหรับนำไป Query อีกที</returns>
        public static string CreateSqlForSearch(string FieldShow, string FieldSearch, string WordsSearch
                                    , string TableName, bool WantAllWord = false
                                    , string WhereConditionPlus = "")
        {
            string sql = string.Format("Select {0} From {1}", FieldShow, TableName);

            // ############################################################
            // ตัวอย่าง Where Condition ทั้งหมดที่ต้องการ
            // (f1=d1,.....,fn=d1) or/and (f1=d2,.....,fn=d2) ........ or/and (f1=dx-1,.....,fn=dx-1) or/and (f1=dx,.....,fn=dx)


            // ###############################################################
            // #####  สร้าง field สำหรับ Where ทั้งหมด  (f1=d1,.....,fn=d1)   #####
            // ###############################################################
            string where = "(";
            foreach (string Str in FieldSearch.Split(','))
                where += Str + " like '%{0}%' or ";
            where = (where + ")").Replace("or )", " )");

            // #######################################################################################
            // ##### นำ where มารวมเข้ากับ คำที่ต้องการค้นหา  ด้วย logic WantAllWord ว่าต้องการทุกคำที่ระบุหรือไม่ ######
            // #######################################################################################
            string OrAndCondition;
            if (WantAllWord == false)
                OrAndCondition = " or ";
            else
                OrAndCondition = " and ";

            string AllwhereCondition = "(";
            foreach (string WordToSearch in WordsSearch.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                AllwhereCondition += string.Format(where, WordToSearch) + OrAndCondition;
            AllwhereCondition = (AllwhereCondition + ")").Replace(OrAndCondition + ")", ")");
            // If WordsSearch.Length > 1 Or WhereConditionPlus.Length > 1 Then
            // sql = sql + " where " + AllwhereCondition
            // If WhereConditionPlus.Length > 1 Then
            // sql &= " and " & WhereConditionPlus
            // End If
            // End If
            if (WordsSearch.Length > 0 | WhereConditionPlus.Length > 0)
            {
                sql = sql + " where ";
                if (WordsSearch.Length > 0)
                    sql = sql + AllwhereCondition;
                if (WhereConditionPlus.Length > 0)
                {
                    if (WordsSearch.Length > 0)
                        sql += " and ";
                    sql += WhereConditionPlus;
                }
            }
            return sql;
        }
        /// <summary>
        /// สำหรับ String ให้ใส่ Single Quote ครอบ 
        /// และป้องกัน ใน String มี Single Quote เป็น data อยู่ด้วย
        /// </summary>
        /// <param name="StrData">String Data</param>
        /// <returns>'String Data'</returns>
        public static string SqlEsc(string StrData)
        {
            StrData = "N'" + StrData.Replace("'", "''") + "'";
            return StrData;
        }

        public static string CreateConnectionStringOled(string Host, string UserName, string Password, string DataBaseName)
        {
            string Connstring = "Provider=SQLOLEDB.1;Data Source={0};Initial Catalog={1};Persist Security Info=False;User ID={2} ;password={3}";
            return string.Format(Connstring, Host, DataBaseName, UserName, Password);
        }
        public static string CreateConnectionStringSQLclient(string Host, string UserName, string Password, string DataBaseName)
        {
            string Connstring = "Data Source={0};Initial Catalog={1};User ID={2};Password={3}";
            return string.Format(Connstring, Host, DataBaseName, UserName, Password);
        }


        /// <summary>เอาทั้ง DataTable ยัดเข้าไปใน Table ใน Data Base โดยอย่างน้อย Column ต้องเท่ากัน</summary>
        public static void CopyDatatableToDatabaseTable(DataTable dt, SqlConnection SqlConn, string TagetTable)
        {

            // Conn.Open()

            var bulkcopy = new SqlBulkCopy(SqlConn);
            bulkcopy.DestinationTableName = TagetTable;
            bulkcopy.WriteToServer(dt);
        }
        /// <summary>เอาทั้ง DataTable ยัดเข้าไปใน Table ใน Data Base โดยอย่างน้อย Column ต้องเท่ากัน</summary>
        public static void CopyDatatableToDatabaseTable(DataTable dt, string SqlConnStr, string TagetTable)
        {
            var Conn = new SqlConnection(SqlConnStr);
            Conn.Open();
            CopyDatatableToDatabaseTable(dt, Conn, TagetTable);
            Conn.Close();
        }


        public enum ConnectionType
        {
            OledDB,
            SQLclient,
            ADO
        }
        public static string GenConnString(string Host, string UserName, string Password, string DataBaseName, ConnectionType ConnectionStringType)
        {
            string ConnectionString = null;
            switch (ConnectionStringType)
            {
                case ConnectionType.SQLclient:
                    {
                        ConnectionString = "Data Source={0};Initial Catalog={1};Persist Security Info=True;User ID={2};Password={3};Network Library=DBMSSOCN;Connect Timeout=15;MultipleActiveResultSets=true";
                        ConnectionString = string.Format(ConnectionString, Host, DataBaseName, UserName, Password);
                        break;
                    }

                case ConnectionType.OledDB:
                    {
                        ConnectionString = "Provider=SQLOLEDB.1;Data Source={0};Initial Catalog={1};Persist Security Info=False;User ID={2} ;password={3};Network Library=DBMSSOCN";
                        ConnectionString = string.Format(ConnectionString, Host, DataBaseName, UserName, Password);
                        break;
                    }
            }
            return ConnectionString;
        }



        /// <summary> สร้างตาราง สำหรับ Import ข้อมูล </summary>
    /// <param name="TableName">กำหนดชื่อตาราง</param>
    /// <param name="NumberColumn">จำนวนคอลลัมน์</param>
        public static string GenCreateTableImport(string TableName, int NumberColumn)
        {
            string sql = "   CREATE TABLE [{0}] ( {1} [Id] [int] IDENTITY(1,1) NOT NULL) ON [PRIMARY]";
            string sqlCol = "";
            for (int n = 1, loopTo = NumberColumn; n <= loopTo; n++)
                sqlCol += string.Format("[Col{0}] [nvarchar](500) NULL ,", n.ToString("000"));
            sql = string.Format(sql, TableName, sqlCol);
            return sql;
        }

        /// <summary> สร้างสคริป สำหรับ สร้างตาราง เพื่อ Import ข้อมูล
    /// </summary>
    /// <param name="TableName">กำหนดชื่อตาราง</param>
    /// <param name="dataTableForGenSqlCreate">DataTable ที่ต้องการนำชื่อมาสร้าง Table ใน SQL Server</param>
        public static string GenCreateTableByDataTableImport(string TableName, DataTable dataTableForGenSqlCreate)
        {
            string sql = "   CREATE TABLE [{0}] ";
            sql += " ( ";
            sql += " {1} "; // , ไม่ต้องมี เพราะอยู่ใน Loop
            sql += " [Id] [int] IDENTITY(1,1) NOT NULL ";  // Id ต้องอยู่หลังสุด เพราะคำสั่งนี้จะใช้กับ SqlBulkCopy
            sql += " ) ON [PRIMARY] ";

            string sqlCol = "";
            for (int n = 0, loopTo = dataTableForGenSqlCreate.Columns.Count - 1; n <= loopTo; n++)
            {
                string ColName = dataTableForGenSqlCreate.Columns[n].ColumnName;
                sqlCol += string.Format("[{0}] [nvarchar](500) NULL ,", ColName);
            }
            sql = string.Format(sql, TableName, sqlCol);
            return sql;
        }


        // ############################################################################################# 
        // https://www.codeproject.com/Articles/10503/Simplest-code-to-convert-an-ADO-NET-DataTable-to-a
        private static ADODB.DataTypeEnum TranslateType(Type columnType)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    {
                        return ADODB.DataTypeEnum.adBoolean;
                    }

                case "System.Byte":
                    {
                        return ADODB.DataTypeEnum.adUnsignedTinyInt;
                    }

                case "System.Char":
                    {
                        return ADODB.DataTypeEnum.adChar;
                    }

                case "System.DateTime":
                    {
                        return ADODB.DataTypeEnum.adDate;
                    }

                case "System.Decimal":
                    {
                        return ADODB.DataTypeEnum.adCurrency;
                    }

                case "System.Double":
                    {
                        return ADODB.DataTypeEnum.adDouble;
                    }

                case "System.Int16":
                    {
                        return ADODB.DataTypeEnum.adSmallInt;
                    }

                case "System.Int32":
                    {
                        return ADODB.DataTypeEnum.adInteger;
                    }

                case "System.Int64":
                    {
                        return ADODB.DataTypeEnum.adBigInt;
                    }

                case "System.SByte":
                    {
                        return ADODB.DataTypeEnum.adTinyInt;
                    }

                case "System.Single":
                    {
                        return ADODB.DataTypeEnum.adSingle;
                    }

                case "System.UInt16":
                    {
                        return ADODB.DataTypeEnum.adUnsignedSmallInt;
                    }

                case "System.UInt32":
                    {
                        return ADODB.DataTypeEnum.adUnsignedInt;
                    }

                case "System.UInt64":
                    {
                        return ADODB.DataTypeEnum.adUnsignedBigInt;
                    }

                default:
                    {
                        return ADODB.DataTypeEnum.adVarChar;
                    }
            }
        }

        public static ADODB.Recordset ConvertToRecordset(DataTable inTable)
        {
            var result = new ADODB.Recordset();
            result.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
            ADODB.Fields resultFields = result.Fields;
            var inColumns = inTable.Columns;
            foreach (DataColumn inColumn in inColumns)
                 //resultFields.Append(inColumn.ColumnName, TranslateType(inColumn.DataType), inColumn.MaxLength
            //, If(inColumn.AllowDBNull, ADODB.FieldAttributeEnum.adFldIsNullable, ADODB.FieldAttributeEnum.adFldUnspecified)
            //, Nothing)

                resultFields.Append(inColumn.ColumnName, TranslateType(inColumn.DataType), inColumn.MaxLength, inColumn.AllowDBNull ? ADODB.FieldAttributeEnum.adFldIsNullable : ADODB.FieldAttributeEnum.adFldUnspecified, null);
            result.Open(System.Reflection.Missing.Value, System.Reflection.Missing.Value, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, 0);

            foreach (DataRow dr in inTable.Rows)
            {
                result.AddNew(System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                for (int columnIndex = 0, loopTo = inColumns.Count - 1; columnIndex <= loopTo; columnIndex++)
                    resultFields[columnIndex].Value  = dr[columnIndex];
            }
            return result;
        }
        // ############################################################################################# 

        /// <summary>  Qry ข้อมูลออกมาจาก AS400 => TextToColumn => SaveFile => Import To SQL Server 's Table  </summary>
    /// <param name="FileName">ชื่อไฟล์ใน AS400 ที่ต้องการ</param>
    /// <param name="Member">Member ของ File AS400</param>
    /// <param name="Separator">ตัวขั่นคอลัมน์</param>
    /// <param name="ConnString">Connection String ของ SQL Server ปลายทาง</param>
    /// <param name="ToTableName">ชื่อตาราง ปลายทาง</param>
    /// <param name="PathSaveFile">ตำแหน่งเซฟไฟล์ ถ้าต้องการ Backup ให้ใส่ Parameter</param>
    /// <remarks>เรียกใช้ ฟังก์ชั่น สำเร็จรูปอื่นๆ</remarks>
        public static void ImportQryAS400ToTableDataBase(string FileName, string Member, string Separator, string ConnString, string ToTableName, string PathSaveFile = null, bool AppendData = false)
        {
            DataTable dtToImport = null;
            string AS400_FileName = FileName;

            dtToImport = FncAS400.QryAS400ToDatatableV2(AS400_FileName, Member);
            // ###############################################

            // ###############################################

            if (dtToImport.Columns.Count == 1 & Separator != null)
            {
                // มี Col เดียว เพราะเพิ่ง Qry มาจาก AS400 ใหม่ๆ หรือ เป็น DT มาจากไฟล์ W Excel ที่ยังไม่ TextToCol 
                dtToImport = FncAS400.As400DataTableTextToColumn(ref dtToImport, Separator, true, true);
                dtToImport = FncDataTable.DatatableTrimCell(dtToImport);
            }
            else
                // ไม่ต้อง TxtToCol เซฟไฟล์เก็บไว้อย่างเดียว
                dtToImport = FncSave_LoadGridFile.Trim_DataTable(dtToImport);
            if (PathSaveFile != null)
            {
                PathSaveFile = FncFileFolder.NewFileNameUnique(PathSaveFile);
                FncSave_LoadGridFile.DataTableSaveToTxtFile1(ref dtToImport, PathSaveFile, Separator);
            }
            // ###################
            var SqlServer = new MsSql_Manager(ConnString); // (txt_Host.Text, txt_User.Text, txt_Passw.Text, txt_Database.Text)
                                                                  // ## นำ ข้อมูล As400 เข้า DataBase
                                                                  // สร้างคำสั่งสร้างตารางสำหรับ นำ ไฟล์ As400 เข้า DataBase
            string Imp_TableName = ToTableName;

            string SqlCrtImptTable = GenCreateTableImport(Imp_TableName, dtToImport.Columns.Count);


            if (AppendData == false)
            {
                // สร้างตาราง ตรวจสอบว่ามีตารางหรือไม่แล้ว ถ้ามีให้ลบก่อนสร้าง
                if (SqlServer.TableExists(Imp_TableName) == true)
                {
                    SqlServer.DeleteTable(Imp_TableName);
                }
                SqlServer.ExecuteNonQuery(SqlCrtImptTable);
            }
            // ###################
            SqlServer.CopyDatatableToDatabaseTable(dtToImport, Imp_TableName);

            // MsgBox("บันทึกไฟล์เข้าฐานข้อมูลแล้ว")
            Interaction.Beep();
            SqlServer.CloseConnection();
        }

        /// <summary>
    /// Qry ข้อมูลออกมาจาก AS400 => TextToColumn => SaveFile => Import To SQL Server 's Table
    /// พร้อม ข้อมูลของไฟล์ ที่ได้จาก เว็บคุณ Turbo
    /// อัพเดตจาก ฟังก์ชั่น ImportQryAS400ToTableDataBase
    /// </summary>
    /// <param name="FileName">ชื่อไฟล์ใน AS400 ที่ต้องการ</param>
    /// <param name="Member">Member ของ File AS400</param>
    /// <param name="Separator">ตัวขั่นคอลัมน์</param>
    /// <param name="ConnString">Connection String ของ SQL Server ปลายทาง</param>
    /// <param name="ToTableName">ชื่อตาราง ปลายทาง</param>
    /// <param name="PathSaveFile">ตำแหน่งเซฟไฟล์ ถ้าต้องการ Backup ให้ใส่ Parameter</param>
    /// <param name="NewTbOrCollec">True =>ลบ Table แล้วสร้างใหม่ หรือ False =>ใช้ Table เดิมสะสมไปเรื่อยๆ</param>
    /// <remarks>เรียกใช้ ฟังก์ชั่น สำเร็จรูปอื่นๆ</remarks>
        public static void ImportQryAS400ToTableDataBaseWithFileData(string FileName, string Member, string Separator, string ConnString, string ToTableName, string PathSaveFile = null, bool NewTbOrCollec = true, int NumberColumnImprtTb = default(int))
        {
            // เตรียม ข้อมูลของไฟล์ W 
            var FileW_Data = new FileWDataAS400(FileName);
            var Col_SaveDate = new DataColumn("SaveDate", typeof(string));
            var Col_cDate = new DataColumn("cDate", typeof(string));
            var Col_cTime = new DataColumn("cTime", typeof(string));
            var Col_chgDate = new DataColumn("chgDate", typeof(string));
            var Col_chgTime = new DataColumn("chgTime", typeof(string));
            var Col_errorMsg = new DataColumn("errorMsg", typeof(string));
            var Col_fName = new DataColumn("fName ", typeof(string));
            var Col_mName = new DataColumn("mName", typeof(string));
            var Col_mRecords = new DataColumn("mRecords", typeof(string));
            Col_SaveDate.DefaultValue = DateAndTime.Now.ToString("yyyy-MM-dd HH:mm:ss.f");
            Col_cDate.DefaultValue = FileW_Data.CreateDate; // FileWDataAS400.CreateDate
            Col_cTime.DefaultValue = FileW_Data.CreateTime;
            Col_chgDate.DefaultValue = FileW_Data.ChangeDate;
            Col_chgTime.DefaultValue = FileW_Data.ChangeTime;
            Col_errorMsg.DefaultValue = FileW_Data.ErrMSG;
            Col_fName.DefaultValue = FileW_Data.FileName;
            Col_mName.DefaultValue = FileW_Data.MemberName;
            Col_mRecords.DefaultValue = FileW_Data.NumberRecord;


            DataTable dtToImport = null;
            string AS400_FileName = FileName;

            dtToImport = FncAS400.QryAS400ToDatatableV2(AS400_FileName, Member);
            // ###############################################

            // ###############################################

            if (dtToImport.Columns.Count == 1 & Separator != null)
            {
                // มี Col เดียว เพราะเพิ่ง Qry มาจาก AS400 ใหม่ๆ หรือ เป็น DT มาจากไฟล์ W Excel ที่ยังไม่ TextToCol 
                dtToImport = FncAS400.As400DataTableTextToColumn(ref dtToImport, Separator, true, true);
                dtToImport = FncDataTable.DatatableTrimCell(dtToImport);
            }
            else
                // ไม่ต้อง TxtToCol เซฟไฟล์เก็บไว้อย่างเดียว
                dtToImport = FncSave_LoadGridFile.Trim_DataTable(dtToImport);

            // เพิ่ม Column Data ก่อน Import หรือ เซฟไฟล์ Text
            int Pos = 0;
            dtToImport.Columns.Add(Col_SaveDate); dtToImport.Columns[Col_SaveDate.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_cDate); dtToImport.Columns[Col_cDate.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_cTime); dtToImport.Columns[Col_cTime.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_chgDate); dtToImport.Columns[Col_chgDate.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_chgTime); dtToImport.Columns[Col_chgTime.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_errorMsg); dtToImport.Columns[Col_errorMsg.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_fName); dtToImport.Columns[Col_fName.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_mName); dtToImport.Columns[Col_mName.ColumnName].SetOrdinal(Pos); Pos += 1;
            dtToImport.Columns.Add(Col_mRecords); dtToImport.Columns[Col_mRecords.ColumnName].SetOrdinal(Pos); Pos += 1;

            // บันทึกเป็น Text File ด้วย
            if (PathSaveFile != null)
            {
                PathSaveFile = FncFileFolder.NewFileNameUnique(PathSaveFile);
                FncSave_LoadGridFile.DataTableSaveToTxtFile1(ref dtToImport, PathSaveFile, Separator);
            }
            // ###################
            var SqlServer = new MsSql_Manager(ConnString); // (txt_Host.Text, txt_User.Text, txt_Passw.Text, txt_Database.Text)
                                                                  // ## นำ ข้อมูล As400 เข้า DataBase
                                                                  // สร้างคำสั่งสร้างตารางสำหรับ นำ ไฟล์ As400 เข้า DataBase
            string Imp_TableName = ToTableName;

            // ###################
            // NewTbOrCollec = True  เมื่อต้องการ  สร้าง Table ใหม่
            if (NewTbOrCollec == true)
            {

                // กำหนดจำนวน Column ของ Table
                int CountColCreateTb;
                if (NumberColumnImprtTb == default(int))
                    CountColCreateTb = dtToImport.Columns.Count;
                else
                    CountColCreateTb = NumberColumnImprtTb;

                string SqlCrtImptTable = GenCreateTableImport(Imp_TableName, CountColCreateTb);

                // สร้างตาราง ตรวจสอบว่ามีตารางหรือไม่แล้ว ถ้ามีให้ลบก่อนสร้าง
                if (SqlServer.TableExists(Imp_TableName))
                    SqlServer.DeleteTable(Imp_TableName);
                SqlServer.ExecuteNonQuery(SqlCrtImptTable);
            }
            // ###################
            SqlServer.CopyDatatableToDatabaseTable(dtToImport, Imp_TableName);

            // MsgBox("บันทึกไฟล์เข้าฐานข้อมูลแล้ว")
            Interaction.Beep();
            SqlServer.CloseConnection();
        }
        /// <summary> นำไฟล์ EXCEL => DataTable => TextToColumn => ลง Table ใน DataBase => บันทึกเป็นไฟล์ ? </summary>
    /// <param name="XLSFileName">ที่อยู่ไฟล์ EXCEL ที่ต้องการ</param>
    /// <param name="Separator">ตัวใช้แบ่ง คอลัมน์</param>
    /// <param name="ConnString">Connection String ของ SQL Server ปลายทาง</param>
    /// <param name="ToTableName">ชื่อตาราง ปลายทาง</param>
    /// <param name="PathSaveFile">ตำแหน่งเซฟไฟล์ ถ้าต้องการ Backup ให้ใส่ Parameter</param>
    /// <remarks>เรียกใช้ ฟังก์ชั่น สำเร็จรูปอื่นๆ</remarks>
        public static void ImportXLSFileToTableDataBase(string XLSFileName, string Separator, string ConnString, string ToTableName, string PathSaveFile = null)
        {
            DataTable dtToImport = null;

            // 590711 เพิ่ม การใช้ อักขรพิเศษ จะเพิ่ม เท่าที่เจอ หรือ เท่าที่เพิ่มได้
            if ((Strings.Left(Separator, 2) ?? "") == "[#" & (Strings.Right(Separator, 2) ?? "") == "#]")
            {
                switch (Strings.Mid(Separator, 3, Strings.Len(Separator) - 4))
                {
                    case "vbTab":
                        {
                            Separator = Constants.vbTab;
                            break;
                        }

                    default:
                        {
                            //Separator = Separator;
                            break;
                        }
                }
            }

            if ((Path.GetExtension(XLSFileName) ?? "") == ".txt")
                dtToImport = FncSave_LoadGridFile.LoadTxtToDataTable(XLSFileName, "<&>");
            else
                // dtToImport = FncExcel.ConvertExcelFileToDataTableV5(txt_File.Text, numUD.Value)
                dtToImport = FncExcel.ConvertExcelFileToDataTableV5(XLSFileName, 1, 2);

            if (dtToImport.Columns.Count == 1 & (Separator != null | !string.IsNullOrEmpty(Separator)))
                // มี Col เดียว เพราะเพิ่ง Qry มาจาก AS400 ใหม่ๆ หรือ เป็น DT มาจากไฟล์ W Excel ที่ยังไม่ TextToCol 
                dtToImport = FncAS400.As400DataTableTextToColumn(ref dtToImport, Separator, true, true);
            else
                // ไม่ต้อง TxtToCol เซฟไฟล์เก็บไว้อย่างเดียว
                dtToImport = FncSave_LoadGridFile.Trim_DataTable(dtToImport);
            if (PathSaveFile != null)
            {
                PathSaveFile = FncFileFolder.NewFileNameUnique(PathSaveFile);
                FncSave_LoadGridFile.DataTableSaveToTxtFile1(ref dtToImport, PathSaveFile, Separator);
            }
            // ###################
            var SqlServer = new MsSql_Manager(ConnString); // (txt_Host.Text, txt_User.Text, txt_Passw.Text, txt_Database.Text)
                                                                  // ## นำ ข้อมูล As400 เข้า DataBase
                                                                  // สร้างคำสั่งสร้างตารางสำหรับ นำ ไฟล์ As400 เข้า DataBase
            string Imp_TableName = ToTableName;

            string SqlCrtImptTable = GenCreateTableImport(Imp_TableName, dtToImport.Columns.Count); // + 1) '+1 เผื่อไว้แก้ตาราง

            // สร้างตาราง ตรวจสอบว่ามีตารางหรือไม่แล้ว ถ้ามีให้ลบก่อนสร้าง
            if (SqlServer.TableExists(Imp_TableName))
                SqlServer.DeleteTable(Imp_TableName);
            SqlServer.ExecuteNonQuery(SqlCrtImptTable);

            // ###################
            SqlServer.CopyDatatableToDatabaseTable(dtToImport, Imp_TableName);

            // MsgBox("บันทึกไฟล์เข้าฐานข้อมูลแล้ว")
            Interaction.Beep();
        }

        /// <summary> นำไฟล์ EXCEL => DataTable => TextToColumn => ลง Table ใน DataBase => บันทึกเป็นไฟล์ ? </summary>
    /// <param name="XLSFileName">ที่อยู่ไฟล์ EXCEL ที่ต้องการ</param>
    /// <param name="Separator">ตัวใช้แบ่ง คอลัมน์</param>
    /// <param name="ConnString">Connection String ของ SQL Server ปลายทาง</param>
    /// <param name="ToTableName">ชื่อตาราง ปลายทาง</param>
    /// <param name="PathSaveFile">ตำแหน่งเซฟไฟล์ ถ้าต้องการ Backup ให้ใส่ Parameter</param>
    /// <remarks>เรียกใช้ ฟังก์ชั่น สำเร็จรูปอื่นๆ</remarks>
        public static void ImportXLSFileToTableDataBase_WithAttribute(string XLSFileName
                                                       , string Separator
                                                       , string ConnString
                                                       , string ToTableName
                                                       , string PathSaveFile = null
                                                       , ArrayList ArL_AttrNameAndVal = null
                                                       , bool WantColumnNameTableByDataTable = false)
        {
            DataTable dtToImport = null;

            // 590711 เพิ่ม การใช้ อักขรพิเศษ จะเพิ่ม เท่าที่เจอ หรือ เท่าที่เพิ่มได้
            if ((Strings.Left(Separator, 2) ?? "") == "[#" & (Strings.Right(Separator, 2) ?? "") == "#]")
            {
                switch (Strings.Mid(Separator, 3, Strings.Len(Separator) - 4))
                {
                    case "vbTab":
                        {
                            Separator = Constants.vbTab;
                            break;
                        }

                    default:
                        {
                            //Separator = Separator;
                            break;
                        }
                }
            }

            if ((Path.GetExtension(XLSFileName) ?? "") == ".txt")
                dtToImport = FncSave_LoadGridFile.LoadTxtToDataTable(XLSFileName, "<&>");
            else
                // dtToImport = FncExcel.ConvertExcelFileToDataTableV5(txt_File.Text, numUD.Value)
                dtToImport = FncExcel.ConvertExcelFileToDataTableV5(XLSFileName, 1, 2);

            if (dtToImport.Columns.Count == 1 & (Separator != null | !string.IsNullOrEmpty(Separator)))
                // มี Col เดียว เพราะเพิ่ง Qry มาจาก AS400 ใหม่ๆ หรือ เป็น DT มาจากไฟล์ W Excel ที่ยังไม่ TextToCol 
                dtToImport = FncAS400.As400DataTableTextToColumn(ref dtToImport, Separator, true, true);
            else
                // ไม่ต้อง TxtToCol เซฟไฟล์เก็บไว้อย่างเดียว
                dtToImport = FncSave_LoadGridFile.Trim_DataTable(dtToImport);
            if (PathSaveFile != null)
            {
                PathSaveFile = FncFileFolder.NewFileNameUnique(PathSaveFile);
                FncSave_LoadGridFile.DataTableSaveToTxtFile1(ref dtToImport, PathSaveFile, Separator);
            }
            // ###################
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
            var SqlServer = new MsSql_Manager(ConnString); // (txt_Host.Text, txt_User.Text, txt_Passw.Text, txt_Database.Text)
                                                                  // ## นำ ข้อมูล As400 เข้า DataBase
                                                                  // สร้างคำสั่งสร้างตารางสำหรับ นำ ไฟล์ As400 เข้า DataBase
            string Imp_TableName = ToTableName;

            string SqlCrtImptTable; // = FncDataBaseTool.GenCreateTableImport(Imp_TableName, dtToImport.Columns.Count) '+ 1) '+1 เผื่อไว้แก้ตาราง
            if (WantColumnNameTableByDataTable == true)
                SqlCrtImptTable = GenCreateTableByDataTableImport(Imp_TableName, dtToImport);
            else
                SqlCrtImptTable = GenCreateTableImport(Imp_TableName, dtToImport.Columns.Count);

            // สร้างตาราง ตรวจสอบว่ามีตารางหรือไม่แล้ว ถ้ามีให้ลบก่อนสร้าง
            if (SqlServer.TableExists(Imp_TableName))
                SqlServer.DeleteTable(Imp_TableName);
            SqlServer.ExecuteNonQuery(SqlCrtImptTable);

            // ###################
            SqlServer.CopyDatatableToDatabaseTable(dtToImport, Imp_TableName);

            // MsgBox("บันทึกไฟล์เข้าฐานข้อมูลแล้ว")
            Interaction.Beep();
        }

        // New No Test
        private static DataTable ConvertToDataTable(ref ArrayList arraylist)
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Column1");
            dataTable.Columns.Add("Column2");


            for (int count = 1, loopTo = arraylist.Count; count <= loopTo; count++)
            {
                var drow = dataTable.NewRow();
                drow[0] = arraylist[count];
                drow[1] = arraylist[count];
                dataTable.Rows.Add(drow);
            }
            return dataTable;
        }

        /// <summary>
    /// สร้าง Temporary Table ด้วย การ Select into จาก Qry หรือ View
    /// เพื่อความรวดเร็วในการ Qry ต่อชีท เพราะ ทุกๆชีทจะต้องไปเปิด View ใหม่ทุกครั้ง
    /// </summary>
    /// <param name="tbName">Table หรือ View ที่ต้องการ นำมาสร้างเป็น Temp Table</param>
    /// <param name="AliasViewName">ชื่อแฝง กำหนดกรณี ต้องการ ใช้ชื่ออื่น</param>
    /// <returns>ส่งกลับเป็นชื่อ Temp Table  =>  Temp_{TableName}</returns>
    /// <remarks>
    /// 591102 เพิ่ม  AliasViewName
    /// </remarks>
        public static string CreateTempTableByTable(MsSql_Manager CSql, string tbName, string AliasViewName = null)
        {
            string TempTbName = tbName;
            if (AliasViewName != null | AliasViewName != "")
                TempTbName = AliasViewName;

            //string TempTable = string.Format("Temp_{0}", TempTbName);
            string TempTable = string.Format("{0}", TempTbName);
            if (CSql.TableExists(TempTable))
                CSql.DeleteTable(TempTable);

            string sqlSelectInto = "select * into " + TempTable + " from " + tbName;
            CSql.ExecuteNonQuery(sqlSelectInto);
            return TempTable;
        }
    }
}
