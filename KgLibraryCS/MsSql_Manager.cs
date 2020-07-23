using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Windows.Forms;

namespace kgLibraryCs
{

    public class MsSql_Manager
    {

        public string ConnectionString { set; get; }
        public SqlConnection Connection = new SqlConnection();//{ set; get; }
        public SqlTransaction Transaction { set; get; }

        //################################################################
        #region ***Class Constructor

        public MsSql_Manager(SqlConnection SqlConn) { Connection = SqlConn; }
        public MsSql_Manager() { ; }

        public MsSql_Manager(string myConnectString) { SetConnectionString(myConnectString); }

        public MsSql_Manager(String Host, String UserName, String Password, String DataBaseName)
        {
            ConnectionString = "Data Source={0};Initial Catalog={1};Persist Security Info=True;User ID={2};Password={3};Network Library=DBMSSOCN;MultipleActiveResultSets=true";
            ConnectionString = String.Format(ConnectionString, Host, DataBaseName, UserName, Password);
            SetConnectionString(ConnectionString);
        }
        #endregion
        //################################################################
        #region ***Connection Management

        public string SetConnectionString(string myConnectString)
        {
            if (Connection.ConnectionString != myConnectString)
            {
                Connection.Close();
            }
            Connection.ConnectionString = myConnectString;
            OpenConnecAuto();
            return Connection.ConnectionString;
        }

        public System.Data.ConnectionState OpenConnecAuto()
        {
            if (Connection.State == ConnectionState.Closed & Connection.ConnectionString == "")
            {
                return Connection.State;
            }
            if (Connection.State == ConnectionState.Closed & Connection.ConnectionString != "")
            {
                Connection.Open();
                return Connection.State;
            }
            return ConnectionState.Broken;
        }

        public ConnectionState CloseConnection()
        {
            Connection.Close();
            return Connection.State;
        }

        public String GetConnectionString()
        {
            return Connection.ConnectionString;
        }
        #endregion
        //################################################################
        #region ***Excute
        //******************************
        //Retuen Object Table...
        //******************************
        public DataTable QueryDataTable(string strSQL)
        {
            SqlDataAdapter myDataAdapter = new SqlDataAdapter(strSQL, Connection);
            DataTable myDataTable = new DataTable();
            //OpenConnecAuto()
            myDataTable.Clear();
            myDataAdapter.SelectCommand.CommandTimeout = 120;
            myDataAdapter.Fill(myDataTable);
            return myDataTable;
        }

        public SqlDataReader QueryDataReader(string strSQL)
        {
            SqlDataReader myDataReader;
            //OpenConnecAuto();
            SqlCommand myCommand = new SqlCommand(strSQL, Connection);
            myDataReader = myCommand.ExecuteReader();
            return myDataReader;
            //myDataReader.Close();
        }

        public DataSet QueryDataSet(String strSQL)
        {
            SqlDataAdapter myDataAdapter = new SqlDataAdapter(strSQL, Connection);
            DataSet myDataSet = new DataSet();
            //OpenConnecAuto();
            myDataAdapter.Fill(myDataSet, "table1");
            return myDataSet;
        }

        /// <summary>
        /// Qry ออกมาเป็น Array เพื่อนำไปใช้ประโยชน์ต่างๆ เช่น เอาเข้า EXCEL ด้วยวิธี .Resize
        /// การรับ ฝั่ง Excel
        /// xlsSheet1.Range("A3").Resize(SqlToArray(cSQL).GetUpperBound(0)+1, SqlToArray(cSQL).GetUpperBound(1)+1).Value = SqlToArray(cSQL)
        /// </summary>
        /// <param name="cSQL">คำสั่ง SQL Qry</param>
        /// <returns>return var Array</returns>
        public Array QueryToArray(String cSQL)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 60 * 10;
            cmd.Connection = Connection;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = cSQL;

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();

            da.Fill(ds, "XLS"); // ไม่รู้จะตั้งชื่อ อะไร ใช้ XLS ไป่ก่อน เพราะใช้กับ Excel

            // แปลง DataTable -> Ary แบบที่ 1 ทดสอบที่ 130000 Row 30 Column ความเร็วเทียบเท่า คิวรี่ผ่าน SQL Manager
            // Return ใช้เป็น as Object เพราะว่า เวลา ตอนถ่ายค่า มันจะ Convert เป็น Int string ให้เอง เช่น เอาไปลง EXCEL ตัวเลขจะเป็น ตัวเลขให้ 
            // แต่ถ้าเป็น as string เวลาลง Excel มันจะขึ้น Warning สีเขียว เพราะมันคือ เลขที่เป็น string
            // อาจจะมีปัญหากับ พวกตัวเลขที่เป็นรหัส 0 นำหน้า

            // Array DataArray = new Array[ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count];
            object[,] DataArray = new object[ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count];

            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int RunRow = 0; RunRow < ds.Tables[0].Rows.Count - 1; RunRow++)
                {
                    for (int RunCol = 0; RunCol < ds.Tables[0].Columns.Count - 1; RunCol++)
                    {
                        //DataArray.SetValue(ds.Tables[0].Rows[RunRow].ItemArray[RunCol], RunRow, RunCol);
                        DataArray[RunRow, RunCol] = ds.Tables[0].Rows[RunRow][RunCol]; // check***
                    }
                }
            }
            return DataArray;
        }

        //******************************
        // Command
        //******************************

        /// <summary>
        /// SQL Command ,Not Return Value
        /// </summary>
        /// <param name="strSQL">sql command</param>
        /// <returns>Number Row Affected</returns>
        public int ExecuteNonQuery(String strSQL)
        {
            int resNumRowAffected = 0;
            try
            {
                SqlCommand myCommand = new SqlCommand(strSQL, Connection);
                myCommand.CommandTimeout = 600;
                //OpenConnecAuto();
                resNumRowAffected = myCommand.ExecuteNonQuery();
                return resNumRowAffected;
            }
            catch (Exception ex)
            {
                Console.Beep(5000, 2000);
                MessageBox.Show("Command ExecuteNonQuery Error: \n" + ex.Message);
                return resNumRowAffected;
            }
        }
        /// <summary>
        /// Get 1 Data
        /// </summary>
        /// <param name="strSQL">sql command</param>
        /// <returns>ตัวแปร Object จะได้ออกตาม Type DataBase</returns>
        public Object ExecuteScalar(String strSQL)
        {
            Object ObjResult;
            try
            {
                SqlCommand myCommand = new SqlCommand(strSQL, Connection);
                myCommand.CommandTimeout = 3600;
                //OpenConnecAuto()
                ObjResult = myCommand.ExecuteScalar();
                return ObjResult;
            }
            catch (Exception ex)
            {
                return ex;
            }
        }

        #endregion
        //################################################################
        #region Transaction Management

        public void TransactionStart()
        {
            Transaction = Connection.BeginTransaction(IsolationLevel.ReadCommitted);
        }

        public int TransactionExecuteNonQuery(String strSQL)
        {
            int NumRowAffected;
            SqlCommand myCommand = new SqlCommand(strSQL, Connection, Transaction);
            NumRowAffected = myCommand.ExecuteNonQuery();
            return NumRowAffected;
        }

        public void TransactionRollback()
        {
            Transaction.Rollback();
        }

        public void TransactionCommit()
        {
            Transaction.Commit();
        }

        #endregion Transaction Management
        //################################################################

        //################################################################

        #region Special Function

        public Boolean DeleteTable(string TableName)
        {
            try
            {
                string sqlDel = "";
                sqlDel += " DROP TABLE [dbo].[" + TableName + "]";
                this.ExecuteNonQuery(sqlDel);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public Boolean TableExists(String TableName)
        {
            return (int)this.ExecuteScalar("SELECT COUNT(*) FROM sys.objects WHERE name = '" + TableName + "'") > 0 ? true : false;
        }

        /// <summary>
        /// เอาทั้ง DataTable ยัดเข้าไปใน Table ใน Data Base โดยอย่างน้อย Column ต้องเท่ากัน ไม่สนใจชื่อคอลัมน์
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="TagetTable">String</param>
        public void CopyDatatableToDatabaseTable(DataTable dt, String TagetTable)
        {
            SqlBulkCopy bulkcopy = new SqlBulkCopy(Connection);
            bulkcopy.DestinationTableName = TagetTable;
            bulkcopy.WriteToServer(dt);
        }

        /// <summary>
        /// เอา DataTable Copy ไปสร้าง Table ใหม่หรือทับ ตารางเดิม ตาราง สำหรับ Import 
        /// </summary>
        /// <param name="dtImp"></param>
        /// <param name="TagetTableName"></param>
        public void RepalceTableImported(DataTable dtImp, String TagetTableName)
        {
            string sqlCreateTable;
            this.ExecuteNonQuery("drop table " + TagetTableName);
            sqlCreateTable = FncDataBaseTool.GenCreateTableImport(TagetTableName, dtImp.Columns.Count);
            this.ExecuteNonQuery(sqlCreateTable);
            this.CopyDatatableToDatabaseTable(dtImp, TagetTableName);
        }


        #endregion Special Function

    }
}


