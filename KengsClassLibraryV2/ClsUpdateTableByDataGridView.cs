using System.Data;
using Microsoft.VisualBasic;
using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace KengsLibraryCs
{

/// <summary>
/// 590127
/// สำหรับใช้ ผูก Table กับ DataGridView เพื่อ Update
/// วิธีใช้อยู่ใน Class
/// How to VB. . .
/// (1) Dim TestClass As New ClsUpdateTableByDataGridView
/// (2) TestClass = New ClsUpdateTableByDataGridView(DataGridView1, connString)
/// (3) DataGridView1.DataSource = TestClass.Qry("select *  from [Table]")
/// (4) Chang Data In DataGridView
/// (5) TestClass.UpdateTableAfterChangeDgv()
/// </summary>
/// <remarks></remarks>
    public class ClsUpdateTableByDataGridView
    {
        // 
        // อย่างน้อย ต้อง มี Primary Key ออกมาด้วย ถึงจะ อัพเดต ด้วย DataAdapter ได้
        public SqlConnection conn = new SqlConnection(); // (connString)
        public SqlDataAdapter da = new SqlDataAdapter();
        public SqlCommandBuilder cb;//= new SqlCommandBuilder(da); // ต้องมีตัวนี้ ถึงจะอัพเดตได้ แต่ไม่ต้องเรียกใช้งานอย่างอื่นเลย
        public DataSet ds = new DataSet();

        public DataGridView DataGridView;

        public ClsUpdateTableByDataGridView(ref DataGridView Dgv, string connStr)
        {
            cb = new SqlCommandBuilder(da);
            if (Strings.Len(connStr) > 0)
                conn = new SqlConnection(connStr);
            DataGridView = Dgv;
        }
        public ClsUpdateTableByDataGridView(ref DataGridView Dgv, SqlConnection connNew)
        {
            cb = new SqlCommandBuilder(da);
            conn = connNew;
            DataGridView = Dgv;
        }

        /// <summary>
        /// Qry Select เพื่อลง DataGridView
        /// </summary>
        /// <param name="QrySel">String Select</param>
        /// <returns>DataTable</returns>
        /// <remarks></remarks>
        public DataTable Qry(string QrySel)
        {
            try
            {
                da.SelectCommand = new SqlCommand(QrySel, conn);

                // da.FillSchema(ds, SchemaType.Source, "employee") ' ไม่ต้องมีก็ได้
                da.Fill(ds, "tb");

                var dt = ds.Tables["tb"];

                // Dim r As New Random(System.DateTime.Now.Millisecond)
                // da.Update(ds, "tb") 'เอาออก เพราะ ถ้า Qry เพื่อ Refesh ล้างข้อมูลที่แก้ หรือ ยกเลิกการแก้ข้อมูล  แล้ว มันจะ แก้ไขข้อมูล เลย

                ds.Clear();
                da.Fill(ds, "tb");

                return ds.Tables["tb"];
            }
            catch (Exception ex)
            {
                Interaction.MsgBox("Error: " + ex.ToString());
            }
            finally
            {
            }
            return null;
        }

        /// <summary>
    /// อัพเดตข้อมูล จาก DataGridView ลงฐานข้อมูล พร้อม ส่งกลับ DataTable ที่แก้ไขแล้ว
    /// </summary>
    /// <returns>DataTable ที่เปลี่ยนแปลงแล้ว</returns>
    /// <remarks></remarks>
        public DataTable UpdateTableAfterChangeDgv()
        {
            da.FillSchema(ds, SchemaType.Mapped, "tb");
            da.Update(ds, "tb");
            return ds.Tables["tb"];
        }
    }
}
