using System.Collections;
using System.Windows.Forms;
using System.Drawing;

namespace KengsLibraryCs
{
    public static class FncDataGrid
    {

        /// <summary>นำ DataGridView มาแอด Column พร้อมกำหนดชื่อ text กว้าง และผูกกับ Column จากตารางที Query มาใส่</summary>
    /// <param name="DTGV">DataGridView ที่ต้องการ</param><param name="ColName">ชื่อคอลลัมน์</param>
    /// <param name="HeaderText">ข้อความที่จะแสดงที่หัวคอลัมน์</param><param name="Width">ขนาดความกว้าง</param>
    /// <param name="BindingColumn">ชื่อ Column ของตารางที่จะผูกด้วย</param>
        public static void DtgvAddColumn(ref DataGridView DTGV, string ColName, string HeaderText, int Width = default(int), string BindingColumn = null, DataGridViewColumnObject ObjCol = -1)
        {
            var DGVColumnObj = new object();

            if ((int)ObjCol > -1)
            {
                switch (ObjCol)
                {
                    case DataGridViewColumnObject.ButtonColumn:
                        {
                            DGVColumnObj = new DataGridViewButtonColumn();
                            break;
                        }

                    case DataGridViewColumnObject.CheckBoxColumn:
                        {
                            DGVColumnObj = new DataGridViewButtonColumn();
                            break;
                        }

                    case DataGridViewColumnObject.ComboBoxColumn:
                        {
                            DGVColumnObj = new DataGridViewComboBoxColumn();
                            break;
                        }

                    case DataGridViewColumnObject.ImageColumn:
                        {
                            DGVColumnObj = new DataGridViewImageColumn();
                            break;
                        }

                    case DataGridViewColumnObject.LinkColumn:
                        {
                            DGVColumnObj = new DataGridViewLinkCell();
                            break;
                        }

                    case DataGridViewColumnObject.TextBoxColumn:
                        {
                            // DGVColumnObj = New DataGridViewTextBoxColumn
                            goto TextCol;
                            break;
                        }
                }

                DGVColumnObj.Name = ColName;
                DGVColumnObj.HeaderText = HeaderText;
                DGVColumnObj.Text = "X";

                // DGVColumnObj.Width = 50
                DGVColumnObj.UseColumnTextForButtonValue = true; // เปิดการแสดง Text ที่ปุ่ม
                DTGV.Columns.Add(DGVColumnObj);
            }
            else
            {
            TextCol:
                ;
                DTGV.Columns.Add(ColName, HeaderText);
            }
            // ######################################

            if (Width != default(int))
                DTGV.Columns[ColName].Width = Width;

            if (BindingColumn != null)
                DTGV.Columns[ColName].DataPropertyName = BindingColumn;
        }

        enum DataGridViewColumnObject
        {
            ButtonColumn,
            CheckBoxColumn,
            ComboBoxColumn,
            ImageColumn,
            LinkColumn,
            TextBoxColumn
        }
        struct ColumnDataGridData
        {
            public string Name;
            public string HeaderText;
            public int Width;
            public object ObjCoumn;
            public string DataPropertyName;
        }
        public static bool IsHeaderButtonCell(DataGridView GridView, DataGridViewCellEventArgs e)
        {
            return GridView.Columns[e.ColumnIndex] is DataGridViewButtonColumn && !(e.RowIndex == -1);
        }
        public static bool IsHeaderObjectCell(DataGridView GridView, DataGridViewCellEventArgs e, DataGridViewColumnObject TypeObject)
        {
            var ret = default(bool);
            switch (TypeObject)
            {
                case DataGridViewColumnObject.ButtonColumn:
                    {
                        ret = GridView.Columns[e.ColumnIndex] is DataGridViewButtonColumn && !(e.RowIndex == -1);
                        break;
                    }

                case DataGridViewColumnObject.CheckBoxColumn:
                    {
                        ret = GridView.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn && !(e.RowIndex == -1);
                        break;
                    }

                case DataGridViewColumnObject.ComboBoxColumn:
                    {
                        ret = GridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && !(e.RowIndex == -1);
                        break;
                    }

                case DataGridViewColumnObject.ImageColumn:
                    {
                        ret = GridView.Columns[e.ColumnIndex] is DataGridViewImageColumn && !(e.RowIndex == -1);
                        break;
                    }

                case DataGridViewColumnObject.LinkColumn:
                    {
                        ret = GridView.Columns[e.ColumnIndex] is DataGridViewLinkColumn && !(e.RowIndex == -1);
                        break;
                    }

                case DataGridViewColumnObject.TextBoxColumn:
                    {
                        ret = GridView.Columns[e.ColumnIndex] is DataGridViewTextBoxColumn && !(e.RowIndex == -1);
                        break;
                    }
            }
            return ret;
        }

        /// <summary>
    /// ค้นหาข้อมูลใน Row
    /// </summary>
    /// <param name="Dgv"></param>
    /// <param name="ColName"></param>
    /// <param name="FindVal"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static DataGridViewRow[] fnDgv_FindValue(ref DataGridView Dgv, string ColName, string FindVal)
        {
            var arlRowRes = new ArrayList();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                if (row.Cells[ColName].Value.ToString().Contains(FindVal))
                    arlRowRes.Add(row);
            }
            var RetRow = arlRowRes.ToArray(typeof(DataGridViewRow)) as DataGridViewRow[];
            return RetRow;
        }

        public static DataGridViewButtonColumn CreateColumnButton(string Name, string HeaderText, int Width = default(int), string Text = "")
        {
            var Btn = new DataGridViewButtonColumn();
            {
                var withBlock = Btn;
                withBlock.Name = Name;
                withBlock.HeaderText = HeaderText;
                withBlock.Width = Width;
                withBlock.Text = Text;
                withBlock.UseColumnTextForButtonValue = true; // เปิดการแสดง Text ที่ปุ่ม
            }
            return Btn;
        }
        public static DataGridViewCheckBoxColumn CreateColumnCheckBox(string Name, string HeaderText, int Width = default(int), string Text = "")
        {
            var Chk = new DataGridViewCheckBoxColumn();
            {
                var withBlock = Chk;
                withBlock.Name = Name;
                withBlock.HeaderText = HeaderText;
                withBlock.Width = Width;
                withBlock.ReadOnly = false;

                withBlock.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                withBlock.FlatStyle = FlatStyle.Standard;
                withBlock.CellTemplate = new DataGridViewCheckBoxCell();
                withBlock.CellTemplate.Style.BackColor = Color.Beige;
            }

            return Chk;
        }
    }
}
