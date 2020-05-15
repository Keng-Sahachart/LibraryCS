using System;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualBasic.CompilerServices;

namespace kgLibraryCs
{
/// http://www.codeproject.com/Articles/42437/Toggling-the-States-of-all-CheckBoxes-Inside-a-Dat


/// <summary>
/// Class ที่สร้างเพื่อมา Edit คอลัมน์ ที่เป็น CheckBox ให้มี CheckBox Select All ที่หัวคอลัมน์
/// </summary>
/// <remarks>
/// วิธีใช้คือ
/// 1.ประกาศตัวแปร Class เป็น Global
/// ==> Dim HeadChk1 As New ClassChkAllDgv
/// 2.ที่ Event Form Load สั่งให้ ตัวแปรทำการ Add Header Select All ที่ชื่อคอลัมน์ ที่ต้องการ
/// ==> HeadChk1.AddHeaderCheckBox(dgvSelectAll, "chkBxSelect")
/// 3. ใช้งานตามปกติ
/// </remarks>
    public class ClsChkAllDgv
    {
        private DataGridView dgvSelectAll;
        private string ChkColumnNAme;

        private int TotalCheckBoxes = 0;
        private int TotalCheckedCheckBoxes = 0;
        private CheckBox HeaderCheckBox = null;
        private bool IsHeaderCheckBoxClicked = false;

        ClsChkAllDgv()
        {
        }

        public void AddHeaderCheckBox(ref DataGridView DataGridView, string ColumnName)
        {
            try
            {
                dgvSelectAll = DataGridView;
                ChkColumnNAme = ColumnName;

                HeaderCheckBox = new CheckBox();

                HeaderCheckBox.Size = new Size(15, 15);

                // Add the CheckBox into the DataGridView
                dgvSelectAll.Controls.Add(HeaderCheckBox);


                dgvSelectAll.DataBindingComplete += DataBindingComplete;
                HeaderCheckBox.KeyUp += new KeyEventHandler(HeaderCheckBox_KeyUp);
                HeaderCheckBox.MouseClick += new MouseEventHandler(HeaderCheckBox_MouseClick);
                dgvSelectAll.CellValueChanged += new DataGridViewCellEventHandler(dgvSelectAll_CellValueChanged);
                dgvSelectAll.CurrentCellDirtyStateChanged += new EventHandler(dgvSelectAll_CurrentCellDirtyStateChanged);
                dgvSelectAll.CellPainting += new DataGridViewCellPaintingEventHandler(dgvSelectAll_CellPainting);
            }
            catch //(Exception ex)
            {
            }
        }




        // ##########################################################################################################
        // ###    
        // ##########################################################################################################
        private void ResetHeaderCheckBoxLocation(int ColumnIndex, int RowIndex)
        {
            // Get the column header cell bounds
            var oRectangle = dgvSelectAll.GetCellDisplayRectangle(ColumnIndex, RowIndex, true);

            var oPoint = new Point();

            oPoint.X = Conversions.ToInteger(oRectangle.Location.X + (oRectangle.Width - HeaderCheckBox.Width) / (double)2 + 1);
            oPoint.Y = Conversions.ToInteger(oRectangle.Location.Y + (oRectangle.Height - HeaderCheckBox.Height) / (double)2 + 1);

            // Change the location of the CheckBox to make it stay on the header
            HeaderCheckBox.Location = oPoint;
        }

        private void HeaderCheckBoxClick(CheckBox HCheckBox)
        {
            IsHeaderCheckBoxClicked = true;

            foreach (DataGridViewRow Row in dgvSelectAll.Rows)
                ((DataGridViewCheckBoxCell)Row.Cells[ChkColumnNAme]).Value = HCheckBox.Checked;

            dgvSelectAll.RefreshEdit();

            TotalCheckedCheckBoxes = HCheckBox.Checked ? TotalCheckBoxes : 0;

            IsHeaderCheckBoxClicked = false;
        }

        private void RowCheckBoxClick(DataGridViewCheckBoxCell RCheckBox)
        {
            if (RCheckBox != null)
            {
                // Modifiy Counter;            
                if (Conversions.ToBoolean(RCheckBox.Value) && TotalCheckedCheckBoxes < TotalCheckBoxes)
                    TotalCheckedCheckBoxes += 1;
                else if (TotalCheckedCheckBoxes > 0)
                    TotalCheckedCheckBoxes -= 1;

                // Change state of the header CheckBox.
                if (TotalCheckedCheckBoxes < TotalCheckBoxes)
                    HeaderCheckBox.Checked = false;
                else if (TotalCheckedCheckBoxes == TotalCheckBoxes)
                    HeaderCheckBox.Checked = true;
            }
        }

        // ##########################################################################################################
        // ###    
        // ##########################################################################################################


        private void DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            HeaderCheckBox.Checked = false; // ล้างการเลือกทุกครั้งที่ รับค่ามาใหม่

            TotalCheckBoxes = dgvSelectAll.RowCount;
            TotalCheckedCheckBoxes = 0;

            if (dgvSelectAll.RowCount == 0)
                HeaderCheckBox.Enabled = false;
            else
                HeaderCheckBox.Enabled = true;
        }
      
        private void dgvSelectAll_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!IsHeaderCheckBoxClicked)
                RowCheckBoxClick((DataGridViewCheckBoxCell)dgvSelectAll[e.ColumnIndex, e.RowIndex]);
        }

        private void dgvSelectAll_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvSelectAll.CurrentCell is DataGridViewCheckBoxCell)
                dgvSelectAll.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void HeaderCheckBox_MouseClick(object sender, MouseEventArgs e)
        {
            HeaderCheckBoxClick((CheckBox)sender);
        }

        private void HeaderCheckBox_KeyUp(object sender, KeyEventArgs e)
        {
            if ((int)e.KeyCode == (int)Keys.Space)
                HeaderCheckBoxClick((CheckBox)sender);
        }

        private void dgvSelectAll_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex == 0)
                ResetHeaderCheckBoxLocation(e.ColumnIndex, e.RowIndex);
        }
    }
}
