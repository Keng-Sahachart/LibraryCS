/*####################################################################################################################
 คลาสที่สืบทอดจาก DataGridView ปกติ  ที่ มีคำสั่ง ป้องกัน การคลิก คอลัมน์ ที่กำหนด
 วิธีใช้ :DataGridView1.ColumnToSkip = 2  => คอลัมน์ที่ 2 ไม่สามารถ คลิกได้ 
 credit : https://social.msdn.microsoft.com/Forums/vstudio/en-US/641b99ba-9d6d-45d1-a2c6-0e5b998d1b03/lockdisable-columnrowcells-in-datagridview?forum=vbgeneral
####################################################################################################################*/
using System.Windows.Forms;

namespace KengsLibraryCs
{

    // DataGridView
    public class CtrlDataGridView_SkipColumn : DataGridView
    {

        private int mColumnToSkip = -1;
        public int ColumnToSkip
        {
            get
            {
                return mColumnToSkip;
            }
            set
            {
                mColumnToSkip = value;
            }
        }

        protected new override bool SetCurrentCellAddressCore(int columnIndex, int rowIndex, bool setAnchorCellAddress, bool validateCurrentCell, bool throughMouseClick)
        {
            if (columnIndex == mColumnToSkip && mColumnToSkip != -1)
            {
                if (mColumnToSkip == ColumnCount - 1)
                    return base.SetCurrentCellAddressCore(0, rowIndex + 1, setAnchorCellAddress, validateCurrentCell, throughMouseClick);
                else if (ColumnCount != 0)
                    return base.SetCurrentCellAddressCore(columnIndex + 1, rowIndex, setAnchorCellAddress, validateCurrentCell, throughMouseClick);
            }

            return base.SetCurrentCellAddressCore(columnIndex, rowIndex, setAnchorCellAddress, validateCurrentCell, throughMouseClick);
        }

        protected new override void SetSelectedCellCore(int columnIndex, int rowIndex, bool selected)
        {
            if (columnIndex == mColumnToSkip)
            {
                if (mColumnToSkip == ColumnCount - 1)
                    base.SetSelectedCellCore(0, rowIndex + 1, selected);
                else if (ColumnCount != 0)
                    base.SetSelectedCellCore(columnIndex + 1, rowIndex, selected);
            }
            else
                base.SetSelectedCellCore(columnIndex, rowIndex, selected);
        }

    }
}
