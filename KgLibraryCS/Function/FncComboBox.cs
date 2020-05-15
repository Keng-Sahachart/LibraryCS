using System.Data;
using System.Windows.Forms;

namespace kgLibraryCs
{
    public static class FncComboBox
    {
        /// <summary>ค้นหา Index ของ Item / ใน กรณีที่ ComboBox ทำ DataSource ด้วย DataTable
        /// ค้นหาโดยการ นำ Item ที่เป็น DataRow ออกมาวนลูปค้นหา ข้อมูลใน Column </summary>
        /// <param name="ComboBoxCtl">ComboBox ที่จะค้นหา</param>
        /// <param name="ValueToFind">ข้อความที่จะค้นหา</param>
        /// <param name="InColumnName">กำหนด ชื่อ Column ที่จะค้นหาข้อมูล</param>
        public static int GetIndexOfValueInDtRow(ref ComboBox ComboBoxCtl, string ValueToFind, string InColumnName)
        {
            for (int Index = 0, loopTo = ComboBoxCtl.Items.Count - 1; Index <= loopTo; Index++)
            {
                DataRow dtRow = (DataRow)ComboBoxCtl.Items[Index];
                string DataVal = dtRow[InColumnName].ToString();
                if ((DataVal ?? "") == (ValueToFind ?? ""))
                    return Index;
            }
            return -1;
        }
    }
}
