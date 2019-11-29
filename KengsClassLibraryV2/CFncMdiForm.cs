using System.Windows.Forms;

namespace KengsLibraryCs
{
    public static class CFncMdiForm
    {
        // ###############################################################
        /// <summary>
    /// หาฟอร์มลูก ที่ได้เปิดเรียกใช้มาแล้ว ว่ามีชื่อนี้หรือไม่
    /// </summary>
    /// <param name="frmName">ชื่อฟอร์มลูกที่จะค้นหา</param>
    /// <param name="MotherForm">ฟอร์มแม่</param>
    /// <returns>ฟอร์มลูกที่พบ</returns>
    /// <remarks></remarks>
        public static Form FindChildForm(string frmName, Form MotherForm)
        {
            Form frmFound = null;
            if (MotherForm.HasChildren)
            {
                foreach (Form children in MotherForm.MdiChildren)
                {
                    if ((children.Name ?? "") == (frmName ?? ""))
                        frmFound = children;
                }
            }
            else
                frmFound = null;
            return frmFound;
        }

        /// <summary>
    /// เช็กว่าเปิดฟอร์มลูกขึ้นมาหรือยัง
    /// </summary>
    /// <param name="frmName">ชื่อฟอร์มลูก</param>
    /// <param name="MainForm">ฟอร์มแม่</param>
    /// <returns>มี หรือ ไม่มี</returns>
    /// <remarks></remarks>
        public static bool CheckIfOpen(string frmName, ref Form MainForm)
        {
            //Form frm;
            foreach (Form frm in MainForm.MdiChildren)
            {
                if ((frm.Name ?? "") == (frmName ?? ""))
                {
                    frm.Focus();
                    return true;
                }
            }
            return false;
        }
        /// <summary>
    /// เรียกฟอร์มลูก โดยไม่ให้เรียกเปิดซ้ำๆ
    /// </summary>
    /// <param name="FrmChild">ฟอร์มลูก</param>
    /// <remarks></remarks>
        public static Form ChildFormShowOne(ref Form FrmChild, ref Form MotherForm)
        {
            // Dim FrmMdi As New FrmChild
            Form Frm;
            string FrmName = FrmChild.Name;
            if (CheckIfOpen(FrmName, ref MotherForm) == false)
            {
                Frm = FrmChild; // New Department
                Frm.MdiParent = MotherForm;
            }
            else
                Frm = FindChildForm(FrmName, MotherForm);
            Frm.Show();
            Frm.WindowState = FormWindowState.Maximized;
            return Frm;
        }

        /// <summary>
    /// รวบรวม รายชื่อของ ChildForm มาใส่ใน ToolStripMenuItem เพื่อแสดงรายชื่อของ Child ที่เปิดขึ้นมา เพื่อเรียกใช้อีกที
    /// ตัวอย่างเช่น Excel เปิดขึ้นมาหลายไฟล์ จะมี Menu แสดง ว่าเปิดไฟล์ ไหนขึ้นมาบ้างแล้ว
    /// </summary>
    /// <param name="MdiForm">MdiForm</param>
    /// <param name="MenuDropDownItems">Menu ที่จะแสดงฟอร์มลูก</param>
    /// <remarks></remarks>
        public static void RefreshChildFormListToMenuItem(ref Form MdiForm, ref ToolStripMenuItem MenuDropDownItems)
        {
            MdiForm.Refresh();
            MenuDropDownItems.DropDownItems.Clear();
            //Form children;
            if (MdiForm.HasChildren)
            {
                foreach (Form children in MdiForm.MdiChildren)
                {
                    var x = children.WindowState;
                    MenuDropDownItems.DropDownItems.Add(children.Text);
                }
            }
        }
    }
}
