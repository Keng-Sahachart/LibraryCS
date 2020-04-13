using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace KengsLibraryCs
{
    public static class CFncRadioButton
    {

        /// <summary>
    /// ค้นหา คอนโทรล RadioButton ใน Control ที่เป็น Container
    /// </summary>
    /// <param name="CtrlContainer">คอนโทรล ที่บรรจุ Radio</param>
    /// <returns>RadioButton ที่ค้นพบ หรือ nothing</returns>
    /// <remarks></remarks>
        public static RadioButton GetRadChecked(ref Control CtrlContainer)
        {
            var RadChecked = new RadioButton();
            RadChecked = CtrlContainer.Controls.OfType<RadioButton>().Where(r => r.Checked == true).FirstOrDefault();
            if (RadChecked == null)
                return null;
            return RadChecked;
        }
    }
}
