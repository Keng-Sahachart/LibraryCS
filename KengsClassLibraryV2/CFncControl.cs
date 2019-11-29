using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;

namespace KengsLibraryCs
{
    public static class CFncControl
    {

        /// <summary>
    /// หา Control ที่เป็นตามชนิดที่เรากำหนด เพื่อนำไปใช้งาน
    /// </summary>
    /// <param name="BigControlContainer">Control ที่เป็น Container เช่น Form,GroupBox,Panel</param>
    /// <param name="typeOfControl">ระบุชนิดของ Control ที่จะหา</param>
    /// <returns>Control ตามชนิด เป็น Array</returns>
    /// <remarks></remarks>
        public static Control[] GetControls(ref Control BigControlContainer, string typeOfControl = null)
        {
            var allControls = new List<Control>();
            // this loop will get all the controls on the form
            // no matter what the level of container nesting
            // thanks to jmcilhinney at vbforums
            var ctl = BigControlContainer.GetNextControl(BigControlContainer, true);
            while (ctl != null)
            {
                allControls.Add(ctl);
                ctl = BigControlContainer.GetNextControl(ctl, true);
            }

            Control[] Ctls;
            if (typeOfControl == null)
                Ctls = allControls.ToArray();
            else
                Ctls = allControls.Where(c => c.GetType().ToString().ToLower().Contains(typeOfControl.ToLower())).ToArray();

            // Display the results. 
            return Ctls;
        }
    }
}
