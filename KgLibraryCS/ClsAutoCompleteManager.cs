using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// ยังไม่เสร็จ 
// ยังใช้ไม่ได้
namespace kgLibraryCs
{
    public class ClsAutoCompleteManager
    {
        
        ArrayList Arr_autoComplete = new ArrayList();
        int NumControl=0;

        public void AddControl(TextBox TxtBx)
        {
            NumControl++;

            AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
            TxtBx.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            TxtBx.AutoCompleteSource = AutoCompleteSource.CustomSource;
            TxtBx.AutoCompleteCustomSource = autoComplete;

            Arr_autoComplete.Add(new {TxtBx.Name, TxtBx, autoComplete});
        }

    }
}
