using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Windows.Forms;

/*
########################################################
#### คลาสสำหรับ ซ่อน และ แสดง tabpage 
#### ซึ่ง TabPage ปกติ จะไม่มี Properties Visible ให้ ใช้
#### วิธีใช้: ประกาศ คลาส ไว้เป็น ตัวแปร Global
#### Credit: http://stackoverflow.com/questions/18058034/hiding-and-showing-tabpages-in-vb-net-tabmanager
########################################################
dim clsTabManager as new ClsTabManager 'ตัวแปร Global

-->Hiding a tabpage:
clsTabManager.SetInvisible(tabPage)

-->Showing a tabpage (call from any class/form):
clsTabManager.SetVisible(FormWithTabControl.tabPage, FormWithTabControl.TabControl)

-->Showing a tabpage (call from Form where TabControl resides):
clsTabManager.SetVisible(tabPage, TabControl)

*/

namespace KengsLibraryCs
{

    public class ClsTabManager
    {
        private struct TabPageData
        {
            internal int Index;
            internal TabControl Parent;
            internal TabPage Page;

            internal TabPageData(int index__1, TabControl parent__2, TabPage page__3)
            {
                Index = index__1;
                Parent = parent__2;
                Page = page__3;
            }

            internal static string GetKey(TabControl tabCtrl, TabPage tabPage)
            {
                string key = "";
                if (tabCtrl != null && tabPage != null)
                    key = String.Format("{0}:{1}", tabCtrl.Name, tabPage.Name);
                return key;
            }
        }

        private Dictionary<string, TabPageData> hiddenPages = new Dictionary<string, TabPageData>();

        public void SetVisible(TabPage page, TabControl parent)
        {
            if (parent != null && !parent.IsDisposed)
            {
                TabPageData tpinfo;
                string key = TabPageData.GetKey(parent, page);

                if (hiddenPages.ContainsKey(key))
                {
                    tpinfo = hiddenPages[key];

                    if (tpinfo.Index < parent.TabPages.Count)
                        parent.TabPages.Insert(tpinfo.Index, tpinfo.Page);
                    else
                        // add the page in the same position it had
                        parent.TabPages.Add(tpinfo.Page);

                    hiddenPages.Remove(key);
                }
            }
        }

        public void SetInvisible(TabPage page)
        {
            if (IsVisible(page))
            {
                TabControl tabCtrl = (TabControl)page.Parent;
                TabPageData tpinfo;
                tpinfo = new TabPageData(tabCtrl.TabPages.IndexOf(page), tabCtrl, page);
                tabCtrl.TabPages.Remove(page);
                hiddenPages.Add(TabPageData.GetKey(tabCtrl, page), tpinfo);
            }
        }

        public bool IsVisible(TabPage page)
        {
            return page != null && page.Parent != null;
        }

        public void CleanUpHiddenPage(TabPage page)
        {
            foreach (TabPageData info in hiddenPages.Values)
            {
                if (info.Parent != null && info.Parent.Equals((TabControl)page.Parent))
                    info.Page.Dispose();
            }
        }

        public void CleanUpAllHiddenPages()
        {
            foreach (TabPageData info in hiddenPages.Values)
                info.Page.Dispose();
        }
    }

}
