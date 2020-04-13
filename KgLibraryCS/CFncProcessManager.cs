using System.Diagnostics;
using Microsoft.VisualBasic;
using System;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace KengsLibraryCs
{
    public static class CFncProcessManager
    {
        [DllImport("user32.dll"
        , CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, ref int lpdwProcessId);

        /// <summary>
    /// ฟังก์ชั่น ที่ไว้ประกาศตัวแปร Application
    /// </summary>
    /// <param name="Application"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static object CreateObj(string Application)
        {
            //object xlAppObj;
            //xlAppObj = Interaction.CreateObject("Excel.Application");
            Excel.Application xlAppObj = new Excel.Application();
            int xlHWND = Conversions.ToInteger(xlAppObj.Hwnd);
            int ProcIdXL = 0;
            GetWindowThreadProcessId((IntPtr)xlHWND, ref ProcIdXL);
            var xproc = Process.GetProcessById(ProcIdXL);
            return null;
        }

        public static string CheckProcess(string ProcessName)
        {
            Process[] myProcesses;
            bool statusProcess;
            statusProcess = true;
            myProcesses = Process.GetProcessesByName(ProcessName);
            foreach (Process instance in myProcesses)
            {
                statusProcess = false;
                return Conversions.ToString(statusProcess);
            }
            return Conversions.ToString(statusProcess);
        }
        public static string CloseProc(string sProcName)
        {
            var Proc = Process.GetProcessesByName(sProcName);
            try
            {
                Proc[0].Kill();
                return "Process killed";
            }
            catch
            {
                return "Can't find process";
            }
        }


        /// <summary>ปิด Process ที่มีสถานะค้าง หรือไม่ได้ใช้งาน </summary>
    /// <param name="ProcessName"></param>
        public static void KillRemainProcess(string ProcessName)   // ควรใช้ตอนเสร็จกระบวนการแล้วเท่านั้น อาจจะไปปิดโปรเซสที่รอการทำงานอยู่ 
        {
            try
            {
                Process[] pProcess = null;
                pProcess = Process.GetProcessesByName(ProcessName);
                foreach (Process p in pProcess)
                {
                    // เช็คเฉพาะที่เป็น Process ค้าง
                    if (p.MainWindowTitle.Length == 0)
                        p.Kill();
                }
            }
            catch (Exception ex)
            {
            }
        }
        //public static object CleanUp(string procName)
        //{
        //    object objProcList;
        //    object objWMI;
        //    //object objProc;
        //    try
        //    {
        //        // create WMI object instance
        //        objWMI = Interaction.GetObject("winmgmts:");
        //        Debug.Print("Cleaning up " + procName);
        //        if (!(objWMI == null))
        //        {
        //            // create object collection of Win32 processes
        //            objProcList = objWMI.InstancesOf("win32_process");
        //            foreach (object objProc in objProcList) // iterate through enumerated collection
        //            {
        //                if (Operators.ConditionalCompareObjectEqual(UCase(objProc.Name), Strings.UCase(procName), false))
        //                {
        //                    objProc.Terminate(0);
        //                    Debug.Print(procName + " was terminated");
        //                }
        //            }
        //        }
        //    }
        //    catch
        //    {

        //    }

        //    objProcList = null;
        //    objWMI = null;
        //    return default(object);
        //}
    }
}
