﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

///https://stackoverflow.com/questions/14522540/close-a-messagebox-after-several-seconds
//var userResult = AutoClosingMessageBox.Show("Yes or No?", "Caption", 1000, MessageBoxButtons.YesNo);
//if(userResult == System.Windows.Forms.DialogResult.Yes) { 
//    // do something
//}
//...

/* ใช้แทน MsgBoxAutoClose */
namespace kgLibraryCs
{
    public class AutoClosingMessageBox
    {
        System.Threading.Timer _timeoutTimer;
        string _caption;
        DialogResult _result;
        DialogResult _timerResult;
        AutoClosingMessageBox(string text, string caption, int timeout_ms, MessageBoxButtons buttons = MessageBoxButtons.OK, DialogResult timerResult = DialogResult.None)
        {
            _caption = caption;
            _timeoutTimer = new System.Threading.Timer(OnTimerElapsed,
                null, timeout_ms, System.Threading.Timeout.Infinite);
            _timerResult = timerResult;
            using (_timeoutTimer)
                _result = MessageBox.Show(text, caption, buttons);
        }
        /// <summary>
        /// Show Msg ภายในเวลาที่กำหนด 
        /// </summary>
        /// <param name="text">ข้อความที่จะแสดง</param>
        /// <param name="caption">หัว Title</param>
        /// <param name="timeout_ms">หมดเวลา ms</param>
        /// <param name="buttons">ปุ่มที่จะแสดง</param>
        /// <param name="timerResult">?</param>
        /// <returns></returns>
        public static DialogResult Show(string text, string caption, int timeout_ms, MessageBoxButtons buttons = MessageBoxButtons.OK, DialogResult timerResult = DialogResult.None)
        {
            return new AutoClosingMessageBox(text, caption, timeout_ms, buttons, timerResult)._result;
        }

        void OnTimerElapsed(object state)
        {
            IntPtr mbWnd = FindWindow("#32770", _caption); // lpClassName is #32770 for MessageBox
            if (mbWnd != IntPtr.Zero)
                SendMessage(mbWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            _timeoutTimer.Dispose();
            _result = _timerResult;
        }
        const int WM_CLOSE = 0x0010;
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
    }
}