using System.Diagnostics;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Devices;
using System.Collections;
using System;
using System.Xml.Linq;
using System.Net;
using System.Windows.Forms;
using System.IO;
using System.Web;


namespace KengsLibraryCs
{
    public static class CFncNet
    {
        /// <summary>ดาวน์โหลด ไฟล์ และเก็บไฟล์ไว้ : ใช้คำสั่ง WebClient1.DownloadFile </summary>
    /// <param name="URL_Download">URL ที่ต้องการ ดาวน์โหลดไฟล์</param>
    /// <param name="PathSaveFile">Path ชื่อไฟล์ นามสกุลที่ต้องการเก็บไฟล์ไว้</param>
        public static void DownloadFileFromURLtoSave(string URL_Download, string PathSaveFile)
        {
            var WebClient1 = new WebClient();
            WebClient1.Encoding = System.Text.Encoding.UTF8;
            // Dim HttpLink As String
            // HttpLink = "http://192.10.10.62:9080/Computer/FileXfer?libname=QS36F&filename=wf.h17&membername=m560206&userid=pcs&passwd=pcu8"
            // StatusList()
            try
            {
                WebClient1.DownloadFile(URL_Download, PathSaveFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : Download File Fail !! : " + ex.Message);
            }
        }

        /// <summary>ดาวน์โหลด ไฟล์ และเก็บไฟล์ไว้ โดยใช้ คำสั่ง My.Computer.Network.DownloadFile </summary>
    /// <param name="URL_Download">URL ที่ต้องการ ดาวน์โหลดไฟล์</param>
    /// <param name="PathSaveFile">Path ชื่อไฟล์ นามสกุลที่ต้องการเก็บไฟล์ไว้</param>
    /// <param name="TimeOutSec">หมดเวลา</param>
    /// <param name="ShowStatus">แสดงสถานะ</param> <param name="OverWrite">เขียนทับ</param>
    /// <param name="UserName">User</param> <param name="Password">Password</param>
        public static void DownloadFileFromURLtoSaveV1(string URL_Download, string PathSaveFile, int TimeOutSec = 60, bool ShowStatus = true, bool OverWrite = true, string UserName = "", string Password = "")
        {
            TimeOutSec = TimeOutSec * 1000; // millisecond
            try
            {
                Network net  = new Network();
                net.DownloadFile(URL_Download,PathSaveFile,UserName,Password,ShowStatus,TimeOutSec,OverWrite);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : Download File Fail !! : " + ex.Message);
            }
        }

        // Private Sub GetIPAddressV00()
        // Dim strHostName As String
        // Dim strIPAddress As String
        // strHostName = System.Net.Dns.GetHostName()
        // strIPAddress = System.Net.Dns.GetHostByName(strHostName).AddressList(0).ToString()
        // MessageBox.Show("Host Name: " & strHostName & "; IP Address: " & strIPAddress)
        // End Sub

        public static string GetIPAddressV01()
        {
            System.Web.HttpContext context;
            context = System.Web.HttpContext.Current;
            string sIPAddress = context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (string.IsNullOrEmpty(sIPAddress))
                return context.Request.ServerVariables["REMOTE_ADDR"];
            else
            {
                var ipArray = sIPAddress.Split(new char[] { ',' });
                return ipArray[0];
            }
        }
        public static string[] GetAllIPAddressV02()
        {
            string host = Dns.GetHostName();
            var ip = Dns.GetHostEntry(host);

            var IpList = new ArrayList();

            foreach (var IpInList in ip.AddressList)
                IpList.Add(IpInList.ToString());

            var ret = (string[])IpList.ToArray(typeof(string));

            return ret; // IpList.ToArray()
        }

        /// <summary>
    /// ตรวจสอบ Public IP ของเรา จากเว็บ http://checkip.dyndns.org/
    /// อนาคต อาจจะใช้ไม่ได้
    /// </summary>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string GetPublicIP()
        {
            string direction = "";
            var request = WebRequest.Create("http://checkip.dyndns.org/");
            using (var response = request.GetResponse())
            {
                using (var stream = new StreamReader(response.GetResponseStream()))
                {
                    direction = stream.ReadToEnd();
                }
            }

            // Search for the ip in the html
            int first = direction.IndexOf("Address: ") + 9;
            int last = direction.LastIndexOf("</body>");
            direction = direction.Substring(first, last - first);    // หาตำแหน่ง จาก Tag

            return direction;
        }

        /// <summary>
    /// Ping ได้ค่า เป็น Avg Ping
    /// </summary>
    /// <param name="hostNameOrAddress"></param>
    /// <param name="PingTimes"></param>
    /// <returns>>0 คือ ติดต่อได้ / -1 คือ ติดต่อไม่ได้ </returns>
    /// <remarks></remarks>
        public static long GetPingMs(ref string hostNameOrAddress, int PingTimes = 3)
        {
            var ping = new System.Net.NetworkInformation.Ping();
            var PingTimeOut = default(long);
            var PingAvg = default(long);
            System.Net.NetworkInformation.PingReply pingRes;

            for (int RunPing = 1, loopTo = PingTimes; RunPing <= loopTo; RunPing++)
            {
                pingRes = ping.Send(hostNameOrAddress);
                PingAvg += pingRes.RoundtripTime;
                if ((int)pingRes.Status == (int)System.Net.NetworkInformation.IPStatus.TimedOut)
                    PingTimeOut += 1;
            }

            PingAvg = PingAvg / (long)PingTimes;

            if (PingTimeOut == PingTimes)
                PingAvg = -1;
            else if (PingTimeOut != PingTimes & PingAvg < 1)
                PingAvg = 1;
            return PingAvg; // PingTimeOut 'ping.Send(hostNameOrAddress).RoundtripTime
        }

        /// <summary>
    /// สั่งเรียกไฟล์ EXE ผ่าน Network แล้วไม่ให้ขึ้น Security Warning
    /// โดยสั่งปิด Security ก่อน
    /// </summary>
    /// <param name="PathExeToRun"></param>
    /// <remarks></remarks>
        //public static bool RunExeOnNetwork_SecurityOff(string PathExeToRun)
        //{
        //    bool RunExeOnNetw_SecurityOffRet = default(bool);
        //    var oShell = Interaction.CreateObject("Wscript.Shell");
        //    var oEnv = oShell.Environment("PROCESS");
        //    oEnv("SEE_MASK_NOZONECHECKS") = 1; // ปิด Security Warning
        //    try
        //    {
        //        Process.Start(PathExeToRun);
        //        RunExeOnNetw_SecurityOffRet = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        RunExeOnNetw_SecurityOffRet = false;
        //    }
        //    finally
        //    {
        //        oEnv.Remove("SEE_MASK_NOZONECHECKS");
        //    }// เปิด Security Warning คืน

        //    return RunExeOnNetw_SecurityOffRet;
        //}
    }
}
