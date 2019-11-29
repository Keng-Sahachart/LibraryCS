using Microsoft.VisualBasic;
using System;
using System.Xml;
using System.Net;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualBasic.Devices;
namespace KengsLibraryCs
{



    /// <summary>

/// สำหรับ เช็คข้อมูลของไฟล์ จาก AS400

/// สามารถ เรียกใช้ได้เลย ไม่ต้องประกาศตัว แปร Class ก็ได้

/// เช่น ClsAS400FileData.GetMember("XXXXX")

/// โครงสร้างของ TAG XML เป็นไปตาม ที่คุณ Turbo กำหนดไว้

/// </summary>

/// <remarks></remarks>
    public class ClsAS400FileData
    {

        // ตัวแปร ข้อมูลที่รับได้ จากไฟล์
        public string CreateDate;
        public string CreateTime;

        public string ChangeDate;
        public string ChangeTime;

        public string FileName;
        public string MemberName;
        public int NumberRecord;

        public string ErrMSG;

        private const string hostAS400P8 = "192.10.10.10";
        // เครื่อง Check Data ของ K Turbo
        private const string HostChkMem = "172.28.1.30";
        private string UrlChkMem = "http://" + HostChkMem + ":9080/ComputerMedline/rest/Getmember?filename=";

        public ClsAS400FileData(string iFileName)
        {
            GetDataFromAS400(iFileName);
        }

        ClsAS400FileData()
        {
        }

        public bool GetDataFromAS400(string iFileName)
        {
            if (Strings.Len(iFileName) < 1)
                return false;
            iFileName = iFileName.ToUpper();

            string Url = UrlChkMem + iFileName;
            // URL = "http://" & HostCheck & ":9081/Rest/rest/getMember?filename=" & iFileName
            // UrlChkMem = "http://172.28.1.30:9080/ComputerMedline/rest/Getmember?filename=" & iFileName

            var client = new WebClient();
            string xml_S;
            Network netw = new Network(); ;
            if (netw.Ping(HostChkMem) & netw.Ping(hostAS400P8))
            {
                try
                {
                    xml_S = client.DownloadString(Url);
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            else
                return false;

            var xml = new XmlDocument();

            xml.LoadXml(xml_S);

            ErrMSG = xml.SelectSingleNode("/output")["errorMsg"].InnerText.ToString();
            string FileNameChk = xml.SelectSingleNode("/output")["fName"].InnerText;
            if (string.IsNullOrEmpty(ErrMSG) & !string.IsNullOrEmpty(FileNameChk))
            {
                FileName = Strings.Trim(xml.SelectSingleNode("/output")["fName"].InnerText);
                MemberName = Strings.Trim(xml.SelectSingleNode("/output")["mName"].InnerText);

                CreateDate = Strings.Trim(xml.SelectSingleNode("/output")["cDate"].InnerText);
                CreateTime = Strings.Trim(xml.SelectSingleNode("/output")["cTime"].InnerText);

                ChangeDate = Strings.Trim(xml.SelectSingleNode("/output")["chgDate"].InnerText);
                ChangeTime = Strings.Trim(xml.SelectSingleNode("/output")["chgTime"].InnerText);

                NumberRecord = Conversions.ToInteger(Strings.Trim(xml.SelectSingleNode("/output")["mRecords"].InnerText));
                return true;
            }

            return false;
        }

        /// <summary>
    /// ดึงค่า Member
    /// *ตัวอย่าง การใช้งาน ==>
    /// IsNothing_(ClsAS400FileData.GetMember(txt_Trnfr_FileName01.Text), MemberDate)
    /// </summary>
    /// <param name="iFileName">ชื่อไฟล์ W</param>
    /// <returns>รับค่าออกมาเป็น Member ของไฟล์ W</returns>
    /// <remarks></remarks>
        public static string GetMember(string iFileName)
        {
            var ClsMe = new ClsAS400FileData(); // = Nothing
            if (Strings.Len(iFileName) < 1)
                return null;
            iFileName = iFileName.ToUpper();
            if ((ClsMe.FileName ?? "") != (iFileName ?? ""))
                ClsMe.GetDataFromAS400(iFileName);
            if ((ClsMe.FileName ?? "") == (iFileName ?? ""))
                return ClsMe.MemberName;
            return null;
        }


        /// <summary>
    /// เช็คว่ามีไฟล์อยู่ไหม
    /// </summary>
    /// <param name="iFileName">ชื่อไฟล์ W</param>
    /// <param name="iMember">ค่า Member ใส่หรือไม่ใส่ก็ได้</param>
    /// <returns>True มี หรือ False ไม่มี</returns>
    /// <remarks></remarks>
        public bool CheckFileExits(string iFileName, string iMember = "")
        {
            iFileName = iFileName.ToUpper();
            iMember = iMember.ToUpper();
            if ((FileName ?? "") != (iFileName ?? ""))
                GetDataFromAS400(iFileName);

            switch (iMember)
            {
                case  null:
                    {
                        if ((FileName ?? "") == (iFileName ?? ""))
                            return true;
                        break;
                    }
                case "":
                    {
                        if ((FileName ?? "") == (iFileName ?? ""))
                            return true;
                        break;
                    }
                default:
                    {
                        if ((FileName ?? "") == (iFileName ?? "") & (iMember ?? "") == (MemberName ?? ""))
                            return true;
                        break;
                    }
            }
            return false;
        }

        public string GetFileInfo(string iFileName = null)
        {
            if (FileName == null)
                GetMember(iFileName);

            string textData = "ข้อมูลไฟล์ มีดังนี้ " + Constants.vbNewLine;
            textData += "FileName : " + FileName + Constants.vbNewLine;
            textData += "MemberName : " + MemberName + Constants.vbNewLine;
            textData += "CreateDate : " + CreateDate + Constants.vbNewLine;
            textData += "CreateTime : " + CreateTime + Constants.vbNewLine;
            textData += "ChangeDate : " + ChangeDate + Constants.vbNewLine;
            textData += "ChangeTime : " + ChangeTime + Constants.vbNewLine;
            textData += "NumberRecord : " + Conversions.ToString(NumberRecord) + Constants.vbNewLine;
            textData += "ErrMSG : " + ErrMSG + Constants.vbNewLine;

            return textData;
        }
    }
}
