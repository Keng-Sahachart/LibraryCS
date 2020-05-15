using Microsoft.VisualBasic;
using System;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

namespace kgLibraryCs
{
    public static class FncDateTime
    {

        /// <summary> เดือน ชื่อไทย ย่อหรือ เต็ม </summary>
        /// <param name="NumMonth">เลขเดือน 1-12</param> <param name="Abbreviate">ย่อหรือไม่</param>
        /// <returns>ชื่อเดือน หรือ ผิดพลาด ถ้า ไม่ใช่ 1-12</returns>
        public static string ThaiMonthName(int NumMonth, bool Abbreviate = true)
        {
            Interaction.IIf(NumMonth > 12 | NumMonth < 1, NumMonth == 0, null);
            var ThaiMonth = new[] { "ผิดพลาด", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฏาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤษจิกายน", "ธันวาคม" };
            var ThaiMonthAbbreviate = new[] { "ผิดพลาด", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค." };
            if (Abbreviate == true)
                return ThaiMonthAbbreviate[NumMonth];
            return ThaiMonth[NumMonth];
        }

        public static string WeekDayNameThai(int nDayOfWeek, bool IsAbbr = false)
        {
            string RetDayName;
            var ThaiDayFullName = new[] { "อาทิตย์", "จันทร์", "อังคาร", "พุุธ", "พฤหัสบดี", "ศุกร์", "เสาร์" };
            var ThaiDayAbbrName = new[] { "อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส." };

            // Dim EngDayFullName() As String = {"อาทิตย์", "จันทร์", "อังคาร", "พุุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"}
            // Dim EngiDayAbbrName() As String = {"Mon", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส."}
            if (IsAbbr == false)
                RetDayName = ThaiDayFullName[nDayOfWeek];
            else
                RetDayName = ThaiDayAbbrName[nDayOfWeek];

            return RetDayName;
        }

        /// <summary>
        /// ปีไทย พ.ศ. ออกมา โดยดูจาก จำนวนว่าน่าจะเป็นปีอะไร
        /// </summary>
        /// <param name="AD_Or_BE">ปี ค.ศ. หรือ พ.ศ.</param>
        /// <param name="Want_2Digit">ต้องการเป็น 2 Digit ใช่หรือไม่</param>
        /// <returns>ส่งออกเป็นปีไทย พ.ศ.</returns>
        /// <remarks>เป็นฟังก์ชั่น ที่ใช้วิธีการ ประมาณการณ์ ซึ่งอาจไม่ถูกต้องเสมอไป / ใช้ได้อีก 500 ปี</remarks>
        public static int ThaiYear(int AD_Or_BE, bool Want_2Digit = false)
        {
            if (AD_Or_BE < 2500)
                AD_Or_BE += 543;
            if (Want_2Digit == true)
                AD_Or_BE = Conversions.ToInteger(Strings.Right(Conversions.ToString(AD_Or_BE), 2));
            return AD_Or_BE;
        }

        /// <summary>โหมดปี สำหรับฟังก์ชั่น YearMode </summary>
        public enum YearModeChoice
        {
            Buddhist,
            Christian
        }
        /// <summary> ใส่ตัวเลขเข้าไป แล้วกำหนด โหมดปี ว่าจะเอาปี ค.ศ. หรือ พ.ศ. </summary>
        /// <param name="nYear">เลขปี พ.ศ. ควรมากกว่า 2500</param>
        /// <param name="ModeChoice">โหมดปี มีตัวเลือกให้ใส่</param>
        /// <returns>โหมดปี ที่ต้องการ</returns>
        public static int YearMode(int nYear, YearModeChoice ModeChoice)
        {
            if (nYear > 2500)
                nYear -= 543;// ทำการปรับให้เป็น ค.ศ. ก่อน
            if (ModeChoice == (int)YearModeChoice.Buddhist)
                return nYear + 543;
            return nYear;
        }

        /// <summary>หาวันแรกของเดือน จากวันปัจจุบัน</summary>
        public static DateTime GetFirstDayOfMonth(DateTime CurrentDate)
        {
            return new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
        }

        /// <summary> หาวันแรกของเดือน ที่เป็นวันทำงาน (จันทร์-ศุกร์) จากวันปัจจุบัน</summary>
        /// <param name="CurrentDate">DateTime ของเดือนที่ต้องการหา</param>
        /// <returns>ส่งออกมาเป็น DateTime วันแรกที่ทำงานของเดือน</returns>
        public static DateTime GetFirstWorkingDayOfMonth(DateTime CurrentDate)
        {
            if ((int)new DateTime(CurrentDate.Year, CurrentDate.Month, 1).DayOfWeek == (int)DayOfWeek.Saturday)
                return new DateTime(CurrentDate.Year, CurrentDate.Month, 1).AddDays(2);
            else if (new DateTime(CurrentDate.Year, CurrentDate.Month, 1).DayOfWeek == (int)DayOfWeek.Sunday)
                return new DateTime(CurrentDate.Year, CurrentDate.Month, 1).AddDays(1);
            else
                return new DateTime(CurrentDate.Year, CurrentDate.Month, 1).AddDays(0);
        }

        /// <summary>หาวันสุดท้ายของเดือน จากวันปัจจุบัน </summary>
        /// <param name="CurrentDate">DateTime ของเดือนที่ต้องการหา</param>
        /// <returns>ส่งออกมาเป็น DateTime วันสุดท้ายของเดือน</returns>
        public static DateTime GetLastDayOfMonth(DateTime CurrentDate)
        {
            return new DateTime(CurrentDate.Year, CurrentDate.Month, DateTime.DaysInMonth(CurrentDate.Year, CurrentDate.Month));
        }

        public static void DateTimePickerThaiDisplay(ref DateTimePicker DateTimePicker)
        {
            var ThaiCultureInfo = System.Globalization.CultureInfo.GetCultureInfo("th-TH");
            DateTimePicker.Format = DateTimePickerFormat.Custom; // "ddd dd MMM yyyy" '
            var formats = DateTimePicker.Value.GetDateTimeFormats(ThaiCultureInfo);
            DateTimePicker.CustomFormat = formats[8]; // "d ""ddd dd MMM yyyy" '
        }
    }
}
