using Microsoft.VisualBasic;
using System;
using Microsoft.VisualBasic.CompilerServices;


namespace KengsLibraryCs
{
    public class CFncNumber
    {

        /// <summary>
    /// ตัวเลขที่ เครื่องหมายลบ อยู่ท้าย ตัวเลข ให้อยู่หน้า เช่น 123- เป็น -123 ส่วนมากสำหรับ AS400
    /// </summary>
    /// <param name="AnyText">string ตัวเลข</param>
    /// <returns>string</returns>
    /// <remarks></remarks>
        public static string FixTrailingMinusNegativeNumbers(string AnyText)
        {
            string RetVal = AnyText;
            if ((Strings.Right(AnyText, 1) ?? "") == "-")
            {
                if (Information.IsNumeric(AnyText))
                    RetVal = Conversions.ToDouble(AnyText).ToString();
            }
            return RetVal;
        }

        /// <summary>
    /// สุ่ม ตัวเลข ตำที่กำหนด
    /// </summary>
    /// <param name="Min">ตัวเลขต่อสุด</param>
    /// <param name="Max">มากสุด</param>
    /// <returns>ตัวเลขที่ซุ่มได้</returns>
    /// <remarks></remarks>
        public static int GetRandom(int Min, int Max)
        {
            var Generator = new Random();
            return Generator.Next(Min, Max);
        }
    }
}
