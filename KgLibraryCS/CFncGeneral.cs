using Microsoft.VisualBasic;
using System.Linq;
using System;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{
    public class CFncGeneral
    {


        // ########################################################################################
        // #######    Function เกี่ยวกับ Arrray  ####################################################
        // ########################################################################################
        /// <summary> ลบ Item ของ  Array และปรับ Index ใหม่</summary>
    /// <param name="a">ตัวแปร Array</param>  <param name="index">Index ที่จะลบออก</param>
        public static void ArrayRemoveItem<T>(ref T[] a, int index)
        {
            // Move elements after "index" down 1 position.
            Array.Copy(a, index + 1, a, index, Information.UBound(a) - index);
            var oldA = a;
            a = new T[Information.UBound(a) - 1 + 1];
            // Shorten by 1 element.
            if (oldA != null)
                Array.Copy(oldA, a, Math.Min(Information.UBound(a) - 1 + 1, oldA.Length));
        }

        /// <summary> เพิ่ม Item ของ  Array และปรับ Index ใหม่</summary>
    /// <param name="a">ตัวแปร Array</param>  <param name="NewValue">ข้อมูลใหม่ที่จะเพิ่มเข้าไป</param>
        public static void ArrayAddItem<T>(ref T[] a, object NewValue)
        {
            var oldA = a;
            a = new T[Information.UBound(a) + 1 + 1];
            // Shorten by 1 element.
            if (oldA != null)
                Array.Copy(oldA, a, Math.Min(Information.UBound(a) + 1 + 1, oldA.Length));
            a[Information.UBound(a)] = (T)NewValue;
        }

        // ########################################################################################


        // #######################################################################################
        // #########    ทั่วไป     ##########################################################
        // #######################################################################################


        public static object EncodeString(ref string Str)
        {
            var latinEnc = System.Text.Encoding.GetEncoding("iso-8859-1");
            var thaiEnc = System.Text.Encoding.GetEncoding("TIS-620");
            var bytes = latinEnc.GetBytes(Str);
            string textResult = thaiEnc.GetString(bytes);
            return textResult;
        }
        public static object Encode(string str)
        {
            var utf8Encoding = new System.Text.UTF8Encoding();
            byte[] encodedString;

            encodedString = utf8Encoding.GetBytes(str);

            return encodedString.ToString();
        }

        //public static bool CheckThaiChar(string StrNotThai)
        //{
        //    for (int i = 0, loopTo = Strings.Len(StrNotThai) - 1; i <= loopTo; i++)
        //    {
        //        switch (Strings.Asc(StrNotThai[i]))
        //        {
        //            case object _ when 161 <= Strings.Asc(StrNotThai[i]) && Strings.Asc(StrNotThai[i]) <= 251 // โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
        //           :
        //                {
        //                    return true;
        //                }
        //        }
        //    }
        //    return false;
        //}


        /// <summary> นับจำนวนตัวอักษร ภายในข้อความ   </summary>
        public static int CountCharacter(string value, char ch)
        {
            return value.Count((c) => c == ch);
        }

        /// <summary>
    /// ตรวจสอบความถูกต้องของเลขบัตรประชาชน 13 หลัก
    /// </summary>
    /// <param name="IdCardNumber">เลขบัตรประชาชน 13 หลัก</param>
    /// <returns>True / False</returns>
        public static bool CheckIdCard(string IdCardNumber)
        {
            if (Strings.Len(IdCardNumber) != 13)
                return false;
            // ขั้นตอนที่ 1 - เอาเลข 12 หลักมา เขียนแยกหลักกันก่อน (หลักที่ 13 ไม่ต้องเอามา)
            // ขั้นตอนที่ 2 - เอาเลข 12 หลักนั้นมา คูณเข้ากับเลขประจำหลักของมัน
            string IdCardNumberRevrs = Strings.StrReverse(IdCardNumber);
            // Dim Digit12 = IdCardNumber.PadLeft(12)
            var MultiplyInAddr = new int[12];
            for (int nAddrIdCard = 1; nAddrIdCard <= 12; nAddrIdCard++)
                MultiplyInAddr[nAddrIdCard - 1] = Conversion.Val(IdCardNumberRevrs[nAddrIdCard]) * (nAddrIdCard + 1);
            // ขั้นตอนที่ 3 - เอาผลคูณทั้ง 12 ตัวมา บวกกันทั้งหมด จะได้ 13+24+0+10+45+32+7+24+30+8+6+6=205
            var Sum = default(int);
            foreach (var Num in MultiplyInAddr)
                Sum += Num;
            // ขั้นตอนที่ 4 - เอาเลขที่ได้จากขั้นตอนที่ 3 มา mod 11 (หารเอาเศษ) จะได้ 205 mod 11 = 7
            int ByMod11 = Sum % 11;
            // ขั้นตอนที่ 5 - เอา 11 ตั้ง ลบออกด้วย เลขที่ได้จากขั้นตอนที่ 4 จะได้ 11-7 = 4 (เราจะได้ 4 เป็นเลขในหลัก Check Digit)
            // ถ้าเกิด ลบแล้วได้ออกมาเป็นเลข 2 หลัก ให้เอาเลขในหลักหน่วยมาเป็น Check Digit (เช่น 11 ให้เอา 1 มา, 10 ให้เอา 0 มา)
            string CheckDigit = Strings.Right(Conversions.ToString(11 - ByMod11), 1); // (11 - ByMod11).ToString.PadRight(1)
            if ((CheckDigit ?? "") == (Strings.Right(IdCardNumber, 1) ?? ""))
                return true;
            else
                return false;
        }




        // #######################################################################################
        // #########    คำสั่งเสริม สำหรับเขียนโปรแกรม    ##########################################################
        // #######################################################################################
        /// <summary>
    /// เช็ค Nothing และ ส่งค่าใหม่ไป คำสั่งเลียนแบบใน MS SQL
    /// </summary>
    /// <param name="Var">ตัวแปรที่ต้องการตรวจสอบ</param>
    /// <param name="ReturnIfNothing">ค่าใหม่</param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static object IsNothing_(object Var, object ReturnIfNothing)
        {
            if (Var == null)
                return ReturnIfNothing;
            return Var;
        }



        /// <summary>
    /// คำสั่งให้ Button พรางตัวบน Picture ใช้ในกรณี ให้ Button วางอยู่บน Picture
    /// </summary>
    /// <param name="ContainerPic">Control ที่เก็บ PicBox ไว้ เช่น Form หรือ GroupBox</param>
    /// <param name="btn">Button ที่อยู่บน PictureBox</param>
    /// <param name="pbx">PictureBox</param>
    /// <remarks>580625
    /// http://www.vbforums.com/showthread.php?457826-transparent-button-over-a-picturebox
    /// </remarks>
        public static void EmbedButtonInPictureBox(Control ContainerPic, Button btn, PictureBox pbx)
        {
            var buttonLocation = pbx.PointToClient(ContainerPic.PointToScreen(btn.Location));

            btn.Parent = pbx;
            btn.Location = buttonLocation;

            var buttonBackground = new Bitmap(btn.Width, btn.Height);

            if (pbx.Image != null)
            {
                using (var g = Graphics.FromImage(buttonBackground))
                {
                    g.DrawImage(pbx.Image, new Rectangle(0, 0, buttonBackground.Width, buttonBackground.Height), btn.Bounds, GraphicsUnit.Pixel);
                }

                btn.BackgroundImage = buttonBackground;
            }
        }



        /// <summary>
    /// เช็คว่า ไฟล์ที่เลือกเป็น รูปภาพหรือไม่ ด้วยการ นำไฟล์มาเปิดวาดใหม่
    /// </summary>
    /// <param name="filename">Path ตำแหน่งของรูปภาพ</param>
    /// <returns>True=ภาพ, Flase=ไม่ใช่ภาพ</returns>
    /// <remarks>อ้างอิง http://stackoverflow.com/questions/8846654/read-image-and-determine-if-its-corrupt-c-sharp </remarks>
        public static bool IsImage(string filename)
        {
            try
            {
                // Dim imageData As Byte()
                // Using MS = New MemoryStream(imageData)
                // Using bmp = New Bitmap(MS)
                // End Using
                // End Using

                using (var bmp = new Bitmap(filename))
                {
                }
                return true;
            }
            catch //(Exception ex)
            {
                return false;
            }
        }
    }
}
