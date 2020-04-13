using System.Text.RegularExpressions;

namespace KengsLibraryCs
{
    public static class CFncREGEX
    {
        /// <summary>
    /// เอา DataValue ของ TagName ออกมาจาก ValueTag
    /// </summary>
    /// <param name="ValueTag">ข้อมูล String ทั้งหมดที่บรรจุ TagName เอาไว้</param>
    /// <param name="TagName">ชื่อของข้อมูลที่จะเอา</param>
    /// <returns>ข้อมูลใน TagName</returns>
    /// <remarks>ValueTag แยก Tag ด้วย [TagName1=Val1][TagName2=Val2]</remarks>
        public static string GetTag(string ValueTag, string TagName)
        {
            if (ValueTag == null)
                return null;// ถ้าไม่มี
                            // Return nodeEntry.Value
            string Pattern = string.Format(@"\[{0}=(?<DATA>.*?)\]", TagName); // "\[DepType=(?<DATA>.*?)\]"
            var Matches = Regex.Matches(ValueTag, Pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (Matches != null && Matches.Count > 0)
            {
                // For Each Match As Match In Matches 'ตัวอย่างว่าพบหลายตัว
                if (Matches[0].Groups["DATA"].Success)
                    // Return Match.Groups("DATA").Value
                    return Matches[0].Groups["DATA"].Value;
            }
            return null; // ถ้าไม่มี
        }

        /// <summary>
    /// เพิ่มหรือเปลี่ยนแปลง Val ข้อมูลของ TagName  ใน ValueTag
    /// </summary>
    /// <param name="ValueTag">ข้อมูล String ทั้งหมดที่บรรจุ TagName เอาไว้</param>
    /// <param name="TagName">ชื่อของข้อมูลที่จะเอา</param>
    /// <param name="Value">ข้อมูลที่จะเป็น Value ของ TagName</param>
    /// <returns>แท็กใหม่</returns>
    /// <remarks>ValueTag แยก Tag ด้วย [TagName1=Val1][TagName2=Val2]</remarks>
        public static string SetTag(ref string ValueTag, string TagName, string Value)
        {
            // nodeEntry.Value = Value
            string NewTagValue = string.Format("[{0}={1}]", TagName, Value);
            if (GetTag(ValueTag, TagName) == null)
                // ถ้าไม่มีให้เพิ่ม Tag เข้าไป
                ValueTag += NewTagValue;
            else
            {
                // แก้ไข ข้อมูลที่มีอยู่แล้ว
                string PaternTag = string.Format(@"\[{0}=(.*?)\]", TagName);
                string AllText = ValueTag;
                ValueTag = Regex.Replace(AllText, PaternTag, NewTagValue);
            }
            return ValueTag;
        }

        public static bool MatchPattern(string Str, string Pattern)
        {
            var Matches = Regex.Match(Str, Pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (Matches.Success)
                return true;
            return false;
        }
    }
}
