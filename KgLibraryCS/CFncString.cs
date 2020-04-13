using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KengsLibraryCs
{
    public class CFncString
    {
        public static string Left(string Str, int length)
        {
            string result = Str.Substring(0, length);
            return result;
        }
        public static string Right(string Str, int length)
        {
            string result = Str.Substring(Str.Length - length, length);
            return result;
        }
        public static string Mid(string Str, int startIndex, int length)
        {
            string result = Str.Substring(startIndex, length);
            return result;
        }
    }
}
