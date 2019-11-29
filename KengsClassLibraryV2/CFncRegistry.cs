using System;
using Microsoft.Win32;

namespace KengsLibraryCs
{
    public static class CFncRegistry
    {
        public static string  GetRegisterByPath(string PathVal,string ValueName )
        {
            return (string) Registry.GetValue(PathVal, ValueName, String.Empty);

        }

        public static string SetRegisterByPath(string PathVal, string ValueName, string Value)
        {
            Registry.SetValue(PathVal, ValueName, Value);
            return (string) Registry.GetValue(PathVal, ValueName, String.Empty);
        }

        //string RegisterExists(string PathVal, string ValueName)
        //{
            
        //}
    }
}