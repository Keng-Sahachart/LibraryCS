using System;
using System.Collections;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using Microsoft.VisualBasic;


using System.Text.RegularExpressions;

namespace kgLibraryCs
{
    class FncArray
    {

        #region Convert Array to Any

        /// <summary>
        /// แปลง Array To DataTable
        /// </summary>
        /// <param name="ArrayLine"></param>
        /// <param name="Delimiter"></param>
        /// <returns></returns>
        public DataTable ArrayToDataTable(object[] ArrayLine, string Delimiter = "~")
        {
            DataTable dt = new DataTable();

            int TopDataCount = 0;
            foreach (var THisLine in ArrayLine) // My.Computer.FileSystem.ReadAllText(Filename).Split(Environment.NewLine)
            {
                var DataAry = Regex.Split(THisLine.ToString(),Delimiter);
                // StringWatch &= DataAry
                dt.Rows.Add(DataAry);
                if (DataAry.GetType().IsArray == true)
                {
                    TopDataCount = (DataAry.Count() > TopDataCount ? DataAry.Count(): TopDataCount);
                }
            }

            return dt;
        }

    }
    #endregion Convert Array to Any

}

