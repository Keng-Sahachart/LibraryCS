using System.Data;
using System.Collections.Generic;
using System.Collections;
using System;
using System.ComponentModel;

namespace kgLibraryCs
{
    public static class FncEntityFramework
    {
        /// <summary>
    /// แปลง .ToList จาก Entity Framework ให้เป็น DataTable / อาจจะช้าเพราะ วนลูป Row, Column
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="ListData">ตัวแปร List เช่น จาก .ToList</param>
    /// <returns>DataTable</returns>
    /// <remarks></remarks>
        public static DataTable ConvertListToDataTable<T>(IList<T> ListData)
        {
            var Properties = TypeDescriptor.GetProperties(typeof(T));
            var Table = new DataTable();
            foreach (PropertyDescriptor prop in Properties)  // วนสร้าง Column ก่อน
            {
                if (prop.IsBrowsable == true & prop.IsReadOnly == false)
                    Table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            foreach (T item in ListData) // วน ใส่ Data ตาม Row
            {
                var row = Table.NewRow();
                foreach (PropertyDescriptor prop in Properties) // วน ใส่ Data ตาม Column
                {
                    if (prop.IsBrowsable == true & prop.IsReadOnly == false)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                Table.Rows.Add(row);
            }
            return Table;
        }

        /// <summary>
        /// แปลง .ToList จาก Entity Framework ให้เป็น DataTable / อาจจะช้าเพราะ วนลูป Row, Column
        /// </summary>
        /// <param name="parIList">IEnumerable ที่ยังไม่ .ToList</param>
        /// <returns>DataTable</returns>
        /// <remarks>ยังมี Column ของ Entity ติดมาด้วย</remarks>
        public static DataTable ConvertListToDataTable_(IEnumerable parIList)
        {
            var ret = new DataTable();
            try
            {
                System.Reflection.PropertyInfo[] ppi = null;
                if (parIList == null)
                    return ret;
                foreach (var itm in parIList)
                {
                    if (ppi == null)
                    {
                        ppi = itm.GetType().GetProperties();
                        foreach (System.Reflection.PropertyInfo pi in ppi)
                        {
                            var colType = pi.PropertyType;
                            if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(System.Nullable<>)))
                                colType = colType.GetGenericArguments()[0];
                            ret.Columns.Add(new DataColumn(pi.Name, colType));
                        }
                    }
                    var dr = ret.NewRow();
                    foreach (System.Reflection.PropertyInfo pi in ppi)
                        dr[pi.Name] = pi.GetValue(itm, null) == null ? DBNull.Value : pi.GetValue(itm, null);
                    ret.Rows.Add(dr);
                }
                foreach (DataColumn c in ret.Columns)
                    c.ColumnName = c.ColumnName.Replace("_", " ");
            }
            catch //(Exception ex)
            {
                ret = new DataTable();
            }
            return ret;
        }
    }
}
