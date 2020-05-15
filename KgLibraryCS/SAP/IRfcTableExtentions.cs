using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SAP.Middleware.Connector;

public partial class IRfcTableExtentions
{

    /// Converts SAP table to .NET DataTable table

    public static System.Data.DataTable ToDataTable(IRfcTable sapTable, string name)
    {
        var adoTable = new System.Data.DataTable(name);
        // ... Create ADO.Net table.
        int liElement = 0;
        while (liElement < sapTable.ElementCount)
        {
            RfcElementMetadata metadata = sapTable.GetElementMetadata(liElement);
            adoTable.Columns.Add(metadata.Name, GetDataType(metadata.DataType));
            Math.Max(System.Threading.Interlocked.Increment(ref liElement), liElement - 1);
        }

        // Transfer rows from SAP Table ADO.Net table.
        foreach (IRfcStructure row in sapTable)
        {
            DataRow ldr = adoTable.NewRow();
            liElement = 0;
            while (liElement < sapTable.ElementCount)
            {
                RfcElementMetadata metadata = sapTable.GetElementMetadata(liElement);
                var switchExpr = metadata.DataType;
                switch (switchExpr)
                {
                    case RfcDataType.DATE:
                        {
                            ldr[metadata.Name] = row.GetString(metadata.Name).Substring(0, 4) + row.GetString(metadata.Name).Substring(5, 2) + row.GetString(metadata.Name).Substring(8, 2);
                            break;
                        }

                    case RfcDataType.BCD:
                        {
                            ldr[metadata.Name] = row.GetDecimal(metadata.Name);
                            break;
                        }

                    case RfcDataType.CHAR:
                        {
                            ldr[metadata.Name] = row.GetString(metadata.Name);
                            break;
                        }

                    case RfcDataType.STRING:
                        {
                            ldr[metadata.Name] = row.GetString(metadata.Name);
                            break;
                        }

                    case RfcDataType.INT2:
                        {
                            ldr[metadata.Name] = row.GetInt(metadata.Name);
                            break;
                        }

                    case RfcDataType.INT4:
                        {
                            ldr[metadata.Name] = row.GetInt(metadata.Name);
                            break;
                        }

                    case RfcDataType.FLOAT:
                        {
                            ldr[metadata.Name] = row.GetDouble(metadata.Name);
                            break;
                        }

                    default:
                        {
                            ldr[metadata.Name] = row.GetString(metadata.Name);
                            break;
                        }
                }

                Math.Max(System.Threading.Interlocked.Increment(ref liElement), liElement - 1);
            }

            adoTable.Rows.Add(ldr);
        }

        return adoTable;
    }

    private static Type GetDataType(RfcDataType rfcDataType)
    {
        switch (rfcDataType)
        {
            case RfcDataType.DATE : //rfcDataType.DATE:
                {
                    return typeof(string);
                }
            case RfcDataType.CHAR:
                {
                    return typeof(string);
                }
            case RfcDataType.STRING:
                {
                    return typeof(string);
                }
            case RfcDataType.BCD:
                {
                    return typeof(decimal);
                }
            case RfcDataType.INT2:
                {
                    return typeof(int);
                }
            case RfcDataType.INT4:
                {
                    return typeof(int);
                }
            case RfcDataType.FLOAT:
                {
                    return typeof(double);
                }
            default:
                {
                    return typeof(string);
                }
        }
    }
}