using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public static class CommonExtension
    {
        public static bool ToContains(this DataColumnCollection Columns, string columnName)
        {
            return Columns.Contains(columnName);
        }

        public static string ToStr(this object str)
        {
            return Convert.ToString(str);
        }

        public static string ToStrNull(this object str)
        {
            if (string.IsNullOrWhiteSpace(Convert.ToString(str)))
                return null;

            return Convert.ToString(str);
        }

        public static decimal ToDec(this object str, decimal defaultDec = 0)
        {
            if (str == System.DBNull.Value)
                return defaultDec;
            return Convert.ToDecimal(str);
        }
        public static Int64 ToInt(this object str, Int64 defaultInt = 0)
        {
            if (str == System.DBNull.Value)
                return defaultInt;
            return Convert.ToInt64(str);
        }

        public static double ToDoub(this object str, double defaultDec = 0)
        {
            if (str == System.DBNull.Value || str == null || str.Equals(string.Empty))
                return defaultDec;

            return Convert.ToDouble(str);
        }

        public static double? ToDoubNull(this object str)
        {
            if (str == System.DBNull.Value || str == null || str.Equals(string.Empty))
                return null;

            return Convert.ToDouble(str);
        }
    }
}
