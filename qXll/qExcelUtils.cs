using kx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qXll
{
    public static class qExcelUtils
    {
        public static object ConvertToExcelType(object o)
        {
            try
            {
                if (o is null) return "";
                Type t = o.GetType();
                if (t == typeof(System.String)) return o;
                else if (t == typeof(System.Double)) return o;
                else if (t == typeof(System.Int16)) return o;
                else if (t == typeof(System.Int32)) return o;
                else if (t == typeof(System.Int64)) return o;
                else if (t == typeof(c.Date)) { c.Date v = (c.Date)o; return v.DateTime(); }
                else if (t == typeof(System.TimeSpan)) { System.TimeSpan v = (System.TimeSpan)o; return v.TotalDays; }
                else if (t == typeof(c.Minute)) return o.ToString();
                else if (t == typeof(c.Second)) return o.ToString();
                else if (t == typeof(System.Char[])) return new string((System.Char[])o);
                return o.ToString();
            }
            catch (Exception e) { return "#Type error: " + e.Message; }
        }
    }
}
