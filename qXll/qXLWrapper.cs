using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using kx;
using ExcelDna.Integration;

namespace qXll
{
    public static class qXLWrapper
    {
        [ExcelFunction(Description = "Execute a q command, synchronously or not.")]
        public static object[,] qExecute(
            [ExcelArgument(Description = "q command")] string query,
            [ExcelArgument(Description = "(optional) Synchronous. Default is false.")] bool synchronous,
            [ExcelArgument(Description = "(optional) Server address or ip. Default is \"localhost\".")] string host = "",
            [ExcelArgument(Description = "(optional) Server port. Default is 5001")] int port = 0)
        {
            object[,] o = null; object[,] err = new object[1, 1];
            if (query == "") return o;
            if (host == "") host = "localhost";
            if (port == 0) port = 5001;
            c c = null;
            try { c = new c(host, port); } catch (Exception e) { err[0, 0] = "#Connection Error: " + e.Message; return err; }
            o = new object[1, 1]; DateTime st = DateTime.Now;
            try
            {
                if (synchronous)
                {
                    c.k(query);
                    DateTime et = DateTime.Now;
                    o[0, 0] = "Success. Command executed in " + (et - st).TotalMilliseconds.ToString() + "ms";
                }
                else
                {
                    c.ks(query);
                    o[0, 0] = "Success. Command sent at " + st.ToString();
                }
                return o;
            }
            catch (Exception e) { err[0, 0] = "#Error Command: " + e.Message; return err; }
        }

        [ExcelFunction(Description = "Execute a q query and return a 2 dimensional array. Note: command must return a well formatted table or keyed table.")]
        public static object[,] qQuery(
            [ExcelArgument(Description = "q query")] string query,
            [ExcelArgument(Description = "(optional) Remove headers. Default is false.")] bool noHeaders = false,
            [ExcelArgument(Description = "(optional) Server address or ip. Default is \"localhost\".")] string host = "",
            [ExcelArgument(Description = "(optional) Server port. Default is 5001")] int port = 0)
        {
            object[,] o = null; object[,] err = new object[1, 1];
            if (query == "") return o;
            if (host == "") host = "localhost";
            if (port == 0) port = 5001;
            c c = null;
            try { c = new c(host, port); } catch (Exception e){ err[0, 0] = "#Connection Error: " + e.Message; return err; }
            try
            {
                Object result = c.k(query); // synchronous query
                c.Flip flip = c.td(result); // if the result set is a keyed table, this removes the key. 
                int nRows = c.n(flip.y[0]); if (nRows <= 0) { err[0, 0] = "#Empty results: 0 rows returned."; return err; }
                int nCols = c.n(flip.x); if (nCols <= 0) { err[0, 0] = "#Empty results: 0 columns returned."; return err; }
                int startRow = 0;
                if (noHeaders) { o = new object[nRows, nCols]; }
                else
                {
                    o = new object[nRows + 1, nCols];
                    for (int j = 0; j < nCols; j++)
                        o[0, j] = flip.x[j];
                    startRow = 1;
                }
                for (int i = 0; i < nRows; i++)
                {
                    for (int j = 0; j < nCols; j++)
                        o[i+ startRow, j] = ToCSharpType(c.at(flip.y[j], i)); // c.at extracts the cell from column,row.
                }
                c.Close();
                return o;
            }
            catch (Exception e) { err[0, 0] = "#Error Query: " + e.Message; return err; }
        }

        [ExcelFunction(Description = "Insert a 2 dimensional variant in a q table. Note: command must return a well formatted table or keyed table.")]
        public static object[,] qInsert(
            [ExcelArgument(Description = "Data. 2 dimensional variant")] object data,
            [ExcelArgument(Description = "Table name")] string tableName,
            [ExcelArgument(Description = "(optional) Create table. Default is false. If true, the first row of the variant must contain table headers. Note: If the table already  exists, it will be overridden.")] bool createTable = false,
            [ExcelArgument(Description = "(optional) Number of keyed columns. Default is 0.")] int keyedColumns = 0,
            [ExcelArgument(Description = "(optional) Synchronous. Default is false.")] bool synchronous = false,
            [ExcelArgument(Description = "(optional) Server address or ip. Default is \"localhost\".")] string host = "",
            [ExcelArgument(Description = "(optional) Server port. Default is 5001")] int port = 0)
        {
            object[,] o = null; object[,] err = new object[1, 1];
            object[,] d = null;
            try { d = (object[,])data; } catch (Exception e) { err[0, 0] = "#Error. Wrong Data format: " + e.Message; return err; }
            if (tableName == "") { err[0, 0] = "#Error: Table Name is empty."; return err; }
            if (host == "") host = "localhost";
            if (port == 0) port = 5001;
            c c = null;
            try { c = new c(host, port); } catch (Exception e) { err[0, 0] = "#Connection Error: " + e.Message; return err; }
            long nRows = d.GetLength(0);
            long nCols = d.GetLength(1);
            //Create table
            int startRow = 0;
            if (createTable)
            {
                string qry = d[0, 0].ToString() + ":()";
                if (nCols > 1)
                    for (int j = 1; j < nCols; j++) { qry = qry + " ;" + d[0, j].ToString() + ":()"; }
                qry = "([] " + qry + ")";
                if (keyedColumns > 0) qry = keyedColumns.ToString() + "!" + qry;
                qry = tableName + ": " + qry;
                try { if (synchronous) { c.k(qry); } else { c.ks(qry); } }
                catch (Exception e) { err[0, 0] = "#Error: Could not create table. q error: " + e.Message; return err; }
                startRow = 1;
                nRows -= 1;
            }
            //Bulk insert
            o = new object[1, 1];
            DateTime st = DateTime.Now;
            object[] x = new object[nCols];
            for (int j = 0; j < nCols; j++)
            {
                object[] col = new object[nRows];
                for (int i = 0; i < nRows; i++)
                {
                    col[i] = d[i + startRow, j];
                }
                x[j] = col;
            }
            try
            {
                if (synchronous)
                {
                    c.k("insert", tableName, x);
                    DateTime et = DateTime.Now;
                    o[0, 0] = "Success. Command executed in " + (et - st).TotalMilliseconds.ToString() + "ms";
                }
                else
                {
                    c.ks("insert", tableName, x);
                    o[0, 0] = "Success. Command sent at " + st.ToString();
                }
            }
            catch (Exception e) { err[0, 0] = "#Error: Could not insert data. q error: " + e.Message; return err; }
            return o;
        }
        
        //Utils
        private static object ToCSharpType(object o)
        {
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
    }
}
