using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using kx;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using System.Net.Sockets;

namespace qXll
{
    public static class qExcelFunctions
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
                if (noHeaders)
                {
                    long ubound;
                    if (nRows == 1) ubound = 2;
                    else ubound = nRows;
                    o = new object[ubound, nCols];
                }
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
                        try { o[i + startRow, j] = qExcelUtils.ConvertToExcelType(c.at(flip.y[j], i)); } // c.at extracts the cell from column,row.
                        catch (Exception e) { o[i + startRow, j] = e.Message;  }
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
                //Get default column type
                Type defaultColumnType = typeof(ExcelDna.Integration.ExcelEmpty);
                for (int i = 0; i < nRows; i++)
                {
                    object tmp = d[i + startRow, j];
                    if (tmp!= null)
                    {
                        defaultColumnType = tmp.GetType();
                        if (defaultColumnType != typeof(ExcelDna.Integration.ExcelEmpty))
                            break;
                    }
                }
                //Cast
                for (int i = 0; i < nRows; i++)
                {
                    col[i] = qExcelUtils.ConvertToqType(d[i + startRow, j], defaultColumnType);
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


        [ExcelFunction(Description = "Subscribe to a data point updates from tsub table")]
        public static object qSubscribe(
            [ExcelArgument(Description = "Subscription key (data point indentifier)")] string key,
            [ExcelArgument(Description = "(optional) Server address or ip. Default is \"localhost\".")] string host = "",
            [ExcelArgument(Description = "(optional) Server port. Default is 5001")] int port = 0)
        {
            string[] topics = new string[3];
            topics[0] = key; topics[1] = host; topics[2] = port.ToString();
            return XlCall.RTD("qXll.qExcelRtdServer", null, topics);
        }


        //Process management
        [ExcelFunction(Description = "Start a q process on localhost. Returns the process ID if success, else an error message.")]
        public static object qProcessStart(
        [ExcelArgument(Description = "(optional) visible window. Default is false (hidden).")] bool visible = false,
        [ExcelArgument(Description = "(optional) command line parameters")] string commandlineparams = "",
        [ExcelArgument(Description = "(optional) qhome: folder containaing q.q and q.k. Default is QHOME environment variable")] string qhome = "",
        [ExcelArgument(Description = "(optional) qlic: folder containaing the q license file. Default is QLIC environment variable")] string qlic = ""
        )
        {
            Process process = new Process();
            if (qhome!="") Environment.SetEnvironmentVariable("QHOME", qhome);
            if (qlic != "") Environment.SetEnvironmentVariable("QLIC", qlic);
            process.StartInfo.FileName = qhome + "q.exe";
            process.StartInfo.Arguments = commandlineparams;
            //Check if port already open
            if (commandlineparams.Contains("-p"))
            {
                string[] arguments = commandlineparams.Split(' ');
                for (int i = 0; i < arguments.Length; i++)
                {
                    if (arguments[i] == "-p")
                    {
                        if (i == arguments.Length - 1) { return "#Error: missing port number"; }
                        try
                        {
                            int port = Convert.ToInt32(arguments[i + 1]);
                            if (IsPortOpen("localhost", port, 100)) return "#Error: port " + port.ToString() + " is currently busy.";
                        }
                        catch (Exception e) { return "#Error: port number. " + e.Message; }
                    }
                }

            }
            if (visible)
                process.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            else
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            try
            {
                process.Start();
                return process.Id;
            }
            catch (Exception e)
            {
                return "#Error: " + e.Message;
            }
        }


        [ExcelFunction(Description = "Kill a running q process")]
        public static object qProcessKill(
        [ExcelArgument(Description = "process id")] int processid)
        {
            try
            {
                Process process = Process.GetProcessById(processid);
                if (process.ProcessName == "q")
                {
                    process.Kill();
                    return "Success";
                }
                else { return "#Error: process " + processid.ToString() + " is not a q process"; }
            }
            catch (Exception e)
            {
                return "#Error: " + e.Message;
            }
        }

        //Utils
        private static bool IsPortOpen(string host, int port, int timeoutms)
        {
            try
            {
                TcpClient client = new TcpClient();
                IAsyncResult result = client.BeginConnect(host, port, null, null);
                bool success = result.AsyncWaitHandle.WaitOne(timeoutms);
                if (!success) { return false; }
                client.EndConnect(result);
            }
            catch
            {
                return false;
            }
            return true;
        }

    }
}
