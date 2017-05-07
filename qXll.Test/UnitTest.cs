using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using qXll;

namespace qXll.Test
{
    // To do: open a q process on localhost:5001, before running the test set
    [TestClass]
    public class UnitTest
    {
        //Success
        [TestMethod]
        public void Test_qExecute()
        {
            object[,] o = InsertSampleTable();
            Assert.IsTrue(o[0, 0].ToString().StartsWith("Success"));
        }
        [TestMethod]
        public void Test_qQuery()
        {
            InsertSampleTable();
            //With headers
            object[,] o;
            o = qXLWrapper.qQuery("select from tUnitTest where i<100");
            Assert.AreEqual(o.GetLength(0), 101);
            Assert.AreEqual(o.GetLength(1), 4);
            //Without headers
            o = qXLWrapper.qQuery("select from tUnitTest where i<100", true);
            Assert.AreEqual(o.GetLength(0), 100);
            Assert.AreEqual(o.GetLength(1), 4);
        }
        [TestMethod]
        public void Test_qInsert()
        {
            object[,] o;
            //Create table and insert
            o = qXLWrapper.qInsert(GetSampleTable(), "tUnitTest", true, 0,true);
            Assert.IsTrue(o[0, 0].ToString().StartsWith("Success"));
            //Insert new lines
            o = qXLWrapper.qInsert(GetSampleTableData(), "tUnitTest", false, 0, true);
            Assert.IsTrue(o[0, 0].ToString().StartsWith("Success"));
            //Upsert
            object[,] i = new object[2, 2];
            i[0, 0] = "sym"; i[0, 1] = "price";
            i[1, 0] = "AAPL"; i[1, 1] = 88.88; 
            //Create keyed table and insert

        }
        
        //Error
        [TestMethod]
        public void Test_Error_Connection()
        {
            object[,] o = qXLWrapper.qQuery("t", false, "123");
            Assert.IsTrue(o[0, 0].ToString().StartsWith("#Connection Error"));
        }
        [TestMethod]
        public void Test_Error_Query()
        {
            object[,] o = qXLWrapper.qQuery("wrong command");
            Assert.IsTrue(o[0, 0].ToString().StartsWith("#Error Query"));
        }

        #region Utils

        // Create a table, t, and insert 50,000 rows of random data
        private object[,] InsertSampleTable()
        {
            return qXLWrapper.qExecute("n:1000;tUnitTest:([]sym:n?`1;time:.z.p+til n;price:n?100.;size:n?1000)", true);
        }
        private object[,] GetSampleTable()
        {
            object[,] o = new object[3, 2];
            o[0, 0] = "sym"; o[0, 1] = "price";
            o[1, 0] = "AAPL"; o[1, 1] = 140.54;
            o[2, 0] = "FB"; o[2, 1] = 54.22;
            return o;
        }
        private object[,] GetSampleTableData()
        {
            object[,] o = new object[2, 2];
            o[0, 0] = "BABA"; o[0, 1] = 32.11;
            o[1, 0] = "XOM"; o[1, 1] = 11.45;
            return o;
        }

        #endregion

    }
}
