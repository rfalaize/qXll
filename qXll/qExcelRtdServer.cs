using System;
using System.Collections.Generic;
using System.Threading;
using ExcelDna.Integration.Rtd;
using kx;

namespace qXll
{
    public class qExcelRtdServer : ExcelRtdServer
    {
        private string defaultHost = "localhost";
        private int defaultPort = 5001;

        //excel interop
        public qExcelRtdServer() { }
        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            string subscriptionKey = topicInfo[0];
            string host = topicInfo[1];
            string port = topicInfo[2];
            try
            {
                object o = this.Subscribe(topic, subscriptionKey, host, port);
                System.Diagnostics.Debug.WriteLine("Info - Subscription added : " + subscriptionKey);
                return o;
            }
            catch (Exception e) { return "Error: " + e.Message; }
        }
        protected override void DisconnectData(Topic topic)
        {
            Dictionary<string, qProcessManager> oldProcessManagers = qProcessManagers;
            Dictionary<string, qProcessManager> newProcessManagers = new Dictionary<string, qProcessManager>();
            foreach (KeyValuePair< string, qProcessManager> processManager in oldProcessManagers)
            {
                try
                {
                    processManager.Value.RemoveSubscription(topic);
                    if (processManager.Value.IsAlive() && !newProcessManagers.ContainsKey(processManager.Key))
                        newProcessManagers.Add(processManager.Key, processManager.Value);
                }
                catch (Exception e) { System.Diagnostics.Debug.WriteLine("Error - catched at Disconnect data: " + e.Message); }
            }
            qProcessManagers = newProcessManagers;
        }
        protected override void ServerTerminate()
        {
            System.Diagnostics.Debug.WriteLine("Info - Rtd server closed");
        }

        //q handlers
        private Dictionary<string, qProcessManager> qProcessManagers = new Dictionary<string, qProcessManager>();   //one qSubscription for each q process

        private object Subscribe(Topic topic,string key, string host, string port)
        {
            if (key == "") throw new Exception("key is empty");
            string h = host; if (h == "") h = this.defaultHost;
            int p = 0; try { p = Convert.ToInt16(port); } catch (Exception) { throw new Exception("port must be a number"); }; if (p == 0) p = this.defaultPort;

            //create subscription
            qProcessManager processManager;
            string processAdress = h + ":" + p.ToString();
            if (qProcessManagers.ContainsKey(processAdress)) processManager = qProcessManagers[processAdress];
            else
            {
                processManager = new qProcessManager(h, p);
                qProcessManagers.Add(processAdress, processManager);
            }

            //connect to the q process
            bool isConnected = false;
            if (processManager.connection == null) isConnected = false;
            else isConnected = processManager.connection.Connected;
            if (!isConnected)
            {
                try { processManager.connection = new c(h, p); }
                catch (Exception e) { processManager.connection = null; throw new Exception(e.Message); }
            }

            //request subscription
            try { return processManager.AddSubscription(topic,key); }
            catch (Exception e) { throw new Exception(e.Message); }
            
        }

        /// <summary>
        /// Helper class
        ///     .one qProcessManager is associated to each q process
        ///     .a qProcessManager has a workthread that handles the data stream
        ///     .subscriptions to this process are stored in a dictionary. 
        ///         key = identifier of the data point to subscribe to (in xlsub table, sym column)
        ///         value = excel caller cell (Topic)
        /// </summary>
        public class qProcessManager
        {
            #region Attributes

            private string processAdress;    //host:port
            private string host;
            private int port;
            private Thread workthread;
            private bool workthreadIsAlive;
            private Dictionary<string, HashSet<Topic>> subscriptions;
            public c connection;

            #endregion Attributes

            public qProcessManager(string host, int port)
            {
                this.host = host;
                this.port = port;
                this.processAdress = this.host + ":" + this.port.ToString();
                this.subscriptions = new Dictionary<string, HashSet<Topic>>();
            }

            public string AddSubscription(Topic topic, string key)
            {
                try
                {
                    //send subscription to q process
                    this.connection.ks(".u.sub[`xlsub;`" + key + "]");
                    
                    HashSet<Topic> topics;
                    if (!this.subscriptions.ContainsKey(key))
                    {
                        topics = new HashSet<Topic>() { topic };
                        this.subscriptions.Add(key, topics);
                    }
                    else
                    {
                        topics = this.subscriptions[key];
                        if (!topics.Contains(topic)) topics.Add(topic);
                    }
                    if (this.workthread == null)
                    {
                        workthread = new Thread(new ThreadStart(WorkThreadFunction));   //start working thread
                        this.workthreadIsAlive = true;
                        workthread.Start();
                    }
                }
                catch (Exception e) { throw new Exception(e.Message); }
                return "#Requesting...";
            }
            public void RemoveSubscription(Topic topic)
            {
                Dictionary<string, HashSet<Topic>> oldsubscriptions = this.subscriptions;
                Dictionary<string, HashSet<Topic>> newsubscriptions = new Dictionary<string, HashSet<Topic>>();
                foreach (KeyValuePair<string, HashSet<Topic>> subscription in oldsubscriptions)
                {
                    HashSet<Topic> oldtopics = subscription.Value;
                    HashSet<Topic> newtopics = new HashSet<Topic>();
                    foreach (Topic oldTopic in oldtopics)
                    {
                        if (oldTopic != topic && !newtopics.Contains(oldTopic))
                            newtopics.Add(oldTopic);
                    }
                    if (newtopics.Count > 0)
                        newsubscriptions.Add(subscription.Key, newtopics);
                }
                this.subscriptions = newsubscriptions;
                if (this.subscriptions.Count == 0)
                {
                    //close q connection
                    this.workthreadIsAlive = false; //update flag to close workthread
                    try
                    {
                        this.connection.Close();
                        System.Diagnostics.Debug.WriteLine("Info - Connection to " + this.processAdress + " closed successfully.");
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine("Error - while closing connection to " + this.processAdress + ": " + e.Message);
                    }
                }
            }
            public bool IsAlive()
            {
                return this.subscriptions.Count > 0;
            }

            private void WorkThreadFunction()
            {
                while (true)
                {
                    if (!this.workthreadIsAlive)
                    {
                        System.Diagnostics.Debug.WriteLine("Info - Workthread is not alive: exit.");
                        break;
                    }
                    try
                    {
                        object result = connection.k();
                        ProcessResult(result);
                    }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
                }
            }
            private void ProcessResult(object result)
            {
                object[] o = (object[])result;
                c.Flip flip = c.td(o[2]);
                int nRows = c.n(flip.y[0]);
                int nColumns = c.n(flip.x);
                for (int i = 0; i < nRows; i++)
                {
                    try
                    {
                        string key = c.at(flip.y[0], i).ToString();
                        object value = c.at(flip.y[1], i);
                        UpdateValue(key, value);
                    }
                    catch (Exception e) { System.Diagnostics.Debug.WriteLine("Error - ProcessResult: " + e.Message); }
                }
            }
            private void UpdateValue(string key, object value)
            {
                //propagate the value to subscribed cells
                if (subscriptions.ContainsKey(key))
                {
                    HashSet<Topic> topics = subscriptions[key];
                    foreach (Topic topic in topics)
                    {
                        try
                        {
                            object o = qExcelUtils.ConvertToExcelType(value);
                            topic.UpdateValue(o);
                        }
                        catch (Exception e) { System.Diagnostics.Debug.WriteLine("Error - UpdateValue: " + e.Message); }
                    }
                }
            }

        }

    }
}
