using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using ExcelDna.Logging;
using Log4Net;


namespace CSharpAddIn
{
    class TestArrayTopic : ExcelRtdServer.Topic
    {
        public string _prefix;
        public bool _israndom;

        public TestArrayTopic(ExcelRtdServer server, int topicId) :
            base(server, topicId)
        {
        }
    }

    [ComVisible(true)]
    public class TestRtdServer : ExcelRtdServer
    {
        string _logPath;

        Random _random;
        Timer _timer;
        List<TestArrayTopic> _topics;
        static ILog Logger = LogManager.GetLogger("TestRtdServer");

        public TestRtdServer()
        {
            _logPath = @"C:\temp\ExcelDnaRtd.log";

            _random = new Random();
            _topics = new List<TestArrayTopic>();
            _timer = new Timer(UpdateTopics, null, 0, 1000);
            Log("TimerServer created");
        }

        void Log(string format, params object[] args)
        {
            if (Directory.Exists(@"C:\temp\"))
                //Directory.CreateDirectory(@"C:\temp\");
                File.AppendAllText(_logPath, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " - " + string.Format(format, args) + "\r\n");
        }

        int GetTopicId(Topic topic)
        {
            return topic.TopicId;
        }

        protected override bool ServerStart()
        {
            Log("ServerStart");
            Logger.Debug(">>>>> ServerStart called.");
            return true;
        }

        protected override void ServerTerminate()
        {
            Log("ServerTerminate");
            Logger.Debug(">>>>> ServerTerminate called.");
            _timer.Dispose();
            _timer = null;
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            Log("CreateTopic: TopicId - {0}, topicInfo: {1}, {2}", topicId, topicInfo[0], topicInfo[1]);
            Logger.Debug(">>>>> CreateTopic called.");
            return new TestArrayTopic(this, topicId) { _prefix = topicInfo[0], _israndom = Convert.ToBoolean(topicInfo[1]) };
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Log("ConnectData: TopicId - {0}, topicInfo: {1}", GetTopicId(topic), string.Join(", ", topicInfo));
            Logger.Debug(">>>>> ConnectData called.");
            _topics.Add((TestArrayTopic)topic);
            return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
        }

        protected override void DisconnectData(Topic topic)
        {
            Log("DisconnectData: TopicId - {0}", GetTopicId(topic));
            Logger.Debug(">>>>> DisconnectData called.");
            _topics.Remove((TestArrayTopic)topic);
        }

        void UpdateTopics(object _unused)
        {
            foreach (TestArrayTopic topic in _topics)
            {
                Log("UpdateTopics: TopicId - {0}, Prefix - {1}, Random - {2}, Value - {3}", GetTopicId(topic), topic._prefix, topic._israndom, topic.Value);
                Logger.Debug(">>>>> UpdateTopics called.");
                var value = DateTime.Now.ToString("HH:mm:ss.fff");

                if (topic._israndom)
                    value = topic._prefix + ";" + _random.NextDouble().ToString("F5");
                else
                    value = topic._prefix + ";" + DateTime.Now.ToString("HH:mm:ss.fff");

                topic.UpdateValue(value);
            }
        }
    }
}
