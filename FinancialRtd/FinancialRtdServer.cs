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
    class RealStockTopic : ExcelRtdServer.Topic
    {
        //该次请求的股票代码
        public string StockCode { get; set; }
        //该次请求的股票信息
        public string StockInfo { get; set; }

        public RealStockTopic(ExcelRtdServer server, int topicId) :
            base(server, topicId)
        {
        }
    }

    [ComVisible(true)]
    public class FinancialRtdServer : ExcelRtdServer
    {
        string _logPath;

        Timer _timer;
        List<RealStockTopic> _topics;
        GoogleFinancial _google;
        static ILog Logger = LogManager.GetLogger("FinancialRtdServer");

        public FinancialRtdServer()
        {
            _logPath = @"C:\temp\ExcelDnaRtd.log";

            _topics = new List<RealStockTopic>();
            _google = new GoogleFinancial();
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
            return new RealStockTopic(this, topicId) { StockCode = topicInfo[0], StockInfo = topicInfo[1] };
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Log("ConnectData: TopicId - {0}, topicInfo: {1}", GetTopicId(topic), string.Join(", ", topicInfo));
            Logger.Debug(">>>>> ConnectData called.");
            _topics.Add((RealStockTopic)topic);
            return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
        }

        protected override void DisconnectData(Topic topic)
        {
            Log("DisconnectData: TopicId - {0}", GetTopicId(topic));
            Logger.Debug(">>>>> DisconnectData called.");
            _topics.Remove((RealStockTopic)topic);
        }

        void UpdateTopics(object _unused)
        {
            foreach (RealStockTopic topic in _topics)
            {
                Log("UpdateTopics: TopicId - {0}, StockCode - {1}, StockInfo - {2}, StockValue - {3}", GetTopicId(topic), topic.StockCode, topic.StockInfo, topic.Value);
                Logger.Debug(">>>>> UpdateTopics called.");

                _google.GetRealStock(topic);
            }
        }
    }
}
