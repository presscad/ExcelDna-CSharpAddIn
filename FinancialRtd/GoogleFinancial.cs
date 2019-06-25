using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace CSharpAddIn
{
    class GoogleStockData
    {
        public string StockSymbol { get; set; }
        public string CompanyName { get; set; }
        public string Exchange { get; set; }
        public double Divisor { get; set; }
        public string Currency { get; set; }
        public double Last { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Volume { get; set; }
        public double AVGVolume { get; set; }
        public double MarketCaption { get; set; }
        public double Open { get; set; }
        public double YesterdaysClose { get; set; }
        public double Change { get; set; }
        public double PercentChange { get; set; }
        public double Delay { get; set; }
        public double TradeTimestampv { get; set; }
        public DateTime TradeDate { get; set; }
        public DateTime TradeTime { get; set; }
        public DateTime CurrentDate { get; set; }
        public DateTime CurrentTime { get; set; }
    }

    class GoogleFinancial
    {
        static Random _random;

        public GoogleFinancial()
        {
            _random = new Random();
        }

        public void GetRealStock(List<RealStockTopic> topics)
        {
            //找到所有的股票代码
            List<String> allStockCode = topics.ConvertAll(x => x.StockCode).Distinct().ToList();

            XDocument doc = XDocument.Load("C:\\Financial.xml");
            Dictionary<String, XElement> returnValue = FetchQuoteElements(doc.Root, allStockCode, "name");
            foreach (RealStockTopic topic in topics)
            {
                if (string.IsNullOrEmpty(topic.StockCode) || string.IsNullOrEmpty(topic.StockInfo))
                    continue;

                if (returnValue.ContainsKey(topic.StockCode))
                {
                    double dbl = 0;
                    string value = GetValue(returnValue[topic.StockCode], topic.StockInfo);
                    if (double.TryParse(value, out dbl))
                        topic.UpdateValue(string.Format("{0} - {1}", value, _random.NextDouble().ToString("F5")));
                }
                else
                    topic.UpdateValue("--");
            }
        }

        public void GetRealStock(RealStockTopic topic)
        {
            if (string.IsNullOrEmpty(topic.StockCode) || string.IsNullOrEmpty(topic.StockInfo))
                return;

            XDocument doc = XDocument.Load("C:\\Financial.xml");
            Dictionary<String, XElement> returnValue = FetchQuoteElement(doc.Root, "symbol", "name");

            if (returnValue.ContainsKey(topic.StockCode))
            {
                double dbl = 0.0;
                string value = GetValue(returnValue[topic.StockCode], topic.StockInfo);
                if (double.TryParse(value, out dbl))
                    topic.UpdateValue(string.Format("{0} - {1}", value, _random.NextDouble().ToString("F5")));
            }
            else
                topic.UpdateValue("--");
        }

        private Dictionary<String, GoogleStockData> FetchQuoteData()
        {
            Dictionary<String, GoogleStockData> allReturnData = new Dictionary<string, GoogleStockData>();
            XDocument doc = XDocument.Load("c;\\Financial.xml");
            foreach (XElement finance in doc.Root.Elements("finance"))
            {
                GoogleStockData gsd = new GoogleStockData();
                gsd.StockSymbol = GetAttriValue(finance, "symbol", "name");
                gsd.CompanyName = GetAttriValue(finance, "company", "name");
                gsd.Last = Convert.ToDouble(GetValue(finance, "last"));
                gsd.High = Convert.ToDouble(GetValue(finance, "high"));
                gsd.Low = Convert.ToDouble(GetValue(finance, "low"));
                gsd.Volume = Convert.ToDouble(GetValue(finance, "volume"));
                if (allReturnData.ContainsKey(gsd.StockSymbol))
                    allReturnData[gsd.StockSymbol] = gsd;
                else
                    allReturnData.Add(gsd.StockSymbol, gsd);
            }
            return allReturnData;
        }

        private Dictionary<String, XElement> FetchQuoteElement(XElement root, string element, string attribute)
        {
            Dictionary<String, XElement> retEles = new Dictionary<string, XElement>();
            IEnumerable<XElement>  elems = root.Elements();
            foreach (XElement elem in elems)
            {
                string attriValue = GetAttriValue(elem, element, attribute);

                if (retEles.ContainsKey(attriValue))
                    retEles[attriValue] = elem;
                else
                    retEles.Add(attriValue, elem);
            }
            return retEles;
        }

        private Dictionary<String, XElement> FetchQuoteElement(XElement root, string element)
        {
            Dictionary<String, XElement> retEles = new Dictionary<string, XElement>();
            IEnumerable<XElement> elems = root.Elements();
            foreach (XElement elem in elems)
            {
                string eleValue = GetValue(elem, element);

                if (retEles.ContainsKey(eleValue))
                    retEles[eleValue] = elem;
                else
                    retEles.Add(eleValue, elem);
            }
            return retEles;
        }

        private Dictionary<String, XElement> FetchQuoteElements(XElement root, List<String> elements, string attribute)
        {
            Dictionary<String, XElement> retEles = new Dictionary<string, XElement>();
            IEnumerable<XElement> elems = root.Elements();
            foreach (XElement elem in elems)
            {
                for(int i=0; i<elements.Count; i++)
                {
                    string attriValue = GetAttriValue(elem, elements[i], attribute);

                    if (retEles.ContainsKey(attriValue))
                        retEles[attriValue] = elem;
                    else
                        retEles.Add(attriValue, elem);
                }
            }
            return retEles;
        }

        private Dictionary<String, XElement> FetchQuoteElements(XElement root, List<String> elements)
        {
            Dictionary<String, XElement> retEles = new Dictionary<string, XElement>();
            IEnumerable<XElement> elems = root.Elements();
            foreach (XElement elem in elems)
            {
                for (int i = 0; i < elements.Count; i++)
                {
                    string attriValue = GetValue(elem, elements[i]);

                    if (retEles.ContainsKey(attriValue))
                        retEles[attriValue] = elem;
                    else
                        retEles.Add(attriValue, elem);
                }
            }
            return retEles;
        }

        private string GetAttriValue(XElement elem, string next, string attribute)
        {
            if (elem.Element(next) == null)
                return null;

            if (elem.Element(next).Attribute(attribute) == null)
                return null;

            return elem.Element(next).Attribute(attribute).Value;
        }

        private string GetValue(XElement elem, string next)
        {
            if (elem.Element(next) == null)
                return null;

            return elem.Element(next).Value;
        }
    }
}
