using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "market_overview")]
    class MarketOverview: IModel
    {
        public int Year { get; set; }
        public string YearType { get; set; }
        public string CPG { get; set; }
        public string Retailer { get; set; }
        public string Banner { get; set; }
        public string Country { get; set; }
        public string CategoryGroup { get; set; }
        public string NielsenCategory { get; set; }
        public string Market { get; set; }
        public string MarketDesc { get; set; }
        public string Segment { get; set; }
        public string SubSegment { get; set; }
        public decimal SalesVolume { get; set; }
        public decimal SalesValue { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0
                && YearType == null
                && CPG == null
                && Retailer == null
                && Banner == null
                && Country == null
                && CategoryGroup == null
                && NielsenCategory == null
                && Market == null
                && MarketDesc == null
                && Segment == null
                && SubSegment == null
                && SalesVolume == 0m
                && SalesValue == 0m)
                return true;
            else
                return false;
        }
    }
}
