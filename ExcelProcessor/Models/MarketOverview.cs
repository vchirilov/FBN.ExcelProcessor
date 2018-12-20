using ExcelProcessor.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "market_overview")]
    class MarketOverview: IModel
    {
        [Order(1)]
        public int Year { get; set; }
        [Order(2)]
        public string YearType { get; set; }
        [Order(3)]
        public string CPG { get; set; }
        [Order(4)]
        public string Retailer { get; set; }
        [Order(5)]
        public string Banner { get; set; }
        [Order(6)]
        public string Country { get; set; }
        [Order(7)]
        public string CategoryGroup { get; set; }
        [Order(8)]
        public string NielsenCategory { get; set; }
        [Order(9)]
        public string Market { get; set; }
        [Order(10)]
        public string MarketDesc { get; set; }
        [Order(11)]
        public string Segment { get; set; }
        [Order(12)]
        public string SubSegment { get; set; }
        [Order(13)]
        public decimal SalesVolume { get; set; }
        [Order(14)]
        public decimal SalesValue { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0
                && YearType.IsNullOrEmpty() == true
                && CPG.IsNullOrEmpty() == true
                && Retailer.IsNullOrEmpty() == true
                && Banner.IsNullOrEmpty() == true
                && Country.IsNullOrEmpty() == true
                && CategoryGroup.IsNullOrEmpty() == true
                && NielsenCategory.IsNullOrEmpty() == true
                && Market.IsNullOrEmpty() == true
                && MarketDesc.IsNullOrEmpty() == true
                && Segment.IsNullOrEmpty() == true
                && SubSegment.IsNullOrEmpty() == true
                && SalesVolume == 0m
                && SalesValue == 0m)
                return true;
            else
                return false;
        }
    }
}
