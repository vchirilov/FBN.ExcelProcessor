using ExcelProcessor.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "sellout_data")]
    public class SellOutData: IModel
    {
        [Order(1)] public int Year { get; set; }
        [Order(2)] public string YearType { get; set; }
        [Order(3)] public string CPG { get; set; }
        [Order(4)] public string Retailer { get; set; }
        [Order(5)] public string Banner { get; set; }
        [Order(6)] public string Country { get; set; }
        [Order(7)] public string CategoryGroup { get; set; }
        [Order(8)] public string Category { get; set; }
        [Order(9)] public string Market { get; set; }
        [Order(10)] public string MarketDesc { get; set; }
        [Order(11)] public string Brand { get; set; }
        [Order(12)] public string Segment { get; set; }
        [Order(13)] public string TargetUser { get; set; }
        [Order(14)] public string MPvsNonMP { get; set; }
        [Order(15)] public string Item { get; set; }
        [Order(16)] public string GlobalBrand { get; set; }
        [Order(17)] public decimal ActualNumberInPack { get; set; }
        [Order(18)] public string Format { get; set; }
        [Order(19)] public string Type { get; set; }
        [Order(20)] public string Form { get; set; }
        [Order(21)] public string PackType { get; set; }
        [Order(22)] public decimal ActualPackSize { get; set; }
        [Order(23)] public string CoreBenefit { get; set; }
        [Order(24)] public string Variant { get; set; }
        [Order(25)] public string TargetArea { get; set; }
        [Order(26)] public string SubSegment { get; set; }
        [Order(27)] public string SubBrand { get; set; }
        [Order(28)] public string EAN { get; set; }
        [Order(29)] public decimal TotalSalesVolume { get; set; }
        [Order(30)] public decimal TotalSalesValue { get; set; }
        [Order(31)] public decimal PromoSalesVolume { get; set; }
        [Order(32)] public decimal PromoSalesValue { get; set; }
        [Order(33)] public decimal WD { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0 &&
                YearType.IsNullOrEmpty() == true &&
                CPG.IsNullOrEmpty() == true &&
                Retailer.IsNullOrEmpty() == true &&
                Banner.IsNullOrEmpty() == true &&
                Country.IsNullOrEmpty() == true &&
                CategoryGroup.IsNullOrEmpty() == true &&
                Category.IsNullOrEmpty() == true &&
                Market.IsNullOrEmpty() == true &&
                MarketDesc.IsNullOrEmpty() == true &&
                Brand.IsNullOrEmpty() == true &&
                Segment.IsNullOrEmpty() == true &&
                TargetUser.IsNullOrEmpty() == true &&
                MPvsNonMP.IsNullOrEmpty() == true &&
                Item.IsNullOrEmpty() == true &&
                GlobalBrand.IsNullOrEmpty() == true &&
                ActualNumberInPack == 0m &&
                Format.IsNullOrEmpty() == true &&
                Type.IsNullOrEmpty() == true &&
                Form.IsNullOrEmpty() == true &&
                PackType.IsNullOrEmpty() == true &&
                ActualPackSize == 0m &&
                CoreBenefit.IsNullOrEmpty() == true &&
                Variant.IsNullOrEmpty() == true &&
                TargetArea.IsNullOrEmpty() == true &&
                SubSegment.IsNullOrEmpty() == true &&
                SubBrand.IsNullOrEmpty() == true &&
                EAN.IsNullOrEmpty() == true &&
                TotalSalesVolume == 0m &&
                TotalSalesValue == 0m &&
                PromoSalesVolume == 0m &&
                PromoSalesValue == 0m &&
                WD == 0m)
                return true;
            else
                return false;
        }

    }
}
