using ExcelProcessor.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpgpl_actual")]
    public class CPGPLResults : IModel
    {
        [Order(1)] public int Year { get; set; }
        [Order(2)] public string YearType { get; set; }
        [Order(3)] public int Month { get; set; }
        [Order(4)] public string Retailer { get; set; }
        [Order(5)] public string Banner { get; set; }
        [Order(6)] public string Country { get; set; }
        [Order(7)] public string EAN { get; set; }
        [Order(8)] public decimal SellInVolumeTotal { get; set; }
        [Order(9)] public decimal SellInVolumePromo { get; set; }
        [Order(10)] public decimal SellInVolumeNonPromo { get; set; }
        [Order(11)] public decimal ListPricePerUnit { get; set; }
        [Order(12)] public decimal TTSTotal { get; set; }
        [Order(13)] public decimal TTSOnTotal { get; set; }
        [Order(14)] public decimal TTSOnConditional { get; set; }
        [Order(15)] public decimal TTSOnUnConditional { get; set; }
        [Order(16)] public decimal TTSOffTotal { get; set; }
        [Order(17)] public decimal TTSOffConditional { get; set; }
        [Order(18)] public decimal TTSOffUnConditional { get; set; }
        [Order(19)] public decimal NetNetPrice { get; set; }
        [Order(20)] public decimal CPPTotal { get; set; }
        [Order(21)] public decimal CPPOn { get; set; }
        [Order(22)] public decimal CPPOff { get; set; }
        [Order(23)] public decimal PromoPrice { get; set; }
        [Order(24)] public decimal ThreeNetPrice { get; set; }
        [Order(25)] public decimal COGSTotal { get; set; }
        [Order(26)] public decimal CPGProfitL1Total { get; set; }
        [Order(27)] public decimal CPGProfitL1Promo { get; set; }
        [Order(28)] public decimal CPGProfitL1NonPromo { get; set; }
        [Order(29)] public decimal CPGCODBTotal { get; set; }
        [Order(30)] public decimal CPGProfitL2Total { get; set; }
        [Order(31)] public decimal CPGOverheadTotal { get; set; }
        [Order(32)] public decimal CPGProfitL3Total { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0m
                && YearType.IsNullOrEmpty() == true
                && Month == 0
                && Retailer.IsNullOrEmpty() == true
                && Banner.IsNullOrEmpty() == true
                && Country.IsNullOrEmpty() == true
                && EAN.IsNullOrEmpty() == true
                && SellInVolumeTotal == 0m
                && SellInVolumePromo == 0m
                && SellInVolumeNonPromo == 0m
                && ListPricePerUnit == 0m
                && TTSTotal == 0m
                && TTSOnTotal == 0m
                && TTSOnConditional == 0m
                && TTSOnUnConditional == 0m
                && TTSOffTotal == 0m
                && TTSOffConditional == 0m
                && TTSOffUnConditional == 0m
                && NetNetPrice == 0m
                && CPPTotal == 0m
                && CPPOn == 0m
                && CPPOff == 0m
                && PromoPrice == 0m
                && ThreeNetPrice == 0m
                && COGSTotal == 0m
                && CPGProfitL1Total == 0m
                && CPGProfitL1Promo == 0m
                && CPGProfitL1NonPromo == 0m
                && CPGCODBTotal == 0m
                && CPGProfitL2Total == 0m
                && CPGOverheadTotal == 0m
                && CPGProfitL3Total == 0m)
                return true;
            else
                return false;

        }
    }
}
