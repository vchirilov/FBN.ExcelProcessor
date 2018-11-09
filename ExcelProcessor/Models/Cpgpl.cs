using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpg_pl")]
    public class Cpgpl: IModel
    {
        public int Year { get; set; }
        public string YearType { get; set; }
        public string Retailer { get; set; }
        public string Banner { get; set; }
        public string Country { get; set; }
        public string EAN { get; set; }
        public string EANDescription { get; set; }
        public decimal SellInVolumeTotal { get; set; }
        public decimal SellInVolumePromo { get; set; }
        public decimal SellInVolumeNonPromo { get; set; }
        public decimal ListPricePerUnit { get; set; }
        public decimal TTSTotal { get; set; }
        public decimal TTSOnTotal { get; set; }
        public decimal TTSOnConditional { get; set; }
        public decimal TTSOnUnConditional { get; set; }
        public decimal TTSOffTotal { get; set; }
        public decimal TTSOffConditional { get; set; }
        public decimal TTSOffUnConditional { get; set; }
        public decimal NetNetPrice { get; set; }
        public decimal CPPTotal { get; set; }
        public decimal CPPOn { get; set; }
        public decimal CPPOff { get; set; }
        public decimal PromoPrice { get; set; }
        public decimal ThreeNetPrice { get; set; }
        public decimal COGSTotal { get; set; }
        public decimal CPGProfitL1Total { get; set; }
        public decimal CPGProfitL1Promo { get; set; }
        public decimal CPGProfitL1NonPromo { get; set; }
        public decimal CODBTotal { get; set; }
        public decimal CPGProfitL2Total { get; set; }
        public decimal OverheadTotal { get; set; }
        public decimal CPGProfitL3Total { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0 
                && YearType == null 
                && Retailer == null 
                && Banner == null 
                && Country == null 
                && EAN == null 
                && EANDescription == null 
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
                && CODBTotal == 0m 
                && CPGProfitL2Total == 0m 
                && OverheadTotal == 0m
                && CPGProfitL3Total == 0m)
                return true;
            else
                return false;
        }
    }
}
