﻿using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "retailerpl_actual")]
    public class RetailerPLResults : IModel
    {
        [Order(1)] public int Year { get; set; }
        [Order(2)] public string YearType { get; set; }
        [Order(3)] public int Month { get; set; }
        [Order(4)] public string Retailer { get; set; }
        [Order(5)] public string Banner { get; set; }
        [Order(6)] public string Country { get; set; }
        [Order(7)] public string EAN { get; set; }
        [Order(8)] public decimal SellOutVolumeTotal { get; set; }
        [Order(9)] public decimal SellOutVolumePromo { get; set; }
        [Order(10)] public decimal SellOutVolumeNonPromo { get; set; }
        [Order(11)] public decimal SellOutPriceAverage { get; set; }
        [Order(12)] public decimal SellOutPricePromo { get; set; }
        [Order(13)] public decimal SellOutPriceNonPromo { get; set; }
        [Order(14)] public decimal COGSTotal { get; set; }
        [Order(15)] public decimal COGSPromo { get; set; }
        [Order(16)] public decimal COGSNonPromo { get; set; }
        [Order(17)] public decimal RetailerProfitL1Total { get; set; }
        [Order(18)] public decimal RetailerProfitL1Promo { get; set; }
        [Order(19)] public decimal RetailerProfitL1NonPromo { get; set; }
        [Order(20)] public decimal RetailerCODBTotal { get; set; }
        [Order(21)] public decimal RetailerProfitL2Total { get; set; }
        [Order(22)] public decimal RetailerOverheadTotal { get; set; }
        [Order(23)] public decimal RetailerProfitL3Total { get; set; }
        [Order(24)] public decimal GrossProfit { get; set; }
        [Order(25)] public decimal Revenue { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0
                && YearType == null
                && Month == 0
                && Retailer == null
                && Banner == null
                && Country == null
                && EAN == null
                && SellOutVolumeTotal == 0m
                && SellOutVolumePromo == 0m
                && SellOutVolumeNonPromo == 0m
                && SellOutPriceAverage == 0m
                && SellOutPricePromo == 0m
                && SellOutPriceNonPromo == 0m
                && COGSTotal == 0m
                && COGSPromo == 0m
                && COGSNonPromo == 0m
                && RetailerProfitL1Total == 0m
                && RetailerProfitL1Promo == 0m
                && RetailerProfitL1NonPromo == 0m
                && RetailerCODBTotal == 0m
                && RetailerProfitL2Total == 0m
                && RetailerOverheadTotal == 0m
                && RetailerProfitL3Total == 0m
                && GrossProfit == 0m
                && Revenue == 0m)
                return true;
            else
                return false;
        }
    }
}