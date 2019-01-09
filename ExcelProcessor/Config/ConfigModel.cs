using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Config
{
    public class ConfigModel
    {
        public string ConnectionString { get; set; }
        public decimal Margin { get; set; }
        public string[] MainSheets { get; set; }
        public string[] MonthlySheet { get; set; }
        public string[] TrackingSheets { get; set; }
    }
}