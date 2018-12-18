using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Config
{
    public class ConfigModel
    {
        public string connectionString { get; set; }
        public string[] mainsheets { get; set; }
        public string[] monthlysheet { get; set; }
        public string[] trackingsheets { get; set; }
    }
}





