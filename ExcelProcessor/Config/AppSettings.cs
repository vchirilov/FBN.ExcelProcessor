using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelProcessor.Config
{
    public class AppSettings
    {
        private AppSettings() {}

        private static ConfigModel Instance { get; set; }
        
        public static ConfigModel GetInstance()
        {
            if (Instance == null)
            {
                string jsonText = File.ReadAllText(Path.Combine(Environment.CurrentDirectory, "config.json"));
                Instance = JsonConvert.DeserializeObject<ConfigModel>(jsonText);                
            }

            return Instance;            
        }        
    }
}
