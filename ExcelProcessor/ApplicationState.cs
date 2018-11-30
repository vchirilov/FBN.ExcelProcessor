using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor
{
    public static class ApplicationState
    {
        public static bool HasRequiredSheets { get; set; } = false;
        public static bool HasMonthlyPlanSheet { get; set; } = false;

        public static void Reset()
        {
            HasRequiredSheets = false;
            HasMonthlyPlanSheet = false;
        }
    }
}
