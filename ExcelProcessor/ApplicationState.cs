using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor
{
    public static class ApplicationState
    {
        public static bool HasRequiredSheets { get; set; } = false;
        public static bool HasMonthlyPlanSheet { get; set; } = false;
        public static bool HasTrackingSheets { get; set; } = false;
        public static string FileName { get; set; } = string.Empty;
        public static string UserId { get; set; } = string.Empty;
        public static State State { get; set; } = State.None;
        public static void Reset()
        {
            HasRequiredSheets = false;
            HasMonthlyPlanSheet = false;
            HasTrackingSheets = false;
            FileName = string.Empty;
            UserId = string.Empty;
            State = State.None;
        }
    }

    public enum State
    {
        None,
        CopyingFile,
        ValidatingWorkbook,
        InitializingWorksheet,
        ValidatingHistoricalData,
        ValidatingUniqueValues,
        ValidatingEANs,
        Loading,
        Finished
    };
}
