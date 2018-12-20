using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public static class ApplicationState
    {
        public static bool HasRequiredSheets { get; set; } = false;
        public static bool HasMonthlyPlanSheet { get; set; } = false;
        public static bool HasTrackingSheets { get; set; } = false;
        public static FileInfo File { get; set; } = null;
        public static string UserId { get; set; } = string.Empty;
        public static State State { get; set; } = State.None;
        public static void Reset()
        {
            HasRequiredSheets = false;
            HasMonthlyPlanSheet = false;
            HasTrackingSheets = false;
            File = null;
            UserId = string.Empty;
            State = State.None;

            LogInfo($"Application state has been reseted");
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
