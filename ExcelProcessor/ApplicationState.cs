using ExcelProcessor.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static ExcelProcessor.Helpers.Utility;

namespace ExcelProcessor
{
    public static class ApplicationState
    {
        public static class ImportType
        {
            public static bool IsBase { get; set; } = false;
            public static bool IsMonthly { get; set; } = false;
            public static bool IsTracking { get; set; } = false;
        }

        public static FileInfo File { get; set; } = null;
        public static ImportDetails ImportDetails { get; set;} = null;
        public static State State { get; set; } = State.None;
        public static void Reset()
        {
            LogInfo($"Application state has been reset", false);

            ImportType.IsBase = false;
            ImportType.IsMonthly = false;
            ImportType.IsTracking = false;
            File = null;
            ImportDetails = null;
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
