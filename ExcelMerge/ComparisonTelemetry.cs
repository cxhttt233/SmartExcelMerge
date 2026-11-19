using System;

namespace ExcelMerge
{
    /// <summary>
    /// Optional, non-blocking diagnostic milestones. Core comparison never depends on logging.
    /// No data is uploaded and cell values are never included.
    /// </summary>
    public static class ComparisonTelemetry
    {
        public static event Action<string> StageChanged;

        public static void Report(string phase)
        {
            var handler = StageChanged;
            if (handler == null) return;
            try { handler(phase); }
            catch { /* Telemetry cannot affect comparison correctness. */ }
        }
    }
}
