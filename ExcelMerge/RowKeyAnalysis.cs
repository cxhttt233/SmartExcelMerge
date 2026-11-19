using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public enum RowKeySelectionMode
    {
        Automatic,
        Manual
    }

    public sealed class RowKeyCandidateAnalysis
    {
        public RowKeyCandidateAnalysis()
        {
            ColumnNames = new List<string>();
            Reason = string.Empty;
        }

        public IList<string> ColumnNames { get; set; }
        public double SourceCoverageRate { get; set; }
        public double DestinationCoverageRate { get; set; }
        public double SourceUniqueRate { get; set; }
        public double DestinationUniqueRate { get; set; }
        public double OverlapRate { get; set; }
        public int OverlapCount { get; set; }
        public double Score { get; set; }
        public bool IsValidAutomaticKey { get; set; }
        public bool IsUsableManualKey { get; set; }
        public bool IsPreferredHeader { get; set; }
        public string Reason { get; set; }
        public string DisplayName { get { return string.Join(" + ", ColumnNames ?? new List<string>()); } }
    }

    public sealed class RowKeyAnalysis
    {
        public RowKeyAnalysis()
        {
            SelectedColumnNames = new List<string>();
            Candidates = new List<RowKeyCandidateAnalysis>();
            SelectionReason = string.Empty;
        }

        public RowKeySelectionMode SelectionMode { get; set; }
        public IList<string> SelectedColumnNames { get; set; }
        public IList<RowKeyCandidateAnalysis> Candidates { get; set; }
        public RowKeyCandidateAnalysis SelectedAnalysis { get; set; }
        public double SelectedScore { get; set; }
        public double SelectedOverlapRate { get; set; }
        public int MatchedKeyCount { get; set; }
        public string SelectionReason { get; set; }
        public bool HasSelectedKey { get { return SelectedColumnNames != null && SelectedColumnNames.Count > 0; } }
        public string SelectedDisplayName { get { return HasSelectedKey ? string.Join(" + ", SelectedColumnNames) : string.Empty; } }

        public double SelectedSourceCoverageRate
        {
            get { return SelectedAnalysis == null ? 0 : SelectedAnalysis.SourceCoverageRate; }
        }

        public double SelectedDestinationCoverageRate
        {
            get { return SelectedAnalysis == null ? 0 : SelectedAnalysis.DestinationCoverageRate; }
        }

        public double SelectedSourceUniqueRate
        {
            get { return SelectedAnalysis == null ? 0 : SelectedAnalysis.SourceUniqueRate; }
        }

        public double SelectedDestinationUniqueRate
        {
            get { return SelectedAnalysis == null ? 0 : SelectedAnalysis.DestinationUniqueRate; }
        }
    }

    internal sealed class RowKeyPreviewContext
    {
        public ExcelSheet Source { get; set; }
        public ExcelSheet Destination { get; set; }
        public int SourceHeaderIndex { get; set; }
        public int DestinationHeaderIndex { get; set; }
    }

    public static class RowKeySelectionRuntime
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, List<string>> Manual = new Dictionary<string, List<string>>();
        private static readonly Dictionary<string, RowKeyAnalysis> Analyses = new Dictionary<string, RowKeyAnalysis>();
        private static readonly Dictionary<string, RowKeyPreviewContext> PreviewContexts = new Dictionary<string, RowKeyPreviewContext>();

        private static string Key(int srcSheetIndex, int dstSheetIndex)
        {
            return srcSheetIndex + ":" + dstSheetIndex;
        }

        public static IList<string> GetManualSelection(int srcSheetIndex, int dstSheetIndex)
        {
            lock (Sync)
            {
                List<string> value;
                return Manual.TryGetValue(Key(srcSheetIndex, dstSheetIndex), out value)
                    ? new List<string>(value) : new List<string>();
            }
        }

        public static void SetManualSelection(int srcSheetIndex, int dstSheetIndex, IEnumerable<string> names)
        {
            var value = (names ?? Enumerable.Empty<string>())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            lock (Sync)
            {
                var key = Key(srcSheetIndex, dstSheetIndex);
                if (value.Count == 0) Manual.Remove(key); else Manual[key] = value;
            }
        }

        public static RowKeyAnalysis GetAnalysis(int srcSheetIndex, int dstSheetIndex)
        {
            lock (Sync)
            {
                RowKeyAnalysis value;
                return Analyses.TryGetValue(Key(srcSheetIndex, dstSheetIndex), out value) ? value : null;
            }
        }

        public static RowKeyCandidateAnalysis AnalyzeSelection(
            int srcSheetIndex,
            int dstSheetIndex,
            IEnumerable<string> names)
        {
            RowKeyPreviewContext context;
            lock (Sync)
            {
                if (!PreviewContexts.TryGetValue(Key(srcSheetIndex, dstSheetIndex), out context))
                    return null;
            }

            return RowKeySelectionEngine.AnalyzeSelection(
                context.Source,
                context.Destination,
                context.SourceHeaderIndex,
                context.DestinationHeaderIndex,
                names);
        }

        internal static void SetAnalysis(int srcSheetIndex, int dstSheetIndex, RowKeyAnalysis analysis)
        {
            lock (Sync) Analyses[Key(srcSheetIndex, dstSheetIndex)] = analysis;
        }

        internal static void SetPreviewContext(
            int srcSheetIndex,
            int dstSheetIndex,
            ExcelSheet source,
            ExcelSheet destination,
            int sourceHeaderIndex,
            int destinationHeaderIndex)
        {
            lock (Sync)
            {
                PreviewContexts[Key(srcSheetIndex, dstSheetIndex)] = new RowKeyPreviewContext
                {
                    Source = source,
                    Destination = destination,
                    SourceHeaderIndex = sourceHeaderIndex,
                    DestinationHeaderIndex = destinationHeaderIndex
                };
            }
        }

        public static void Clear()
        {
            lock (Sync)
            {
                Manual.Clear();
                Analyses.Clear();
                PreviewContexts.Clear();
            }
        }
    }
}
