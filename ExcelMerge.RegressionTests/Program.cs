using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMerge.GUI.Views;

namespace ExcelMerge.RegressionTests
{
    internal static class Program
    {
        private static int failures;

        [STAThread]
        private static int Main()
        {
            Run("auto", Auto);
            Run("manual single", Single);
            Run("composite", Composite);
            Run("analysis", Analysis);
            Run("invalid fallback", Invalid);
            Run("key analysis window", KeyAnalysisWindowSmoke);
            Console.WriteLine(failures == 0 ? "All regression tests passed." : failures + " failed.");
            return failures == 0 ? 0 : 1;
        }

        private static ExcelSheet Sheet(params string[][] values)
        {
            var sheet = new ExcelSheet();
            for (var row = 0; row < values.Length; row++)
                sheet.Rows.Add(row, new ExcelRow(row, values[row].Select((value, column) => new ExcelCell(value, column, row))));
            return sheet;
        }

        private static ExcelSheetDiffConfig Config()
        {
            return new ExcelSheetDiffConfig
            {
                UseSmartTableDiff = true,
                SrcHeaderIndex = 0,
                DstHeaderIndex = 0,
                SrcSheetIndex = 0,
                DstSheetIndex = 0
            };
        }

        private static void Auto()
        {
            RowKeySelectionRuntime.Clear();
            var config = Config();
            var summary = ExcelSheet.Diff(
                Sheet(new[] { "ID", "Name" }, new[] { "1", "A" }, new[] { "2", "B" }),
                Sheet(new[] { "ID", "Name" }, new[] { "2", "B2" }, new[] { "1", "A" }),
                config).CreateSummary();
            Assert(summary.ModifiedRowCount == 1 && summary.AddedRowCount == 0 && summary.RemovedRowCount == 0, "auto pairing");
            Assert(config.RowKeyAnalysis.SelectedDisplayName == "ID", "auto ID");
        }

        private static void Single()
        {
            RowKeySelectionRuntime.Clear();
            RowKeySelectionRuntime.SetManualSelection(0, 0, new[] { "Code" });
            var config = Config();
            var summary = ExcelSheet.Diff(
                Sheet(new[] { "Seq", "Code", "Name" }, new[] { "1", "A1", "A" }, new[] { "2", "A2", "B" }),
                Sheet(new[] { "Seq", "Code", "Name" }, new[] { "2", "A2", "B2" }, new[] { "1", "A1", "A" }),
                config).CreateSummary();
            Assert(summary.AddedRowCount == 0 && summary.RemovedRowCount == 0, "manual pairing");
            Assert(config.RowKeyAnalysis.SelectionMode == RowKeySelectionMode.Manual, "manual mode");
        }

        private static void Composite()
        {
            RowKeySelectionRuntime.Clear();
            RowKeySelectionRuntime.SetManualSelection(0, 0, new[] { "Region", "Name" });
            var config = Config();
            var summary = ExcelSheet.Diff(
                Sheet(new[] { "Region", "Name", "Value" }, new[] { "N", "Gate", "10" }, new[] { "S", "Gate", "20" }),
                Sheet(new[] { "Region", "Name", "Value" }, new[] { "S", "Gate", "21" }, new[] { "N", "Gate", "10" }),
                config).CreateSummary();
            Assert(summary.ModifiedRowCount == 1 && summary.AddedRowCount == 0 && summary.RemovedRowCount == 0, "composite pairing");
            Assert(config.RowKeyAnalysis.SelectedColumnNames.Count == 2, "composite selected");
        }

        private static void Analysis()
        {
            RowKeySelectionRuntime.Clear();
            var config = Config();
            ExcelSheet.Diff(
                Sheet(new[] { "ID", "Name" }, new[] { "1", "A" }, new[] { "2", "B" }),
                Sheet(new[] { "ID", "Name" }, new[] { "1", "A" }, new[] { "2", "B2" }),
                config);
            Assert(config.RowKeyAnalysis.Candidates.Any(candidate => candidate.DisplayName == "ID"), "candidate");
            Assert(!string.IsNullOrWhiteSpace(config.RowKeyAnalysis.SelectionReason), "reason");
        }

        private static void Invalid()
        {
            RowKeySelectionRuntime.Clear();
            RowKeySelectionRuntime.SetManualSelection(0, 0, new[] { "Missing" });
            var config = Config();
            ExcelSheet.Diff(Sheet(new[] { "ID" }, new[] { "1" }), Sheet(new[] { "ID" }, new[] { "1" }), config);
            Assert(config.RowKeyAnalysis != null && !config.RowKeyAnalysis.HasSelectedKey, "fallback");
        }

        private static void KeyAnalysisWindowSmoke()
        {
            var analysis = new RowKeyAnalysis
            {
                SelectionMode = RowKeySelectionMode.Automatic,
                SelectedColumnNames = new List<string> { "ID" },
                SelectedOverlapRate = 1,
                MatchedKeyCount = 2,
                SelectionReason = "测试自动主键",
                Candidates = new List<RowKeyCandidateAnalysis>
                {
                    new RowKeyCandidateAnalysis
                    {
                        ColumnNames = new List<string> { "ID" },
                        SourceCoverageRate = 1,
                        DestinationCoverageRate = 1,
                        SourceUniqueRate = 1,
                        DestinationUniqueRate = 1,
                        OverlapRate = 1,
                        Score = 12,
                        Reason = "字段名和唯一性均符合"
                    }
                }
            };

            var window = new KeyAnalysisWindow(analysis, new string[0]);
            Assert(window.Title == "主键分析与选择", "window construction");
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                Console.WriteLine("[PASS] " + name);
            }
            catch (Exception exception)
            {
                failures++;
                Console.WriteLine("[FAIL] " + name + ": " + exception);
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}