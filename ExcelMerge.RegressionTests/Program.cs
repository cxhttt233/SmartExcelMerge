using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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
            Run("live composite preview", LiveCompositePreview);
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
            Assert(config.RowKeyAnalysis.SelectedAnalysis != null, "auto analysis retained");
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
            Assert(config.RowKeyAnalysis.SelectedAnalysis != null, "manual analysis retained");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.SourceUniqueRate == 1, "manual source unique");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.DestinationUniqueRate == 1, "manual destination unique");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.OverlapRate == 1, "manual overlap");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.OverlapCount == 2, "manual match count");
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
            Assert(config.RowKeyAnalysis.SelectedAnalysis != null, "composite analysis retained");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.SourceUniqueRate == 1, "composite source unique");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.DestinationUniqueRate == 1, "composite destination unique");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.OverlapRate == 1, "composite overlap");
            Assert(config.RowKeyAnalysis.SelectedAnalysis.OverlapCount == 2, "composite match count");
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

        private static void LiveCompositePreview()
        {
            RowKeySelectionRuntime.Clear();
            var config = Config();
            ExcelSheet.Diff(
                Sheet(
                    new[] { "Region", "Name", "Value" },
                    new[] { "N", "Gate", "10" },
                    new[] { "S", "Gate", "20" },
                    new[] { "N", "Dam", "30" }),
                Sheet(
                    new[] { "Region", "Name", "Value" },
                    new[] { "N", "Dam", "31" },
                    new[] { "S", "Gate", "20" },
                    new[] { "N", "Gate", "10" }),
                config);

            var preview = RowKeySelectionRuntime.AnalyzeSelection(0, 0, new[] { "Region", "Name" });
            Assert(preview != null, "preview exists");
            Assert(preview.ColumnNames.Count == 2, "preview composite columns");
            Assert(preview.SourceCoverageRate == 1 && preview.DestinationCoverageRate == 1, "preview composite coverage");
            Assert(preview.SourceUniqueRate == 1 && preview.DestinationUniqueRate == 1, "preview composite unique");
            Assert(preview.OverlapRate == 1 && preview.OverlapCount == 3, "preview composite overlap");
            Assert(preview.IsUsableManualKey, "preview usable");

            var window = new KeyAnalysisWindow(config.RowKeyAnalysis, new[] { "Region", "Name" }, 0, 0);
            Assert(window.PreviewSummaryText.Contains("联合主键预览"), "window composite title");
            Assert(window.PreviewSummaryText.Contains("联合非空率"), "window composite coverage");
            Assert(window.PreviewSummaryText.Contains("联合唯一率"), "window composite uniqueness");
            Assert(window.PreviewSummaryText.Contains("匹配唯一键：3"), "window composite match count");
        }

        private static void KeyAnalysisWindowSmoke()
        {
            RowKeySelectionRuntime.Clear();
            var config = Config();
            ExcelSheet.Diff(
                Sheet(new[] { "ID", "Name" }, new[] { "1", "A" }, new[] { "2", "B" }),
                Sheet(new[] { "ID", "Name" }, new[] { "1", "A" }, new[] { "2", "B2" }),
                config);

            var window = new KeyAnalysisWindow(config.RowKeyAnalysis, new[] { "ID" }, 0, 0);
            Assert(window.Title == "主键分析与选择", "window construction");
            Assert(window.PreviewSummaryText.Contains("单主键预览"), "single preview visible");

            var buttons = LogicalDescendants<Button>(window).ToList();
            Assert(buttons.Any(button => Equals(button.Content, "确定") && button.IsDefault), "confirm button");
            Assert(buttons.Any(button => Equals(button.Content, "取消") && button.IsCancel), "cancel button");
            Assert(!buttons.Any(button => Convert.ToString(button.Content).Contains("重新比较")
                || Convert.ToString(button.Content).Contains("重新分析")
                || Convert.ToString(button.Content).Contains("应用并")), "no redundant action button");
        }

        private static IEnumerable<T> LogicalDescendants<T>(DependencyObject root) where T : DependencyObject
        {
            foreach (var child in LogicalTreeHelper.GetChildren(root))
            {
                var dependencyObject = child as DependencyObject;
                if (dependencyObject == null)
                    continue;

                var typed = dependencyObject as T;
                if (typed != null)
                    yield return typed;

                foreach (var descendant in LogicalDescendants<T>(dependencyObject))
                    yield return descendant;
            }
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
