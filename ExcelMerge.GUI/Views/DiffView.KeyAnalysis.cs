using System;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView
    {
        private DispatcherTimer rowKeyTimer;
        private RowKeyAnalysis displayedAnalysis;

        private void RowKeyPanel_Loaded(object sender, RoutedEventArgs e)
        {
            if (rowKeyTimer != null) return;
            rowKeyTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            rowKeyTimer.Tick += (s, args) => RefreshRowKeySummary();
            rowKeyTimer.Start();
            RefreshRowKeySummary();
        }

        private void RefreshRowKeySummary()
        {
            if (SrcSheetCombobox.SelectedIndex < 0 || DstSheetCombobox.SelectedIndex < 0) return;
            var analysis = RowKeySelectionRuntime.GetAnalysis(SrcSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedIndex);
            if (ReferenceEquals(analysis, displayedAnalysis)) return;
            displayedAnalysis = analysis;
            if (analysis == null || !analysis.HasSelectedKey)
            {
                RowKeySummaryText.Text = "主键：自动分析";
                RowKeySummaryText.ToolTip = analysis == null ? null : analysis.SelectionReason;
                return;
            }
            var mode = analysis.SelectionMode == RowKeySelectionMode.Manual ? "手动" : "自动";
            RowKeySummaryText.Text = string.Format("{0}：{1}（重合 {2:P0}）", mode, analysis.SelectedDisplayName, analysis.SelectedOverlapRate);
            RowKeySummaryText.ToolTip = analysis.SelectionReason;
        }

        private void RowKeyAnalysisButton_Click(object sender, RoutedEventArgs e)
        {
            var srcIndex = SrcSheetCombobox.SelectedIndex; var dstIndex = DstSheetCombobox.SelectedIndex;
            var analysis = RowKeySelectionRuntime.GetAnalysis(srcIndex, dstIndex);
            if (analysis == null) { ExecuteDiff(); analysis = RowKeySelectionRuntime.GetAnalysis(srcIndex, dstIndex); }
            if (analysis == null) return;
            var selected = RowKeySelectionRuntime.GetManualSelection(srcIndex, dstIndex);
            var dialog = new KeyAnalysisWindow(analysis, selected) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true) return;
            RowKeySelectionRuntime.SetManualSelection(srcIndex, dstIndex,
                dialog.UseManualSelection ? dialog.SelectedColumns : Enumerable.Empty<string>());
            displayedAnalysis = null;
            ExecuteDiff();
            RefreshRowKeySummary();
        }
    }
}
