using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView
    {
        private DispatcherTimer rowKeyTimer;
        private RowKeyAnalysis displayedAnalysis;
        private TextBlock rowKeySummaryText;

        protected override void OnInitialized(EventArgs e)
        {
            base.OnInitialized(e);
            Dispatcher.BeginInvoke((Action)InitializeRowKeyPanel, DispatcherPriority.Loaded);
        }

        private void InitializeRowKeyPanel()
        {
            if (rowKeySummaryText != null || ToolExpander == null)
                return;

            var headerPanel = ToolExpander.Header as WrapPanel;
            if (headerPanel == null)
                return;

            rowKeySummaryText = new TextBlock
            {
                Text = "主键：自动分析",
                Margin = new Thickness(5, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                MaxWidth = 260,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            var button = new Button
            {
                Content = "主键分析...",
                Margin = new Thickness(3),
                Padding = new Thickness(5, 2, 5, 2)
            };
            button.Click += RowKeyAnalysisButton_Click;

            var content = new StackPanel { Orientation = Orientation.Horizontal };
            content.Children.Add(rowKeySummaryText);
            content.Children.Add(button);

            var group = new GroupBox
            {
                Header = "行匹配",
                Margin = new Thickness(10, 0, 0, 0),
                Content = content
            };

            var insertIndex = Math.Min(4, headerPanel.Children.Count);
            headerPanel.Children.Insert(insertIndex, group);

            rowKeyTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            rowKeyTimer.Tick += (s, args) => RefreshRowKeySummary();
            rowKeyTimer.Start();
            RefreshRowKeySummary();
        }

        private void RefreshRowKeySummary()
        {
            if (rowKeySummaryText == null || SrcSheetCombobox.SelectedIndex < 0 || DstSheetCombobox.SelectedIndex < 0)
                return;

            var analysis = RowKeySelectionRuntime.GetAnalysis(SrcSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedIndex);
            if (ReferenceEquals(analysis, displayedAnalysis))
                return;

            displayedAnalysis = analysis;
            if (analysis == null || !analysis.HasSelectedKey)
            {
                rowKeySummaryText.Text = "主键：自动分析";
                rowKeySummaryText.ToolTip = analysis == null ? null : analysis.SelectionReason;
                return;
            }

            var mode = analysis.SelectionMode == RowKeySelectionMode.Manual ? "手动" : "自动";
            rowKeySummaryText.Text = string.Format("{0}：{1}（重合 {2:P0}）", mode, analysis.SelectedDisplayName, analysis.SelectedOverlapRate);
            rowKeySummaryText.ToolTip = analysis.SelectionReason;
        }

        private void RowKeyAnalysisButton_Click(object sender, RoutedEventArgs e)
        {
            var srcIndex = SrcSheetCombobox.SelectedIndex;
            var dstIndex = DstSheetCombobox.SelectedIndex;
            var analysis = RowKeySelectionRuntime.GetAnalysis(srcIndex, dstIndex);
            if (analysis == null)
            {
                ExecuteDiff();
                analysis = RowKeySelectionRuntime.GetAnalysis(srcIndex, dstIndex);
            }
            if (analysis == null)
                return;

            var selected = RowKeySelectionRuntime.GetManualSelection(srcIndex, dstIndex);
            var dialog = new KeyAnalysisWindow(analysis, selected) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true)
                return;

            RowKeySelectionRuntime.SetManualSelection(srcIndex, dstIndex,
                dialog.UseManualSelection ? dialog.SelectedColumns : Enumerable.Empty<string>());
            displayedAnalysis = null;
            ExecuteDiff();
            RefreshRowKeySummary();
        }
    }
}
