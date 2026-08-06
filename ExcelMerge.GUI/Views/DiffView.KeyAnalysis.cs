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
        private TextBlock rowKeyPrimaryText;
        private TextBlock rowKeyDetailText;

        protected override void OnInitialized(EventArgs e)
        {
            base.OnInitialized(e);
            Dispatcher.BeginInvoke((Action)InitializeRowKeyPanel, DispatcherPriority.Loaded);
        }

        private void InitializeRowKeyPanel()
        {
            if (rowKeyPrimaryText != null || ToolExpander == null)
                return;

            var headerPanel = ToolExpander.Header as WrapPanel;
            if (headerPanel == null)
                return;

            rowKeyPrimaryText = new TextBlock
            {
                Text = "当前主键：正在分析",
                FontWeight = FontWeights.SemiBold,
                MaxWidth = 405,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            rowKeyDetailText = new TextBlock
            {
                Text = "等待比较结果",
                FontSize = 11,
                Foreground = SystemColors.GrayTextBrush,
                Margin = new Thickness(0, 1, 0, 0),
                MaxWidth = 405,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            var textPanel = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(5, 1, 8, 1)
            };
            textPanel.Children.Add(rowKeyPrimaryText);
            textPanel.Children.Add(rowKeyDetailText);

            var changeButton = new Button
            {
                Content = "更改",
                Margin = new Thickness(0, 0, 6, 0),
                Padding = new Thickness(3, 0, 3, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = SystemColors.HotTrackBrush,
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = "选择自动、单字段或联合主键"
            };
            changeButton.Click += RowKeyAnalysisButton_Click;
            DockPanel.SetDock(changeButton, Dock.Right);

            var content = new DockPanel { LastChildFill = true };
            content.Children.Add(changeButton);
            content.Children.Add(textPanel);

            var group = new GroupBox
            {
                Header = "行匹配",
                Margin = new Thickness(10, 0, 0, 0),
                MinWidth = 390,
                MaxWidth = 560,
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
            if (rowKeyPrimaryText == null || rowKeyDetailText == null
                || SrcSheetCombobox.SelectedIndex < 0 || DstSheetCombobox.SelectedIndex < 0)
                return;

            var analysis = RowKeySelectionRuntime.GetAnalysis(SrcSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedIndex);
            if (ReferenceEquals(analysis, displayedAnalysis))
                return;

            displayedAnalysis = analysis;
            if (analysis == null)
            {
                rowKeyPrimaryText.Text = "当前主键：正在分析";
                rowKeyDetailText.Text = "等待比较结果";
                rowKeyPrimaryText.ToolTip = null;
                rowKeyDetailText.ToolTip = null;
                return;
            }

            if (!analysis.HasSelectedKey)
            {
                rowKeyPrimaryText.Text = "当前主键：未固定";
                rowKeyDetailText.Text = "匹配方式：智能相似度匹配";
                rowKeyPrimaryText.ToolTip = analysis.SelectionReason;
                rowKeyDetailText.ToolTip = analysis.SelectionReason;
                return;
            }

            var mode = analysis.SelectionMode == RowKeySelectionMode.Manual ? "手动" : "自动";
            rowKeyPrimaryText.Text = string.Format("当前主键：{0}（{1}）", analysis.SelectedDisplayName, mode);

            var selected = analysis.SelectedAnalysis;
            if (selected == null)
            {
                rowKeyDetailText.Text = string.Format("重合率 {0:P1}｜匹配 {1:N0} 条", analysis.SelectedOverlapRate, analysis.MatchedKeyCount);
            }
            else
            {
                rowKeyDetailText.Text = string.Format(
                    "唯一率 左 {0:P1} / 右 {1:P1}｜重合率 {2:P1}｜匹配 {3:N0} 条",
                    selected.SourceUniqueRate,
                    selected.DestinationUniqueRate,
                    selected.OverlapRate,
                    selected.OverlapCount);
            }

            var tooltip = BuildRowKeyTooltip(analysis);
            rowKeyPrimaryText.ToolTip = tooltip;
            rowKeyDetailText.ToolTip = tooltip;
        }

        private static string BuildRowKeyTooltip(RowKeyAnalysis analysis)
        {
            if (analysis == null)
                return string.Empty;

            if (!analysis.HasSelectedKey || analysis.SelectedAnalysis == null)
                return analysis.SelectionReason;

            var selected = analysis.SelectedAnalysis;
            return string.Format(
                "当前主键：{0}\n模式：{1}\n左侧非空率：{2:P1}\n右侧非空率：{3:P1}\n左侧唯一率：{4:P1}\n右侧唯一率：{5:P1}\n两表重合率：{6:P1}\n匹配唯一键：{7:N0}\n{8}",
                analysis.SelectedDisplayName,
                analysis.SelectionMode == RowKeySelectionMode.Manual ? "手动" : "自动",
                selected.SourceCoverageRate,
                selected.DestinationCoverageRate,
                selected.SourceUniqueRate,
                selected.DestinationUniqueRate,
                selected.OverlapRate,
                selected.OverlapCount,
                analysis.SelectionReason);
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
            var dialog = new KeyAnalysisWindow(analysis, selected, srcIndex, dstIndex)
            {
                Owner = Window.GetWindow(this)
            };
            if (dialog.ShowDialog() != true)
                return;

            RowKeySelectionRuntime.SetManualSelection(srcIndex, dstIndex,
                dialog.UseManualSelection ? dialog.SelectedColumns : Enumerable.Empty<string>());

            rowKeyPrimaryText.Text = "当前主键：正在重新计算";
            rowKeyDetailText.Text = "正在自动刷新表格对比结果";
            displayedAnalysis = null;
            ExecuteDiff();
            RefreshRowKeySummary();
        }
    }
}
