using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelMerge;

namespace ExcelMerge.GUI.Views
{
    public sealed class KeyAnalysisWindow : Window
    {
        private sealed class Item
        {
            private bool selected;

            public Action SelectionChanged { get; set; }
            public bool Selected
            {
                get { return selected; }
                set
                {
                    if (selected == value)
                        return;
                    selected = value;
                    if (SelectionChanged != null)
                        SelectionChanged();
                }
            }

            public string Field { get; set; }
            public string SourceCoverage { get; set; }
            public string DestinationCoverage { get; set; }
            public string SourceUnique { get; set; }
            public string DestinationUnique { get; set; }
            public string Overlap { get; set; }
            public string Score { get; set; }
            public string Reason { get; set; }
        }

        private readonly RowKeyAnalysis analysis;
        private readonly int srcSheetIndex;
        private readonly int dstSheetIndex;
        private readonly RadioButton auto;
        private readonly RadioButton manual;
        private readonly DataGrid grid;
        private readonly List<Item> items;
        private readonly TextBlock previewText;
        private readonly Button confirm;

        public bool UseManualSelection { get; private set; }
        public IList<string> SelectedColumns { get; private set; }
        public string PreviewSummaryText { get { return previewText == null ? string.Empty : previewText.Text; } }

        public KeyAnalysisWindow(
            RowKeyAnalysis analysis,
            IEnumerable<string> selected,
            int srcSheetIndex,
            int dstSheetIndex)
        {
            this.analysis = analysis ?? new RowKeyAnalysis();
            this.srcSheetIndex = srcSheetIndex;
            this.dstSheetIndex = dstSheetIndex;
            var selectedSet = new HashSet<string>(selected ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);

            Title = "主键分析与选择";
            Width = 980;
            Height = 650;
            MinWidth = 760;
            MinHeight = 500;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;

            var root = new Grid { Margin = new Thickness(12) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            var top = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(top, 0);
            root.Children.Add(top);
            top.Children.Add(new TextBlock
            {
                Text = Summary(this.analysis),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8),
                Padding = new Thickness(8),
                Background = SystemColors.ControlLightBrush
            });

            var modes = new StackPanel { Orientation = Orientation.Horizontal };
            auto = new RadioButton
            {
                Content = "自动选择",
                IsChecked = selectedSet.Count == 0,
                Margin = new Thickness(0, 0, 16, 0)
            };
            manual = new RadioButton
            {
                Content = "手动选择（勾选一个字段为单主键，多个字段为联合主键）",
                IsChecked = selectedSet.Count > 0
            };
            modes.Children.Add(auto);
            modes.Children.Add(manual);
            top.Children.Add(modes);

            previewText = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                FontWeight = FontWeights.SemiBold
            };
            top.Children.Add(new Border
            {
                Margin = new Thickness(0, 8, 0, 0),
                Padding = new Thickness(9),
                BorderThickness = new Thickness(1),
                BorderBrush = SystemColors.ActiveBorderBrush,
                Background = SystemColors.ControlLightLightBrush,
                Child = previewText
            });

            items = this.analysis.Candidates.OrderByDescending(candidate => candidate.Score).Select(candidate => new Item
            {
                Selected = candidate.ColumnNames.Count == 1 && selectedSet.Contains(candidate.ColumnNames[0]),
                Field = candidate.DisplayName,
                SourceCoverage = candidate.SourceCoverageRate.ToString("P1"),
                DestinationCoverage = candidate.DestinationCoverageRate.ToString("P1"),
                SourceUnique = candidate.SourceUniqueRate.ToString("P1"),
                DestinationUnique = candidate.DestinationUniqueRate.ToString("P1"),
                Overlap = candidate.OverlapRate.ToString("P1"),
                Score = candidate.Score.ToString("0.00"),
                Reason = candidate.Reason
            }).ToList();

            foreach (var item in items)
                item.SelectionChanged = OnSelectionChanged;

            grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = items
            };
            grid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "选择",
                Binding = new Binding("Selected")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                Width = 52
            });
            Add("字段", "Field", 150);
            Add("左侧非空率", "SourceCoverage", 88);
            Add("右侧非空率", "DestinationCoverage", 88);
            Add("左侧唯一率", "SourceUnique", 88);
            Add("右侧唯一率", "DestinationUnique", 88);
            Add("两表重合率", "Overlap", 88);
            Add("得分", "Score", 65);
            Add("分析结论", "Reason", new DataGridLength(1, DataGridLengthUnitType.Star));
            Grid.SetRow(grid, 1);
            root.Children.Add(grid);

            var bottom = new DockPanel { Margin = new Thickness(0, 10, 0, 0), LastChildFill = true };
            Grid.SetRow(bottom, 2);
            root.Children.Add(bottom);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            DockPanel.SetDock(buttons, Dock.Right);
            bottom.Children.Add(buttons);

            confirm = new Button
            {
                Content = "确定",
                MinWidth = 72,
                Margin = new Thickness(8, 0, 0, 0),
                Padding = new Thickness(8, 3, 8, 3),
                IsDefault = true
            };
            confirm.Click += Confirm;
            buttons.Children.Add(confirm);
            buttons.Children.Add(new Button
            {
                Content = "取消",
                MinWidth = 72,
                Margin = new Thickness(8, 0, 0, 0),
                Padding = new Thickness(8, 3, 8, 3),
                IsCancel = true
            });

            bottom.Children.Add(new TextBlock
            {
                Text = "上方预览会随勾选即时更新；确认后才会刷新主界面的左右表格。",
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            });

            auto.Checked += (sender, args) =>
            {
                grid.IsEnabled = false;
                UpdatePreview();
            };
            manual.Checked += (sender, args) =>
            {
                grid.IsEnabled = true;
                UpdatePreview();
            };
            grid.IsEnabled = selectedSet.Count > 0;
            UpdatePreview();
        }

        private void OnSelectionChanged()
        {
            if (items.Any(item => item.Selected))
                manual.IsChecked = true;
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (previewText == null || confirm == null)
                return;

            if (auto.IsChecked == true)
            {
                confirm.IsEnabled = true;
                previewText.Text = "自动方案（当前结果）\n" + CompactSummary(analysis.SelectedAnalysis, analysis.SelectedDisplayName, false);
                return;
            }

            var selectedColumns = items.Where(item => item.Selected).Select(item => item.Field).ToList();
            if (selectedColumns.Count == 0)
            {
                confirm.IsEnabled = false;
                previewText.Text = "手动主键预览\n请勾选一个字段作为单主键，或勾选多个字段查看联合主键的组合指标。";
                return;
            }

            var preview = RowKeySelectionRuntime.AnalyzeSelection(srcSheetIndex, dstSheetIndex, selectedColumns);
            if (preview == null)
            {
                confirm.IsEnabled = false;
                previewText.Text = "无法生成预览：当前比较数据尚未准备完成。";
                return;
            }

            confirm.IsEnabled = preview.IsUsableManualKey;
            previewText.Text = CompactSummary(preview, string.Join(" + ", selectedColumns), selectedColumns.Count > 1);
        }

        private static string CompactSummary(RowKeyCandidateAnalysis value, string displayName, bool composite)
        {
            if (value == null)
                return "当前未选出可靠主键。";

            var title = composite ? "联合主键预览" : "单主键预览";
            var metricPrefix = composite ? "联合" : string.Empty;
            return string.Format(
                "{0}：{1}\n{2}非空率：左 {3:P1} / 右 {4:P1}\n{2}唯一率：左 {5:P1} / 右 {6:P1}\n两表重合率：{7:P1}｜匹配唯一键：{8:N0}\n结论：{9}",
                title,
                string.IsNullOrWhiteSpace(displayName) ? value.DisplayName : displayName,
                metricPrefix,
                value.SourceCoverageRate,
                value.DestinationCoverageRate,
                value.SourceUniqueRate,
                value.DestinationUniqueRate,
                value.OverlapRate,
                value.OverlapCount,
                value.Reason);
        }

        private void Add(string header, string property, double width)
        {
            Add(header, property, new DataGridLength(width));
        }

        private void Add(string header, string property, DataGridLength width)
        {
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = header,
                Binding = new Binding(property),
                IsReadOnly = true,
                Width = width
            });
        }

        private void Confirm(object sender, RoutedEventArgs e)
        {
            grid.CommitEdit(DataGridEditingUnit.Cell, true);
            grid.CommitEdit(DataGridEditingUnit.Row, true);

            UseManualSelection = manual.IsChecked == true;
            SelectedColumns = items.Where(item => item.Selected).Select(item => item.Field).ToList();
            if (UseManualSelection)
            {
                if (SelectedColumns.Count == 0)
                {
                    MessageBox.Show(this, "手动模式至少选择一个字段。", "主键选择",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var preview = RowKeySelectionRuntime.AnalyzeSelection(srcSheetIndex, dstSheetIndex, SelectedColumns);
                if (preview == null || !preview.IsUsableManualKey)
                {
                    MessageBox.Show(this,
                        preview == null ? "无法分析当前选择。" : preview.Reason,
                        "主键选择",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }
            }

            DialogResult = true;
        }

        private static string Summary(RowKeyAnalysis analysis)
        {
            if (!analysis.HasSelectedKey)
                return "当前主键：未固定\n匹配方式：智能相似度匹配\n" + analysis.SelectionReason;

            var mode = analysis.SelectionMode == RowKeySelectionMode.Manual ? "手动" : "自动";
            var selected = analysis.SelectedAnalysis;
            if (selected == null)
            {
                return string.Format(
                    "当前主键：{0}（{1}）\n两表重合率：{2:P1}；匹配唯一键：{3:N0}\n{4}",
                    analysis.SelectedDisplayName,
                    mode,
                    analysis.SelectedOverlapRate,
                    analysis.MatchedKeyCount,
                    analysis.SelectionReason);
            }

            return string.Format(
                "当前主键：{0}（{1}）\n左侧非空率：{2:P1}；右侧非空率：{3:P1}；左侧唯一率：{4:P1}；右侧唯一率：{5:P1}\n两表重合率：{6:P1}；匹配唯一键：{7:N0}\n{8}",
                analysis.SelectedDisplayName,
                mode,
                selected.SourceCoverageRate,
                selected.DestinationCoverageRate,
                selected.SourceUniqueRate,
                selected.DestinationUniqueRate,
                selected.OverlapRate,
                selected.OverlapCount,
                analysis.SelectionReason);
        }
    }
}
