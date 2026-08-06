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
            public bool Selected { get; set; }
            public string Field { get; set; }
            public string SourceCoverage { get; set; }
            public string DestinationCoverage { get; set; }
            public string SourceUnique { get; set; }
            public string DestinationUnique { get; set; }
            public string Overlap { get; set; }
            public string Score { get; set; }
            public string Reason { get; set; }
        }

        private readonly RadioButton auto;
        private readonly RadioButton manual;
        private readonly DataGrid grid;
        private readonly List<Item> items;
        public bool UseManualSelection { get; private set; }
        public IList<string> SelectedColumns { get; private set; }

        public KeyAnalysisWindow(RowKeyAnalysis analysis, IEnumerable<string> selected)
        {
            analysis = analysis ?? new RowKeyAnalysis();
            var selectedSet = new HashSet<string>(selected ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);

            Title = "主键分析与选择";
            Width = 980;
            Height = 580;
            MinWidth = 760;
            MinHeight = 440;
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
                Text = Summary(analysis),
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
                Content = "手动选择（可勾选一个或多个字段）",
                IsChecked = selectedSet.Count > 0
            };
            modes.Children.Add(auto);
            modes.Children.Add(manual);
            top.Children.Add(modes);

            items = analysis.Candidates.OrderByDescending(candidate => candidate.Score).Select(candidate => new Item
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

            grid = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                ItemsSource = items
            };
            grid.BeginningEdit += (sender, args) =>
            {
                if (args.Column is DataGridCheckBoxColumn)
                    manual.IsChecked = true;
            };
            grid.Columns.Add(new DataGridCheckBoxColumn
            {
                Header = "选择",
                Binding = new Binding("Selected") { Mode = BindingMode.TwoWay },
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

            var confirm = new Button
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
                Text = "勾选字段会自动切换到手动模式。点“确定”后，主界面会自动重新计算并刷新左右表格。",
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            });
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
            if (UseManualSelection && SelectedColumns.Count == 0)
            {
                MessageBox.Show(this, "手动模式至少选择一个字段。", "主键选择",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
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
