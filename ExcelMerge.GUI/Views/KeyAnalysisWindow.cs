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
            Title = "主键分析与选择"; Width = 980; Height = 560; MinWidth = 760; MinHeight = 420;
            WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
            var root = new DockPanel { Margin = new Thickness(12) }; Content = root;
            var top = new StackPanel { Margin = new Thickness(0,0,0,8) }; DockPanel.SetDock(top, Dock.Top); root.Children.Add(top);
            top.Children.Add(new TextBlock { Text = Summary(analysis), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0,0,0,8) });
            var modes = new StackPanel { Orientation = Orientation.Horizontal };
            auto = new RadioButton { Content = "自动选择", IsChecked = selectedSet.Count == 0, Margin = new Thickness(0,0,16,0) };
            manual = new RadioButton { Content = "手动选择（可勾选一个或多个字段）", IsChecked = selectedSet.Count > 0 };
            modes.Children.Add(auto); modes.Children.Add(manual); top.Children.Add(modes);
            items = analysis.Candidates.OrderByDescending(c => c.Score).Select(c => new Item {
                Selected = c.ColumnNames.Count == 1 && selectedSet.Contains(c.ColumnNames[0]), Field = c.DisplayName,
                SourceCoverage = c.SourceCoverageRate.ToString("P1"), DestinationCoverage = c.DestinationCoverageRate.ToString("P1"),
                SourceUnique = c.SourceUniqueRate.ToString("P1"), DestinationUnique = c.DestinationUniqueRate.ToString("P1"),
                Overlap = c.OverlapRate.ToString("P1"), Score = c.Score.ToString("0.00"), Reason = c.Reason }).ToList();
            grid = new DataGrid { AutoGenerateColumns=false, CanUserAddRows=false, CanUserDeleteRows=false, ItemsSource=items };
            grid.Columns.Add(new DataGridCheckBoxColumn { Header="选择", Binding=new Binding("Selected") { Mode=BindingMode.TwoWay }, Width=52 });
            Add("字段","Field",150); Add("左侧非空率","SourceCoverage",88); Add("右侧非空率","DestinationCoverage",88);
            Add("左侧唯一率","SourceUnique",88); Add("右侧唯一率","DestinationUnique",88); Add("两表重合率","Overlap",88);
            Add("得分","Score",65); Add("分析结论","Reason",new DataGridLength(1,DataGridLengthUnitType.Star)); root.Children.Add(grid);
            var bottom = new DockPanel { Margin=new Thickness(0,10,0,0) }; DockPanel.SetDock(bottom,Dock.Bottom); root.Children.Add(bottom);
            bottom.Children.Add(new TextBlock { Text="联合主键按勾选字段的显示顺序组合。优先选择稳定编码，避免使用会随排序变化的序号。", TextWrapping=TextWrapping.Wrap, VerticalAlignment=VerticalAlignment.Center });
            var buttons = new StackPanel { Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Right }; DockPanel.SetDock(buttons,Dock.Right); bottom.Children.Add(buttons);
            var apply = new Button { Content="应用并重新比较", MinWidth=110, Margin=new Thickness(8,0,0,0), Padding=new Thickness(8,3,8,3), IsDefault=true };
            apply.Click += Apply; buttons.Children.Add(apply);
            buttons.Children.Add(new Button { Content="取消", MinWidth=72, Margin=new Thickness(8,0,0,0), Padding=new Thickness(8,3,8,3), IsCancel=true });
            auto.Checked += (s,e) => grid.IsEnabled=false; manual.Checked += (s,e) => grid.IsEnabled=true; grid.IsEnabled = selectedSet.Count > 0;
        }

        private void Add(string header, string property, object width)
        {
            grid.Columns.Add(new DataGridTextColumn { Header=header, Binding=new Binding(property), IsReadOnly=true,
                Width = width is double ? new DataGridLength((double)width) : (DataGridLength)width });
        }

        private void Apply(object sender, RoutedEventArgs e)
        {
            UseManualSelection = manual.IsChecked == true;
            SelectedColumns = items.Where(i => i.Selected).Select(i => i.Field).ToList();
            if (UseManualSelection && SelectedColumns.Count == 0)
            { MessageBox.Show(this,"手动模式至少选择一个字段。","主键选择",MessageBoxButton.OK,MessageBoxImage.Information); return; }
            DialogResult = true;
        }

        private static string Summary(RowKeyAnalysis a)
        {
            if (!a.HasSelectedKey) return "当前匹配方式：未固定主键。\n" + a.SelectionReason;
            return string.Format("当前主键：{0}；两表重合率：{1:P1}；匹配唯一键：{2}。\n{3}", a.SelectedDisplayName, a.SelectedOverlapRate, a.MatchedKeyCount, a.SelectionReason);
        }
    }
}
