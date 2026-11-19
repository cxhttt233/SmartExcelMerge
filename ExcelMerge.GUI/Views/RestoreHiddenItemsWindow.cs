using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelMerge.GUI.Views
{
    public sealed class HiddenItemOption
    {
        public int Index { get; set; }
        public string Label { get; set; }
    }

    // Compact purpose-built picker: no Excel data is modified.
    public sealed class RestoreHiddenItemsWindow : Window
    {
        private readonly List<HiddenItemOption> options;
        private readonly ListBox list;
        private readonly TextBox search;
        public int[] SelectedIndices { get; private set; }
        public bool RestoreAll { get; private set; }

        public RestoreHiddenItemsWindow(string kind, IEnumerable<HiddenItemOption> items)
        {
            options = items.ToList();
            SelectedIndices = new int[0];
            Title = "恢复隐藏的" + kind;
            Width = 510;
            Height = 450;
            MinWidth = 385;
            MinHeight = 300;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanResize;
            Background = Brushes.White;

            var root = new DockPanel { Margin = new Thickness(14) };
            Content = root;

            var intro = new StackPanel { Margin = new Thickness(0, 0, 0, 9) };
            DockPanel.SetDock(intro, Dock.Top);
            root.Children.Add(intro);
            intro.Children.Add(new TextBlock
            {
                Text = "仅恢复你手动隐藏的" + kind + "，不改变差异筛选，也不会修改原始 Excel。",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 7)
            });
            search = new TextBox { Height = 28, ToolTip = "搜索序号或单元格内容" };
            intro.Children.Add(search);

            var bottom = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            DockPanel.SetDock(bottom, Dock.Bottom);
            root.Children.Add(bottom);

            var selectedButton = new Button { Content = "恢复选中", MinWidth = 95, Padding = new Thickness(9, 5, 9, 5), IsDefault = true };
            selectedButton.Click += (s, e) =>
            {
                SelectedIndices = list.SelectedItems.Cast<HiddenItemOption>().Select(i => i.Index).ToArray();
                if (SelectedIndices.Length == 0)
                {
                    MessageBox.Show(this, "请先选择要恢复的项目，可使用 Ctrl / Shift 多选。", "SmartMerge");
                    return;
                }
                DialogResult = true;
            };
            bottom.Children.Add(selectedButton);

            var allButton = new Button { Content = "全部恢复", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(9, 5, 9, 5) };
            allButton.Click += (s, e) => { RestoreAll = true; DialogResult = true; };
            bottom.Children.Add(allButton);

            var cancelButton = new Button { Content = "取消", IsCancel = true, Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(9, 5, 9, 5) };
            bottom.Children.Add(cancelButton);

            list = new ListBox
            {
                SelectionMode = SelectionMode.Extended,
                DisplayMemberPath = "Label",
                ItemsSource = options
            };
            ScrollViewer.SetCanContentScroll(list, true);
            VirtualizingStackPanel.SetIsVirtualizing(list, true);
            root.Children.Add(list);

            search.TextChanged += (s, e) =>
            {
                string keyword = (search.Text ?? "").Trim();
                list.ItemsSource = keyword.Length == 0
                    ? options
                    : options.Where(item => item.Label.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            };
        }
    }
}
