using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using FastWpfGrid;
using ExcelMerge.GUI.Models;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView
    {
        private readonly Dictionary<string, ViewLayoutSnapshot> sessionViewLayouts =
            new Dictionary<string, ViewLayoutSnapshot>(StringComparer.OrdinalIgnoreCase);
        private string activeViewLayoutKey;

        private string GetViewLayoutKey(string sourceSheet, string targetSheet)
        {
            return Path.GetFullPath(SrcPathTextBox.Text) + "|" + sourceSheet + "|" +
                   Path.GetFullPath(DstPathTextBox.Text) + "|" + targetSheet;
        }

        private void RememberCurrentViewLayout()
        {
            var model = SrcDataGrid.Model as DiffGridModel;
            if (model != null && !string.IsNullOrEmpty(activeViewLayoutKey))
                sessionViewLayouts[activeViewLayoutKey] = model.CaptureLayout();
        }

        // Never apply saved model indexes to a different row/column alignment.
        private void RestoreSessionViewLayout(string sourceSheet, string targetSheet)
        {
            activeViewLayoutKey = GetViewLayoutKey(sourceSheet, targetSheet);
            ViewLayoutSnapshot snapshot;
            if (!sessionViewLayouts.TryGetValue(activeViewLayoutKey, out snapshot))
                return;

            var models = BothDiffModels().ToList();
            if (models.Count != 2 || !models.All(model => model.TryRestoreLayout(snapshot)))
            {
                sessionViewLayouts.Remove(activeViewLayoutKey);
                if (ViewLayoutStatusTextBlock != null)
                    ViewLayoutStatusTextBlock.ToolTip = "记录对齐结构发生变化，旧的行列隐藏设置已自动丢弃，避免隐藏错误记录。";
            }
        }

        private static string ShortenPreview(string value)
        {
            value = (value ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return value.Length > 34 ? value.Substring(0, 34) + "…" : value;
        }

        private static string ExcelColumnLetters(int index)
        {
            var result = "";
            for (int x = index + 1; x > 0; x = (x - 1) / 26)
                result = (char)('A' + (x - 1) % 26) + result;
            return result;
        }

        private void RestoreSelectedHiddenRows_Click(object sender, RoutedEventArgs e)
        {
            var model = SrcDataGrid.Model as DiffGridModel;
            if (model == null) return;

            var entries = model.GetManuallyHiddenRows().Select(i => new HiddenItemOption
            {
                Index = i,
                Label = string.Format("第 {0} 行  ·  {1}", i + 1,
                    ShortenPreview(string.Join(" | ", Enumerable.Range(0, Math.Min(3, model.ColumnCount))
                        .Select(c => model.GetCellText(i, c)))))
            }).ToList();
            if (entries.Count == 0)
            {
                MessageBox.Show("没有手动隐藏的行。", "SmartMerge");
                return;
            }

            var dialog = new RestoreHiddenItemsWindow("行", entries) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true) return;
            foreach (var item in BothDiffModels())
            {
                if (dialog.RestoreAll) item.RestoreUserRows();
                else item.RestoreSelectedRows(dialog.SelectedIndices);
            }
            RefreshViewLayoutStatus();
        }

        private void RestoreSelectedHiddenColumns_Click(object sender, RoutedEventArgs e)
        {
            var model = SrcDataGrid.Model as DiffGridModel;
            if (model == null) return;

            var entries = model.GetManuallyHiddenColumns().Select(i => new HiddenItemOption
            {
                Index = i,
                Label = string.Format("{0} 列  ·  {1}", ExcelColumnLetters(i),
                    ShortenPreview(model.GetColumnHeaderText(i)))
            }).ToList();
            if (entries.Count == 0)
            {
                MessageBox.Show("没有手动隐藏的列。", "SmartMerge");
                return;
            }

            var dialog = new RestoreHiddenItemsWindow("列", entries) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true) return;
            foreach (var item in BothDiffModels())
            {
                if (dialog.RestoreAll) item.RestoreUserColumns();
                else item.RestoreSelectedColumns(dialog.SelectedIndices);
            }
            RefreshViewLayoutStatus();
        }

        // Resolve the actual pane even when the command lives in a nested context submenu.
        private FastGridControl GetLayoutTargetGrid(object sender)
        {
            DependencyObject node = sender as DependencyObject;
            while (node != null)
            {
                var contextMenu = node as ContextMenu;
                if (contextMenu != null)
                    return contextMenu.PlacementTarget as FastGridControl ?? copyTargetGrid;

                var menuItem = node as MenuItem;
                if (menuItem != null)
                {
                    var owner = ItemsControl.ItemsControlFromItemContainer(menuItem);
                    if (owner != null && !ReferenceEquals(owner, node))
                    {
                        node = owner;
                        continue;
                    }
                }
                node = LogicalTreeHelper.GetParent(node);
            }
            return copyTargetGrid;
        }

        private static List<int> SelectedRows(FastGridControl grid)
        {
            if (grid == null) return new List<int>();
            var selected = grid.GetSelectedModelCells().Where(c => c.Row.HasValue)
                .Select(c => c.Row.Value).Distinct().ToList();
            if (!selected.Any() && grid.CurrentModelCell.Row.HasValue)
                selected.Add(grid.CurrentModelCell.Row.Value);
            return selected;
        }

        private static List<int> SelectedColumns(FastGridControl grid)
        {
            if (grid == null) return new List<int>();
            var selected = grid.GetSelectedModelCells().Where(c => c.Column.HasValue)
                .Select(c => c.Column.Value).Distinct().ToList();
            if (!selected.Any() && grid.CurrentModelCell.Column.HasValue)
                selected.Add(grid.CurrentModelCell.Column.Value);
            return selected;
        }

        private void RefreshViewLayoutStatus()
        {
            var model = SrcDataGrid.Model as DiffGridModel;
            if (model == null || ViewLayoutStatusTextBlock == null) return;

            if (model.ManuallyHiddenRowCount == 0 && model.ManuallyHiddenColumnCount == 0 &&
                model.PinnedRowCount == 0 && model.PinnedColumnCount == 0)
            {
                ViewLayoutStatusTextBlock.Text = string.Empty;
                ViewLayoutStatusTextBlock.ToolTip = "视图调整只影响当前会话中的显示，不会修改源文件。";
                return;
            }
            ViewLayoutStatusTextBlock.Text = string.Format(
                "视图调整：隐藏 {0} 行 / {1} 列，置顶 {2} 行 / 置左 {3} 列",
                model.ManuallyHiddenRowCount, model.ManuallyHiddenColumnCount,
                model.PinnedRowCount, model.PinnedColumnCount);
            ViewLayoutStatusTextBlock.ToolTip =
                "本次会话可在同一对文件和工作表重新比较后恢复视图；记录位置变化时为了安全会自动重置。所有设置只影响显示。";
        }

        private IEnumerable<DiffGridModel> BothDiffModels()
        {
            var src = SrcDataGrid.Model as DiffGridModel;
            var dst = DstDataGrid.Model as DiffGridModel;
            if (src != null) yield return src;
            if (dst != null) yield return dst;
        }

        private void HideSelectedRows_Click(object sender, RoutedEventArgs e)
        {
            var indices = SelectedRows(GetLayoutTargetGrid(sender));
            if (indices.Count == 0) return;

            var models = BothDiffModels().ToList();
            if (models.Count == 0) return;
            if (!models[0].HideUserRows(indices))
            {
                MessageBox.Show("没有可隐藏的行，或该操作会隐藏全部行。", "SmartMerge");
                return;
            }
            foreach (var model in models.Skip(1)) model.HideUserRows(indices);
            RefreshViewLayoutStatus();
        }

        private void HideSelectedColumns_Click(object sender, RoutedEventArgs e)
        {
            var indices = SelectedColumns(GetLayoutTargetGrid(sender));
            if (indices.Count == 0) return;

            var models = BothDiffModels().ToList();
            if (models.Count == 0) return;
            if (!models[0].HideUserColumns(indices))
            {
                MessageBox.Show("没有可隐藏的列，或该操作会隐藏全部列。", "SmartMerge");
                return;
            }
            foreach (var model in models.Skip(1)) model.HideUserColumns(indices);
            RefreshViewLayoutStatus();
        }

        private void PinSelectedRows_Click(object sender, RoutedEventArgs e)
        {
            var indices = SelectedRows(GetLayoutTargetGrid(sender));
            if (indices.Count == 0) return;
            foreach (var model in BothDiffModels()) model.PinRows(indices);
            RefreshViewLayoutStatus();
        }

        private void PinSelectedColumns_Click(object sender, RoutedEventArgs e)
        {
            var indices = SelectedColumns(GetLayoutTargetGrid(sender));
            if (indices.Count == 0) return;
            foreach (var model in BothDiffModels()) model.PinColumns(indices);
            RefreshViewLayoutStatus();
        }

        private void ShowUserRows_Click(object sender, RoutedEventArgs e)
        {
            foreach (var model in BothDiffModels()) model.RestoreUserRows();
            RefreshViewLayoutStatus();
        }

        private void ShowUserColumns_Click(object sender, RoutedEventArgs e)
        {
            foreach (var model in BothDiffModels()) model.RestoreUserColumns();
            RefreshViewLayoutStatus();
        }

        private void RestoreRowOrder_Click(object sender, RoutedEventArgs e)
        {
            foreach (var model in BothDiffModels()) model.RestoreRowOrder();
            RefreshViewLayoutStatus();
        }

        private void RestoreColumnOrder_Click(object sender, RoutedEventArgs e)
        {
            foreach (var model in BothDiffModels()) model.RestoreColumnOrder();
            RefreshViewLayoutStatus();
        }

        private void RestoreViewLayout_Click(object sender, RoutedEventArgs e)
        {
            foreach (var model in BothDiffModels())
            {
                model.RestoreUserRows();
                model.RestoreUserColumns();
                model.RestoreRowOrder();
                model.RestoreColumnOrder();
            }
            RefreshViewLayoutStatus();
        }
    }
}
