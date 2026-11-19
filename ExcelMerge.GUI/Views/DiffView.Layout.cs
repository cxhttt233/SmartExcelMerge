using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using FastWpfGrid;
using ExcelMerge.GUI.Models;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView
    {
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
                return;
            }
            ViewLayoutStatusTextBlock.Text = string.Format(
                "视图调整：隐藏 {0} 行 / {1} 列，置顶 {2} 行 / 置左 {3} 列",
                model.ManuallyHiddenRowCount, model.ManuallyHiddenColumnCount,
                model.PinnedRowCount, model.PinnedColumnCount);
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
