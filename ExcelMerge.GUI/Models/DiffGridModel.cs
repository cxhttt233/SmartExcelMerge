using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media;
using FastWpfGrid;

namespace ExcelMerge.GUI.Models
{
    public class DiffGridModel : FastGridModelBase
    {
        private const string RemovedPlaceholderText = "///";

        private readonly Action<int, int, string> setCellText;
        private int columnCount;
        private int rowCount;
        private Color? backgroundColor = null;
        private Color? decorationColor = null;
        private CellDecoration decoration = CellDecoration.None;
        private string toolTipText = string.Empty;
        private Dictionary<int, int> rowIndexMap = new Dictionary<int, int>();
        private Dictionary<int, int> columnIndexMap = new Dictionary<int, int>();
        private readonly HashSet<int> manuallyHiddenRows = new HashSet<int>();
        private readonly HashSet<int> manuallyHiddenColumns = new HashSet<int>();
        private readonly HashSet<int> pinnedRows = new HashSet<int>();
        private readonly HashSet<int> pinnedColumns = new HashSet<int>();
        private readonly HashSet<int> diffHiddenRows = new HashSet<int>();

        public int ManuallyHiddenRowCount { get { return manuallyHiddenRows.Count; } }
        public int ManuallyHiddenColumnCount { get { return manuallyHiddenColumns.Count; } }
        public int PinnedRowCount { get { return pinnedRows.Count; } }
        public int PinnedColumnCount { get { return pinnedColumns.Count; } }

        // View-only arrangements: neither sheet differences nor underlying workbooks are modified.
        public bool HideUserRows(IEnumerable<int> indices)
        {
            var valid = indices.Where(i => i >= 0 && i < RowCount).Distinct()
                .Where(i => !manuallyHiddenRows.Contains(i)).ToList();
            if (valid.Count == 0 || RowCount - manuallyHiddenRows.Count - valid.Count < 1)
                return false;

            manuallyHiddenRows.UnionWith(valid);
            RefreshRowArrangement();
            return true;
        }

        public bool HideUserColumns(IEnumerable<int> indices)
        {
            var valid = indices.Where(i => i >= 0 && i < ColumnCount).Distinct()
                .Where(i => !manuallyHiddenColumns.Contains(i)).ToList();
            if (valid.Count == 0 || ColumnCount - manuallyHiddenColumns.Count - valid.Count < 1)
                return false;

            manuallyHiddenColumns.UnionWith(valid);
            RefreshColumnArrangement();
            return true;
        }

        public void RestoreUserRows()
        {
            manuallyHiddenRows.Clear();
            RefreshRowArrangement();
        }

        public void RestoreUserColumns()
        {
            manuallyHiddenColumns.Clear();
            RefreshColumnArrangement();
        }

        public void PinRows(IEnumerable<int> indices)
        {
            pinnedRows.UnionWith(indices.Where(i => i >= 0 && i < RowCount));
            RefreshRowArrangement();
        }

        public void PinColumns(IEnumerable<int> indices)
        {
            pinnedColumns.UnionWith(indices.Where(i => i >= 0 && i < ColumnCount));
            RefreshColumnArrangement();
        }

        public void RestoreRowOrder()
        {
            pinnedRows.Clear();
            RefreshRowArrangement();
        }

        public void RestoreColumnOrder()
        {
            pinnedColumns.Clear();
            RefreshColumnArrangement();
        }

        private void RefreshRowArrangement()
        {
            var hidden = new HashSet<int>(diffHiddenRows);
            hidden.UnionWith(manuallyHiddenRows);
            var frozen = new HashSet<int>(pinnedRows);
            frozen.ExceptWith(hidden);

            rowIndexMap.Clear();
            var visible = frozen.OrderBy(i => i).Concat(
                Enumerable.Range(0, RowCount).Where(i => !hidden.Contains(i) && !frozen.Contains(i)));
            int visual = 0;
            foreach (var index in visible)
                rowIndexMap[visual++] = index;

            SetRowArrange(hidden, frozen);
        }

        private void RefreshColumnArrangement()
        {
            var hidden = new HashSet<int>(manuallyHiddenColumns);
            var frozen = new HashSet<int>(pinnedColumns);
            frozen.ExceptWith(hidden);

            columnIndexMap.Clear();
            var visible = frozen.OrderBy(i => i).Concat(
                Enumerable.Range(0, ColumnCount).Where(i => !hidden.Contains(i) && !frozen.Contains(i)));
            int visual = 0;
            foreach (var index in visible)
                columnIndexMap[visual++] = index;

            SetColumnArrange(hidden, frozen);
        }


        public override int ColumnCount
        {
            get { return columnCount; }
        }

        public override int RowCount
        {
            get { return rowCount; }
        }

        public override Color? BackgroundColor
        {
            get { return backgroundColor; }
        }

        public override Color? DecorationColor
        {
            get { return decorationColor; }
        }

        public override CellDecoration Decoration
        {
            get { return decoration; }
        }

        public override string ToolTipText
        {
            get { return toolTipText; }
        }

        public override TooltipVisibilityMode ToolTipVisibility
        {
            get { return TooltipVisibilityMode.OnlyWhenTrimmed; }
        }

        public int ColumnHeaderIndex { get; private set; }
        public int RowHeaderIndex { get; private set; } = -1;
        public DiffType DiffType { get; private set; }
        public ExcelSheetDiff SheetDiff { get; private set; }

        public bool IsEditable
        {
            get { return DiffType == DiffType.Dest && setCellText != null; }
        }

        public DiffGridModel(ExcelSheetDiff sheetDiff, DiffType type, Action<int, int, string> setCellText = null) : base()
        {
            DiffType = type;
            SheetDiff = sheetDiff;
            this.setCellText = setCellText;

            columnCount = SheetDiff.Rows.Any() ? SheetDiff.Rows.Max(r => r.Value.Cells.Count) : 0;
            rowCount = SheetDiff.Rows.Count();
            
            App.Instance.OnSettingUpdated += () => { InvalidateAll(); };
        }

        public override string GetColumnHeaderText(int column)
        {
            ExcelCellDiff cellDiff;
            if (TryGetCellDiff(ColumnHeaderIndex, column, out cellDiff))
                return GetCellText(cellDiff);

            return string.Empty;
        }

        public override string GetRowHeaderText(int row)
        {
            if (RowHeaderIndex < 0)
                return base.GetRowHeaderText(row);

            ExcelCellDiff cellDiff;
            if (TryGetCellDiff(row, RowHeaderIndex, out cellDiff))
                return GetCellText(cellDiff);

            return string.Empty;
        }

        public override string GetCellText(int row, int column)
        {
            return GetCellText(row, column, false);
        }

        public string GetCellText(int row, int column, bool direct)
        {
            ExcelCellDiff cellDiff;
            if (TryGetCellDiff(row, column, out cellDiff, direct))
                return GetCellText(cellDiff);

            return string.Empty;
        }

        public string GetCellText(FastGridCellAddress address, bool direct)
        {
            if (address.IsEmpty)
                return string.Empty;

            return GetCellText(address.Row.Value, address.Column.Value, true);
        }

        public override void SetCellText(int row, int column, string value)
        {
            int originalRow;
            int originalColumn;
            if (!TryGetEditableCellAddress(row, column, out originalRow, out originalColumn))
                return;

            setCellText(originalRow, originalColumn, value);
        }

        public bool TryGetEditableCellAddress(int row, int column, out int originalRow, out int originalColumn)
        {
            originalRow = -1;
            originalColumn = -1;

            if (!IsEditable)
                return false;

            ExcelCellDiff cellDiff;
            if (!TryGetCellDiff(row, column, out cellDiff))
                return false;

            if (cellDiff.Status == ExcelCellStatus.Removed || IsRemovedPlaceholder(cellDiff))
                return false;

            if (cellDiff.DstCell == null)
                return false;

            originalRow = cellDiff.DstCell.OriginalRowIndex;
            originalColumn = cellDiff.DstCell.OriginalColumnIndex;
            return originalRow >= 0 && originalColumn >= 0;
        }

        public bool TryGetEditableRowAddress(int row, out int originalRow)
        {
            originalRow = -1;

            if (!IsEditable)
                return false;

            ExcelRowDiff rowDiff;
            if (!TryGetRowDiff(row, out rowDiff))
                return false;

            if (rowDiff.IsRemoved())
                return false;

            var cell = rowDiff.Cells.Values.FirstOrDefault(c =>
                c.Status != ExcelCellStatus.Removed
                && c.DstCell != null
                && c.DstCell.OriginalRowIndex >= 0);

            if (cell == null)
                return false;

            originalRow = cell.DstCell.OriginalRowIndex;
            return true;
        }

        private bool TryGetCellDiff(int row, int column, out ExcelCellDiff cellDiff, bool direct = false)
        {
            cellDiff = null;

            ExcelRowDiff rowDiff;
            if (TryGetRowDiff(row, out rowDiff, direct))
            {
                if (direct && columnIndexMap.ContainsKey(column))
                    column = columnIndexMap[column];
                return rowDiff.Cells.TryGetValue(column, out cellDiff);
            }

            return false;
        }

        private bool TryGetRowDiff(int row, out ExcelRowDiff rowDiff, bool direct = false)
        {
            if (direct)
                row = rowIndexMap.ContainsKey(row) ? rowIndexMap[row] : row;

            return SheetDiff.Rows.TryGetValue(row, out rowDiff);
        }

        private string GetCellText(ExcelCellDiff cellDiff)
        {
            if (IsRemovedPlaceholder(cellDiff))
                return RemovedPlaceholderText;

            switch (cellDiff.Status)
            {
                case ExcelCellStatus.None:
                    return cellDiff.SrcCell.Value;
                case ExcelCellStatus.Modified:
                    return DiffType == DiffType.Source ? cellDiff.SrcCell.Value : cellDiff.DstCell.Value;
                case ExcelCellStatus.Added:
                    return DiffType == DiffType.Source ? cellDiff.SrcCell.Value : cellDiff.DstCell.Value;
                case ExcelCellStatus.Removed:
                    return DiffType == DiffType.Source ? cellDiff.SrcCell.Value : cellDiff.DstCell.Value;
            }

            return string.Empty;
        }

        private bool IsRemovedPlaceholder(ExcelCellDiff cellDiff)
        {
            return cellDiff.Status == ExcelCellStatus.Removed
                && DiffType == DiffType.Dest
                && string.IsNullOrEmpty(cellDiff.DstCell.Value);
        }

        private Color? GetColor(ExcelCellStatus status)
        {
            switch (status)
            {
                case ExcelCellStatus.Modified:
                    return App.Instance.Setting.ModifiedColor;
                case ExcelCellStatus.Added:
                    return App.Instance.Setting.AddedColor;
                case ExcelCellStatus.Removed:
                    return App.Instance.Setting.RemovedColor;
            }

            return null;
        }

        public bool IsModifiedRow(int row, bool direct)
        {
            if (direct)
                row = rowIndexMap.ContainsKey(row) ? rowIndexMap[row] : row;

            ExcelRowDiff rowDiff;
            if (SheetDiff.Rows.TryGetValue(row, out rowDiff))
                return rowDiff.IsModified();

            return false;
        }

        public bool IsRemovedRow(int row, bool direct)
        {
            if (direct)
                row = rowIndexMap.ContainsKey(row) ? rowIndexMap[row] : row;

            ExcelRowDiff rowDiff;
            if (SheetDiff.Rows.TryGetValue(row, out rowDiff))
                return rowDiff.IsRemoved();

            return false;
        }

        public bool IsAddedRow(int row, bool direct)
        {
            if (direct)
                row = rowIndexMap.ContainsKey(row) ? rowIndexMap[row] : row;

            ExcelRowDiff rowDiff;
            if (SheetDiff.Rows.TryGetValue(row, out rowDiff))
                return rowDiff.IsAdded();

            return false;
        }

        public override IFastGridCell GetRowHeader(IFastGridView view, int row)
        {
            var header = base.GetRowHeader(view, row) as DiffGridModel;
            if (header == null)
                return header;

            header.backgroundColor = App.Instance.Setting.RowHeaderColor;
            header.decoration = CellDecoration.None;
            header.decorationColor = null;

            return header;
        }

        public override IFastGridCell GetColumnHeader(IFastGridView view, int column)
        {
            var header = base.GetColumnHeader(view, column) as DiffGridModel;
            if (header == null)
                return header;

            header.backgroundColor = App.Instance.Setting.ColumnHeaderColor;
            header.decoration = CellDecoration.None;
            header.decorationColor = null;

            return header;
        }

        public IFastGridCell GetCell(IFastGridView view, int row, int column, bool direct)
        {
            toolTipText = GetCellText(row, column, direct);

            var cell = base.GetCell(view, row, column) as DiffGridModel;
            if (cell == null)
                return cell;

            ExcelCellDiff cellDiff;
            var status = ExcelCellStatus.None;
            var isRemovedPlaceholder = false;
            if (TryGetCellDiff(row, column, out cellDiff, direct))
            {
                status = cellDiff.Status;
                isRemovedPlaceholder = IsRemovedPlaceholder(cellDiff);
            }

            cell.backgroundColor = null;
            cell.decoration = CellDecoration.None;
            cell.decorationColor = null;

            if (App.Instance.Setting.ColorModifiedRow && IsModifiedRow(row, true))
                cell.backgroundColor = App.Instance.Setting.ModifiedRowColor;

            cell.backgroundColor = GetColor(status) ?? cell.backgroundColor;
            if (isRemovedPlaceholder)
            {
                cell.decoration = CellDecoration.DiagonalSlash;
                cell.decorationColor = Colors.DimGray;
            }

            return cell;
        }

        public override IFastGridCell GetCell(IFastGridView view, int row, int column)
        {
            return GetCell(view, row, column, false);
        }

        private FastGridCellAddress GetVisualCellAddress(FastGridCellAddress realCellAddress)
        {
            if (realCellAddress.IsEmpty)
                return FastGridCellAddress.Empty;

            var swapped = rowIndexMap.ToDictionary(i => i.Value, i => i.Key);
            int visualRow;
            if (swapped.TryGetValue(realCellAddress.Row.Value, out visualRow))
            {
                int visualColumn = realCellAddress.Column.Value;
                if (columnIndexMap.Any())
                {
                    var match = columnIndexMap.FirstOrDefault(p => p.Value == realCellAddress.Column.Value);
                    if (!columnIndexMap.ContainsKey(match.Key) || match.Value != realCellAddress.Column.Value)
                        return FastGridCellAddress.Empty;
                    visualColumn = match.Key;
                }
                return new FastGridCellAddress(visualRow, visualColumn, realCellAddress.IsGridHeader);
            }

            return rowIndexMap.Any() ? FastGridCellAddress.Empty : realCellAddress;
        }

        private FastGridCellAddress GetRealCellAddress(FastGridCellAddress visualCellAddress)
        {
            if (visualCellAddress.IsEmpty)
                return FastGridCellAddress.Empty;

            int realRow;
            if (!rowIndexMap.TryGetValue(visualCellAddress.Row.Value, out realRow))
                return rowIndexMap.Any() ? FastGridCellAddress.Empty : visualCellAddress;
            int realColumn;
            if (!columnIndexMap.TryGetValue(visualCellAddress.Column.Value, out realColumn))
                realColumn = visualCellAddress.Column.Value;
            return new FastGridCellAddress(realRow, realColumn, visualCellAddress.IsGridHeader);
        }

        public FastGridCellAddress GetNextModifiedCell(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row + 1, (row, cells) =>
            {
                var next = row == realAddress.Row
                ? cells.Skip(realAddress.Column.Value + 1).FirstOrDefault(c => c.Value.Status == ExcelCellStatus.Modified)
                : cells.FirstOrDefault(c => c.Value.Status == ExcelCellStatus.Modified);

                var address = next.Value != null ? new FastGridCellAddress(row, next.Key) : FastGridCellAddress.Empty;

                return GetVisualCellAddress(address);
            });
        }

        public FastGridCellAddress GetPreviousModifiedCell(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row - 1, (row, cells) =>
            {
                var next = row == realAddress.Row
                    ? cells.Take(realAddress.Column.Value).LastOrDefault(c => c.Value.Status == ExcelCellStatus.Modified)
                    : cells.LastOrDefault(c => c.Value.Status == ExcelCellStatus.Modified);

                var address = next.Value != null ? new FastGridCellAddress(row, next.Key) : FastGridCellAddress.Empty;

                return GetVisualCellAddress(address);
            });
        }

        public FastGridCellAddress GetNextModifiedRow(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row + 1, (row, cells) =>
            {
                if (row == realAddress.Row)
                    return FastGridCellAddress.Empty;

                ExcelRowDiff rowDiff;
                if (TryGetRowDiff(row, out rowDiff, false) && rowDiff.IsModified())
                    return GetVisualCellAddress(new FastGridCellAddress(row, realAddress.Column));

                return FastGridCellAddress.Empty;
            });
        }

        public FastGridCellAddress GetPreviousModifiedRow(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row - 1, (row, cells) =>
            {
                if (row == realAddress.Row)
                    return FastGridCellAddress.Empty;

                ExcelRowDiff rowDiff;
                if (TryGetRowDiff(row, out rowDiff, false) && rowDiff.IsModified())
                    return GetVisualCellAddress(new FastGridCellAddress(row, realAddress.Column));

                return FastGridCellAddress.Empty;
            });
        }

        public FastGridCellAddress GetNextAddedRow(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row + 1, (row, cells) =>
            {
                if (row == realAddress.Row)
                    return FastGridCellAddress.Empty;

                ExcelRowDiff rowDiff;
                if (TryGetRowDiff(row, out rowDiff, false) && rowDiff.IsAdded())
                    return GetVisualCellAddress(new FastGridCellAddress(row, realAddress.Column));

                return FastGridCellAddress.Empty;
            });
        }

        public FastGridCellAddress GetPreviousAddedRow(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row - 1, (row, cells) =>
            {
                if (row == realAddress.Row)
                    return FastGridCellAddress.Empty;

                ExcelRowDiff rowDiff;
                if (TryGetRowDiff(row, out rowDiff, false) && rowDiff.IsAdded())
                    return GetVisualCellAddress(new FastGridCellAddress(row, realAddress.Column));

                return FastGridCellAddress.Empty;
            });
        }

        public FastGridCellAddress GetNextRemovedRow(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row + 1, (row, cells) =>
            {
                if (row == realAddress.Row)
                    return FastGridCellAddress.Empty;

                ExcelRowDiff rowDiff;
                if (TryGetRowDiff(row, out rowDiff, false) && rowDiff.IsRemoved())
                    return GetVisualCellAddress(new FastGridCellAddress(row, realAddress.Column));

                return FastGridCellAddress.Empty;
            });
        }

        public FastGridCellAddress GetPreviousRemovedRow(FastGridCellAddress current)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row - 1, (row, cells) =>
            {
                if (row == realAddress.Row)
                    return FastGridCellAddress.Empty;

                ExcelRowDiff rowDiff;
                if (TryGetRowDiff(row, out rowDiff, false) && rowDiff.IsRemoved())
                    return GetVisualCellAddress(new FastGridCellAddress(row, realAddress.Column));

                return FastGridCellAddress.Empty;
            });
        }

        private bool IsMatch(ExcelCellDiff cell, string text, bool exactMatch, bool caseSensitive, bool useRegex)
        {
            var srcValue = caseSensitive ? cell.SrcCell.Value : cell.SrcCell.Value.ToUpper();
            var dstValue = caseSensitive ? cell.DstCell.Value : cell.DstCell.Value.ToUpper();
            var target = caseSensitive ? text : text.ToUpper();

            if (useRegex)
            {
                var regex = new Regex(target);
                return regex.IsMatch(srcValue) || regex.IsMatch(srcValue);
            }

            if (exactMatch)
                return target == srcValue || target == dstValue;

            return srcValue.Contains(target) || dstValue.Contains(target);
        }

        public FastGridCellAddress GetNextMatchCell(FastGridCellAddress current, string text, bool exactMatch, bool caseSensitive, bool useRegex, bool onlyDiff)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row + 1, (row, cells) =>
            {
                if (onlyDiff && GetVisualCellAddress(realAddress).IsEmpty)
                    return FastGridCellAddress.Empty;

                var next = row == realAddress.Row
                ? cells.Skip(realAddress.Column.Value + 1).FirstOrDefault(c => IsMatch(c.Value, text, exactMatch, caseSensitive, useRegex))
                : cells.FirstOrDefault(c => IsMatch(c.Value, text, exactMatch, caseSensitive, useRegex));

                var address = next.Value != null ? new FastGridCellAddress(row, next.Key) : FastGridCellAddress.Empty;

                return GetVisualCellAddress(address);
            });
        }

        public FastGridCellAddress GetPreviousMatchCell(FastGridCellAddress current, string text, bool exactMatch, bool caseSensitive, bool useRegex, bool onlyDiff)
        {
            var realAddress = GetRealCellAddress(current);
            return GetNextCell(realAddress, (row) => row - 1, (row, cells) =>
            {
                if (onlyDiff && GetVisualCellAddress(realAddress).IsEmpty)
                    return FastGridCellAddress.Empty;

                var next = row == realAddress.Row
                    ? cells.Take(realAddress.Column.Value).LastOrDefault(c => IsMatch(c.Value, text, exactMatch, caseSensitive, useRegex))
                    : cells.LastOrDefault(c => IsMatch(c.Value, text, exactMatch, caseSensitive, useRegex));

                var address = next.Value != null ? new FastGridCellAddress(row, next.Key) : FastGridCellAddress.Empty;

                return GetVisualCellAddress(address);
            });
        }

        public FastGridCellAddress GetNextCell(
                FastGridCellAddress current, Func<int, int> rowSelector, Func<int, SortedDictionary<int, ExcelCellDiff>, FastGridCellAddress> selector)
        {
            if (current.IsEmpty)
                return current;

            var rowIndex = current.Row.Value;
            while (true)
            {
                ExcelRowDiff rowDiff;
                if (!TryGetRowDiff(rowIndex, out rowDiff, false))
                    break;

                var next = selector(rowIndex, rowDiff.Cells);
                if (!next.IsEmpty)
                    return next;

                rowIndex = rowSelector(rowIndex);
            }

            return FastGridCellAddress.Empty;
        }

        public void SetRowHeader(int? column)
        {
            if (column.HasValue)
                RowHeaderIndex = column.Value;
        }

        public void SetRowHeader(string columnHeaderName)
        {
            for (int i = 0; i < columnCount; i++)
            {
                if (columnHeaderName == GetColumnHeaderText(i))
                {
                    SetRowHeader(i);
                    return;
                }
            }
        }

        public void SetColumnHeader(int? row)
        {
            if (row.HasValue)
                ColumnHeaderIndex = row.Value;
        }

        public void HideEqualRows()
        {
            HideEqualRows(true, true, true);
        }

        public void HideEqualRows(bool showAdded, bool showModified, bool showRemoved)
        {
            diffHiddenRows.Clear();
            foreach (var entry in SheetDiff.Rows)
            {
                var row = entry.Value;
                bool added = row.IsAdded();
                bool removed = row.IsRemoved();
                bool modified = row.IsModified() || row.Cells.Any(c => c.Value.Status == ExcelCellStatus.Modified);
                bool visible = added ? showAdded : removed ? showRemoved : modified && showModified;
                if (!visible)
                    diffHiddenRows.Add(entry.Key);
            }
            RefreshRowArrangement();
        }

        public void ShowEqualRows()
        {
            diffHiddenRows.Clear();
            RefreshRowArrangement();
        }

    }
}
