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
                return rowDiff.Cells.TryGetValue(column, out cellDiff);

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
                return new FastGridCellAddress(visualRow, realCellAddress.Column.Value, realCellAddress.IsGridHeader);

            return rowIndexMap.Any() ? FastGridCellAddress.Empty : realCellAddress;
        }

        private FastGridCellAddress GetRealCellAddress(FastGridCellAddress visualCellAddress)
        {
            if (visualCellAddress.IsEmpty)
                return FastGridCellAddress.Empty;

            int realRow;
            if (rowIndexMap.TryGetValue(visualCellAddress.Row.Value, out realRow))
                return new FastGridCellAddress(realRow, visualCellAddress.Column.Value, visualCellAddress.IsGridHeader);

            return rowIndexMap.Any() ? FastGridCellAddress.Empty : visualCellAddress;
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
            rowIndexMap.Clear();
            var equalRows = new HashSet<int>();
            var originalIndex = 0;
            var index = 0;

            foreach (var r in SheetDiff.Rows)
            {
                if (!r.Value.Cells.All(c => c.Value.Status == ExcelCellStatus.None))
                    rowIndexMap.Add(index++, originalIndex);
                else
                    equalRows.Add(r.Key);

                originalIndex++;
            }

            SetRowArrange(equalRows, new HashSet<int>());
        }

        public void ShowEqualRows()
        {
            SetRowArrange(new HashSet<int>(), new HashSet<int>());

            rowIndexMap.Clear();
        }
    }
}
