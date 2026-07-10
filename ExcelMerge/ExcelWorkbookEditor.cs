using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public static class ExcelWorkbookEditor
    {
        public static bool IsEditableWorkbook(string path)
        {
            var type = ExcelUtility.GetWorkbookType(path);
            return type == ExcelWorkbookType.XLS || type == ExcelWorkbookType.XLSX;
        }

        public static void SetCellText(string path, string sheetName, int rowIndex, int columnIndex, string value)
        {
            UpdateWorkbook(path, sheetName, sheet =>
            {
                var cell = GetOrCreateCell(sheet, rowIndex, columnIndex);
                SetCellValue(cell, value);
            });
        }

        public static void SetCellTextBlock(string path, string sheetName, int startRowIndex, int startColumnIndex, IEnumerable<IEnumerable<string>> values)
        {
            UpdateWorkbook(path, sheetName, sheet =>
            {
                var rowOffset = 0;
                foreach (var rowValues in values)
                {
                    var columnOffset = 0;
                    foreach (var value in rowValues)
                    {
                        var cell = GetOrCreateCell(sheet, startRowIndex + rowOffset, startColumnIndex + columnOffset);
                        SetCellValue(cell, value);
                        columnOffset++;
                    }

                    rowOffset++;
                }
            });
        }

        public static void SetCellTexts(string path, string sheetName, IEnumerable<Tuple<int, int, string>> values)
        {
            var pendingValues = values
                .Where(v => v.Item1 >= 0 && v.Item2 >= 0)
                .ToList();
            if (!pendingValues.Any())
                return;

            UpdateWorkbook(path, sheetName, sheet =>
            {
                foreach (var value in pendingValues)
                {
                    var cell = GetOrCreateCell(sheet, value.Item1, value.Item2);
                    SetCellValue(cell, value.Item3);
                }
            });
        }

        public static void InsertBlankRows(string path, string sheetName, int rowIndex, int count)
        {
            if (count <= 0)
                return;

            UpdateWorkbook(path, sheetName, sheet =>
            {
                var insertIndex = Math.Max(0, rowIndex);
                var lastRow = sheet.LastRowNum;
                if (insertIndex <= lastRow)
                    sheet.ShiftRows(insertIndex, lastRow, count, true, false);

                var template = sheet.GetRow(Math.Min(Math.Max(insertIndex + count, 0), sheet.LastRowNum))
                    ?? sheet.GetRow(Math.Max(insertIndex - 1, 0));

                for (var i = 0; i < count; i++)
                {
                    var row = sheet.GetRow(insertIndex + i) ?? sheet.CreateRow(insertIndex + i);
                    ApplyRowStyle(template, row, clearValues: true);
                }
            });
        }

        public static void DuplicateRowsBelow(string path, string sheetName, IEnumerable<int> rowIndices)
        {
            var sourceRows = rowIndices.Distinct().OrderBy(i => i).ToList();
            if (!sourceRows.Any())
                return;

            UpdateWorkbook(path, sheetName, sheet =>
            {
                var snapshots = sourceRows.Select(i => RowSnapshot.Create(sheet.GetRow(i))).ToList();
                var insertIndex = sourceRows.Max() + 1;
                var lastRow = sheet.LastRowNum;
                if (insertIndex <= lastRow)
                    sheet.ShiftRows(insertIndex, lastRow, snapshots.Count, true, false);

                for (var i = 0; i < snapshots.Count; i++)
                {
                    var row = sheet.GetRow(insertIndex + i) ?? sheet.CreateRow(insertIndex + i);
                    snapshots[i].ApplyTo(row);
                }
            });
        }

        public static void DeleteRows(string path, string sheetName, IEnumerable<int> rowIndices)
        {
            var rows = rowIndices.Distinct().OrderByDescending(i => i).ToList();
            if (!rows.Any())
                return;

            UpdateWorkbook(path, sheetName, sheet =>
            {
                foreach (var rowIndex in rows)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row != null)
                        sheet.RemoveRow(row);

                    if (rowIndex < sheet.LastRowNum)
                        sheet.ShiftRows(rowIndex + 1, sheet.LastRowNum, -1, true, false);
                }
            });
        }

        public static string SaveWorkingCopyToOriginal(string workingPath, string originalPath)
        {
            if (string.IsNullOrWhiteSpace(originalPath))
                throw new InvalidOperationException("Original destination path is empty.");

            if (!File.Exists(workingPath))
                throw new FileNotFoundException("Edited working copy was not found.", workingPath);

            if (!File.Exists(originalPath))
                throw new FileNotFoundException("Original destination file was not found.", originalPath);

            var backupPath = CreateBackup(originalPath);
            File.Copy(workingPath, originalPath, true);
            return backupPath;
        }

        public static void RestoreOriginalToWorkingCopy(string originalPath, string workingPath)
        {
            if (string.IsNullOrWhiteSpace(originalPath))
                throw new InvalidOperationException("Original destination path is empty.");

            if (!File.Exists(originalPath))
                throw new FileNotFoundException("Original destination file was not found.", originalPath);

            File.Copy(originalPath, workingPath, true);
        }

        private static void UpdateWorkbook(string path, string sheetName, Action<ISheet> update)
        {
            if (!IsEditableWorkbook(path))
                throw new NotSupportedException("Only .xls and .xlsx files can be edited in the diff window.");

            var tempPath = path + ".tmp";
            IWorkbook workbook = null;
            try
            {
                using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = WorkbookFactory.Create(input);
                }

                var sheet = workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0);
                update(sheet);

                using (var output = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    workbook.Write(output);
                }

                ExcelUtility.DisposeWorkbook(workbook);
                workbook = null;

                File.Copy(tempPath, path, true);
            }
            finally
            {
                if (workbook != null)
                    ExcelUtility.DisposeWorkbook(workbook);

                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        private static ICell GetOrCreateCell(ISheet sheet, int rowIndex, int columnIndex)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            return row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        }

        private static void SetCellValue(ICell cell, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                cell.SetCellType(CellType.Blank);
                return;
            }

            if (cell.CellType == CellType.Numeric)
            {
                double numericValue;
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out numericValue)
                    || double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out numericValue))
                {
                    cell.SetCellValue(numericValue);
                    return;
                }
            }

            if (cell.CellType == CellType.Boolean)
            {
                bool boolValue;
                if (bool.TryParse(value, out boolValue))
                {
                    cell.SetCellValue(boolValue);
                    return;
                }
            }

            if (cell.CellType == CellType.Formula)
                cell.SetCellType(CellType.String);

            cell.SetCellValue(value);
        }

        private static void ApplyRowStyle(IRow source, IRow target, bool clearValues)
        {
            if (source == null || target == null)
                return;

            target.Height = source.Height;
            if (source.IsFormatted)
                target.RowStyle = source.RowStyle;

            if (source.LastCellNum <= 0)
                return;

            for (var column = 0; column < source.LastCellNum; column++)
            {
                var sourceCell = source.GetCell(column);
                if (sourceCell == null)
                    continue;

                var targetCell = target.GetCell(column) ?? target.CreateCell(column);
                targetCell.CellStyle = sourceCell.CellStyle;
                if (clearValues)
                    targetCell.SetCellType(CellType.Blank);
            }
        }

        private static string CreateBackup(string originalPath)
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(root))
                root = Path.GetTempPath();

            var backupDir = Path.Combine(root, "ExcelSmartDiff", "p4v-edit-backup", DateTime.Now.ToString("yyyyMMdd_HHmmss_fff"));
            Directory.CreateDirectory(backupDir);

            var backupPath = Path.Combine(backupDir, Path.GetFileName(originalPath));
            File.Copy(originalPath, backupPath, true);
            return backupPath;
        }

        private sealed class RowSnapshot
        {
            private readonly short height;
            private readonly bool isFormatted;
            private readonly ICellStyle rowStyle;
            private readonly List<CellSnapshot> cells;

            private RowSnapshot(IRow row)
            {
                height = row == null ? (short)-1 : row.Height;
                isFormatted = row != null && row.IsFormatted;
                rowStyle = isFormatted ? row.RowStyle : null;
                cells = new List<CellSnapshot>();

                if (row == null || row.LastCellNum <= 0)
                    return;

                for (var i = 0; i < row.LastCellNum; i++)
                {
                    var cell = row.GetCell(i);
                    if (cell != null)
                        cells.Add(CellSnapshot.Create(cell));
                }
            }

            public static RowSnapshot Create(IRow row)
            {
                return new RowSnapshot(row);
            }

            public void ApplyTo(IRow row)
            {
                if (height >= 0)
                    row.Height = height;

                if (isFormatted)
                    row.RowStyle = rowStyle;

                foreach (var cell in cells)
                    cell.ApplyTo(row);
            }
        }

        private sealed class CellSnapshot
        {
            private int columnIndex;
            private CellType cellType;
            private ICellStyle style;
            private string stringValue;
            private double numericValue;
            private bool boolValue;
            private string formula;

            public static CellSnapshot Create(ICell cell)
            {
                var snapshot = new CellSnapshot
                {
                    columnIndex = cell.ColumnIndex,
                    cellType = cell.CellType,
                    style = cell.CellStyle
                };

                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        snapshot.boolValue = cell.BooleanCellValue;
                        break;
                    case CellType.Numeric:
                        snapshot.numericValue = cell.NumericCellValue;
                        break;
                    case CellType.Formula:
                        snapshot.formula = cell.CellFormula;
                        break;
                    case CellType.String:
                        snapshot.stringValue = cell.StringCellValue;
                        break;
                }

                return snapshot;
            }

            public void ApplyTo(IRow row)
            {
                var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
                cell.CellStyle = style;

                switch (cellType)
                {
                    case CellType.Boolean:
                        cell.SetCellValue(boolValue);
                        break;
                    case CellType.Numeric:
                        cell.SetCellValue(numericValue);
                        break;
                    case CellType.Formula:
                        cell.SetCellFormula(formula);
                        break;
                    case CellType.String:
                        cell.SetCellValue(stringValue);
                        break;
                    default:
                        cell.SetCellType(CellType.Blank);
                        break;
                }
            }
        }
    }
}
