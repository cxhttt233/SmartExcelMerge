using System.Collections.Generic;
using System;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var actualRowIndex = 0;
            int maxLastCell = 0;
            int processedCells = 0;
            ComparisonTelemetry.Report("reader.begin last_row=" + sheet.LastRowNum +
                " physical_rows=" + sheet.PhysicalNumberOfRows);
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                    continue;

                var cells = new List<ExcelCell>();
                var cellsInRow = Math.Max(0, (int)row.LastCellNum);
                maxLastCell = Math.Max(maxLastCell, cellsInRow);
                processedCells += cellsInRow;
                if (processedCells >= 50000 && processedCells % 50000 < cellsInRow)
                    ComparisonTelemetry.Report("reader.progress row=" + rowIndex +
                        " cells=" + processedCells + " max_columns=" + maxLastCell);
                for (int columnIndex = 0; columnIndex < cellsInRow; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);

                    cells.Add(new ExcelCell(stringValue, columnIndex, rowIndex));
                }

                yield return new ExcelRow(actualRowIndex++, cells);
            }
            ComparisonTelemetry.Report("reader.done rows=" + actualRowIndex +
                " cells=" + processedCells + " max_columns=" + maxLastCell);
        }
    }
}
