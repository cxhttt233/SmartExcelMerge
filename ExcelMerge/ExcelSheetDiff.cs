using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelSheetDiff
    {
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }
        public SortedDictionary<int, ExcelColumnStatus> Columns { get; private set; }

        public ExcelSheetDiff()
        {
            Rows = new SortedDictionary<int, ExcelRowDiff>();
            Columns = new SortedDictionary<int, ExcelColumnStatus>();
        }

        public ExcelRowDiff CreateRow(ExcelRowStatus? status = null)
        {
            var row = new ExcelRowDiff(Rows.Any() ? Rows.Keys.Last() + 1 : 0, status);
            Rows.Add(row.Index, row);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            var addedColumnCount = Columns.Count(c => c.Value == ExcelColumnStatus.Inserted);
            var removedColumnCount = Columns.Count(c => c.Value == ExcelColumnStatus.Deleted);
            foreach (var row in Rows)
            {
                if (row.Value.IsAdded())
                    addedRowCount++;
                else if (row.Value.IsRemoved())
                    removedRowCount++;

                if (row.Value.IsModified())
                    modifiedRowCount++;

                modifiedCellCount += row.Value.ModifiedCellCount;
            }

            return new ExcelSheetDiffSummary
            {
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
                AddedColumnCount = addedColumnCount,
                RemovedColumnCount = removedColumnCount,
            };
        }
    }
}
