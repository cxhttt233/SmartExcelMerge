using System.Linq;
using System.Collections.Generic;

namespace ExcelMerge
{
    public enum ExcelRowStatus
    {
        None,
        Modified,
        Added,
        Removed,
    }

    public class ExcelRowDiff
    {
        private ExcelRowStatus? status;

        public int Index { get; private set; }
        public SortedDictionary<int, ExcelCellDiff> Cells { get; private set; }
        public ExcelRowStatus Status
        {
            get
            {
                if (status.HasValue)
                    return status.Value;

                if (IsAdded())
                    return ExcelRowStatus.Added;

                if (IsRemoved())
                    return ExcelRowStatus.Removed;

                return Cells.Any(c => c.Value.Status == ExcelCellStatus.Modified)
                    ? ExcelRowStatus.Modified
                    : ExcelRowStatus.None;
            }
        }

        public ExcelRowDiff(int index)
            : this(index, null)
        {
        }

        public ExcelRowDiff(int index, ExcelRowStatus? status)
        {
            Index = index;
            Cells = new SortedDictionary<int, ExcelCellDiff>();
            this.status = status;
        }

        public void SetStatus(ExcelRowStatus status)
        {
            this.status = status;
        }

        public ExcelCellDiff CreateCell(ExcelCell src, ExcelCell dst, int columnIndex, ExcelCellStatus status)
        {
            var cell = new ExcelCellDiff(columnIndex, Index, src, dst, status);
            Cells.Add(cell.ColumnIndex, cell);

            return cell;
        }

        public bool IsModified()
        {
            if (status.HasValue)
                return status.Value == ExcelRowStatus.Modified;

            return Cells.Any(c => c.Value.Status == ExcelCellStatus.Modified);
        }

        public bool IsAdded()
        {
            if (status.HasValue)
                return status.Value == ExcelRowStatus.Added;

            return Cells.All(c => c.Value.Status == ExcelCellStatus.Added);
        }

        public bool IsRemoved()
        {
            if (status.HasValue)
                return status.Value == ExcelRowStatus.Removed;

            return Cells.All(c => c.Value.Status == ExcelCellStatus.Removed);
        }

        public int ModifiedCellCount
        {
            get { return Cells.Count(c => c.Value.Status == ExcelCellStatus.Modified); }
        }
    }
}
