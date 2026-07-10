namespace ExcelMerge
{
    public struct ExcelSheetDiffSummary
    {
        public int ModifiedCellCount { get; set; }
        public int AddedRowCount { get; set; }
        public int RemovedRowCount { get; set; }
        public int ModifiedRowCount { get; set; }
        public int AddedColumnCount { get; set; }
        public int RemovedColumnCount { get; set; }

        public bool HasDiff
        {
            get
            {
                return ModifiedCellCount + AddedRowCount + RemovedRowCount + ModifiedRowCount + AddedColumnCount + RemovedColumnCount > 0;
            }
        }
    }
}
