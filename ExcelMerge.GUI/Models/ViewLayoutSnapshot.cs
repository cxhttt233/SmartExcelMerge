using System;

namespace ExcelMerge.GUI.Models
{
    // In-memory only. Never persisted over an unknown workbook revision.
    public sealed class ViewLayoutSnapshot
    {
        public string Signature { get; set; }
        public int[] HiddenRows { get; set; }
        public int[] HiddenColumns { get; set; }
        public int[] PinnedRows { get; set; }
        public int[] PinnedColumns { get; set; }

        public ViewLayoutSnapshot()
        {
            Signature = string.Empty;
            HiddenRows = new int[0];
            HiddenColumns = new int[0];
            PinnedRows = new int[0];
            PinnedColumns = new int[0];
        }
    }
}
