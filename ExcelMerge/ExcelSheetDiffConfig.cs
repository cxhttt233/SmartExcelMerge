using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge
{
    public class ExcelSheetDiffConfig
    {
        public ExcelSheetDiffConfig()
        {
            UseSmartTableDiff = true;
            SrcRowHeaderIndex = -1;
            DstRowHeaderIndex = -1;
            SrcRowHeaderName = string.Empty;
            DstRowHeaderName = string.Empty;
            ManualRowKeyNames = new List<string>();
        }

        public int SrcSheetIndex { get; set; }
        public int DstSheetIndex { get; set; }
        public int SrcHeaderIndex { get; set; }
        public int DstHeaderIndex { get; set; }
        public bool UseSmartTableDiff { get; set; }
        public int SrcRowHeaderIndex { get; set; }
        public int DstRowHeaderIndex { get; set; }
        public string SrcRowHeaderName { get; set; }
        public string DstRowHeaderName { get; set; }
        public IList<string> ManualRowKeyNames { get; set; }
        public RowKeyAnalysis RowKeyAnalysis { get; set; }
    }
}
