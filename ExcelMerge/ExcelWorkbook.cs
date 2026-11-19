using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }

        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);

            ComparisonTelemetry.Report("workbook.open " + Path.GetFileName(path));
            var wb = new ExcelWorkbook();
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                ComparisonTelemetry.Report("workbook.parse_npoi " + Path.GetFileName(path) + " bytes=" + input.Length);
                var srcWb = WorkbookFactory.Create(input);
                ComparisonTelemetry.Report("workbook.parsed " + Path.GetFileName(path) + " sheets=" + srcWb.NumberOfSheets);
                try
                {
                    for (int i = 0; i < srcWb.NumberOfSheets; i++)
                    {
                        var srcSheet = srcWb.GetSheetAt(i);
                        ComparisonTelemetry.Report("sheet.read " + Path.GetFileName(path) + " sheet=" + i +
                            " physical_rows=" + srcSheet.PhysicalNumberOfRows + " last_row=" + srcSheet.LastRowNum);
                        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                        var parsedSheet = ExcelSheet.Create(srcSheet, config);
                        stopwatch.Stop();
                        ComparisonTelemetry.Report("sheet.done " + Path.GetFileName(path) + " sheet=" + i +
                            " kept_rows=" + parsedSheet.Rows.Count + " elapsed_ms=" + stopwatch.ElapsedMilliseconds);
                        wb.Sheets.Add(srcSheet.SheetName, parsedSheet);
                    }
                }
                finally
                {
                    ExcelUtility.DisposeWorkbook(srcWb);
                }
            }

            ComparisonTelemetry.Report("workbook.done " + Path.GetFileName(path) + " sheets=" + wb.Sheets.Count);
            return wb;
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            if (Path.GetExtension(path) == ".csv")
            {
                return new[] { System.IO.Path.GetFileName(path) };
            }

            if (Path.GetExtension(path) == ".tsv")
            {
                return new[] { System.IO.Path.GetFileName(path) };
            }

            var names = new List<string>();
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var wb = WorkbookFactory.Create(input);
                try
                {
                    for (int i = 0; i < wb.NumberOfSheets; i++)
                        names.Add(wb.GetSheetAt(i).SheetName);
                }
                finally
                {
                    ExcelUtility.DisposeWorkbook(wb);
                }
            }

            return names;
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromCsv(path, config));

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromTsv(path, config));

            return wb;
        }
    }
}
