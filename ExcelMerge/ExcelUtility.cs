using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public class ExcelUtility
    {
        public static object GetCellValue(ICell cell)
        {
            if (cell == null)
                return null;

            return GetCellValue(cell, cell.CellType);
        }

        private static object GetCellValue(ICell cell, CellType type)
        {
            if (cell != null)
            {
                switch (type)
                {
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            return cell.DateCellValue;
                        }
                        else
                        {
                            return cell.NumericCellValue;
                        }
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue;
                    case CellType.Formula:
                        return GetCellValue(cell, cell.CachedFormulaResultType);
                }
            }

            return string.Empty;
        }

        public static string GetCellStringValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;

            return GetCellValue(cell).ToString();
        }

        public static void CreateWorkbook(string path, ExcelWorkbookType workbookType)
        {
            if (!ValidateExtension(path, workbookType))
                throw new ArgumentException("The specified Excel type and path extension do not match.");

            var workbook = CreateWorkbook(workbookType);
            try
            {
                var sheet = workbook.CreateSheet();

                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    workbook.Write(fileStream);
                }
            }
            finally
            {
                DisposeWorkbook(workbook);
            }
        }

        internal static void DisposeWorkbook(IWorkbook workbook)
        {
            var disposable = workbook as IDisposable;
            if (disposable != null)
                disposable.Dispose();
        }
        private static IWorkbook CreateWorkbook(ExcelWorkbookType workbookType)
        {
            switch (workbookType)
            {
                case ExcelWorkbookType.XLS: return new HSSFWorkbook() as IWorkbook;
                case ExcelWorkbookType.XLSX: return new XSSFWorkbook() as IWorkbook;
                default: break;
            }

            throw new ArgumentException("The specified excel type is not supported instantiating.");
        }

        private static bool ValidateExtension(string path, ExcelWorkbookType workbookType)
        {
            switch (workbookType)
            {
                case ExcelWorkbookType.XLS: return Path.GetExtension(path) == ".xls";
                case ExcelWorkbookType.XLSX: return Path.GetExtension(path) == ".xlsx";
                default: break;
            }

            return false;
        }

        public static ExcelWorkbookType GetWorkbookType(string path)
        {
            var extension = Path.GetExtension(path);
            switch (extension)
            {
                case ".xls": return ExcelWorkbookType.XLS;
                case ".xlsx": return ExcelWorkbookType.XLSX;
                default: break;
            }

            return ExcelWorkbookType.None;
        }

        public static ExcelWorkbookType GetWorkboolTypeStrict(string path)
        {
            var type = GetWorkbookType(path);

            if (type == ExcelWorkbookType.None)
            {
                if (IsXLS(path))
                    type = ExcelWorkbookType.XLS;
                else if (IsXLSX(path))
                    type = ExcelWorkbookType.XLSX;
            }

            return type;
        }

        public static bool IsXLS(string path)
        {
            try
            {
                using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var workbook = WorkbookFactory.Create(input);
                    try
                    {
                        return workbook is HSSFWorkbook;
                    }
                    finally
                    {
                        DisposeWorkbook(workbook);
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        public static bool IsXLSX(string path)
        {
            try
            {
                using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var workbook = WorkbookFactory.Create(input);
                    try
                    {
                        return workbook is XSSFWorkbook;
                    }
                    finally
                    {
                        DisposeWorkbook(workbook);
                    }
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
