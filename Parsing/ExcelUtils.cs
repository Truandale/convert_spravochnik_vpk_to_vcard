using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Converter.Parsing
{
    /// <summary>
    /// Утилиты для работы с Excel файлами
    /// </summary>
    public static class ExcelUtils
    {
        /// <summary>
        /// Открывает Excel файл (поддерживает .xls и .xlsx)
        /// </summary>
        public static IWorkbook Open(string filePath)
        {
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".xlsx" => new XSSFWorkbook(fileStream),
                ".xls" => new HSSFWorkbook(fileStream),
                _ => throw new NotSupportedException($"Формат файла {extension} не поддерживается")
            };
        }

        /// <summary>
        /// Получает строковое значение ячейки
        /// </summary>
        public static string GetCellString(IRow row, int cellIndex)
        {
            var cell = row.GetCell(cellIndex);
            if (cell == null) return "";

            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue ?? "",
                CellType.Numeric => cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Formula => GetFormulaValue(cell),
                _ => ""
            };
        }

        private static string GetFormulaValue(ICell cell)
        {
            try
            {
                return cell.CachedFormulaResultType switch
                {
                    CellType.String => cell.StringCellValue ?? "",
                    CellType.Numeric => cell.NumericCellValue.ToString(),
                    CellType.Boolean => cell.BooleanCellValue.ToString(),
                    _ => ""
                };
            }
            catch
            {
                return "";
            }
        }
    }
}
