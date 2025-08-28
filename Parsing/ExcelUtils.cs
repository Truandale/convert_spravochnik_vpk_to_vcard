using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
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
            Debug.WriteLine($"[DEBUG] ExcelUtils.Open ВХОДНАЯ ТОЧКА: path={filePath}, ext={Path.GetExtension(filePath)}");
            
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            Debug.WriteLine($"[DEBUG] ExcelUtils.Open: файл открыт, расширение={extension}");
            
            switch (extension)
            {
                case ".xlsx":
                    Debug.WriteLine($"[DEBUG] Создаем XSSFWorkbook для {filePath}");
                    try
                    {
                        var xlsxWorkbook = new XSSFWorkbook(fileStream);
                        Debug.WriteLine($"[DEBUG] XSSFWorkbook успешно создан для {filePath}");
                        return xlsxWorkbook;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ERROR] Ошибка создания XSSFWorkbook: {ex.GetType().Name}: {ex.Message}");
                        throw;
                    }
                case ".xls":
                    Debug.WriteLine($"[DEBUG] Создаем HSSFWorkbook для {filePath}");
                    try
                    {
                        var xlsWorkbook = new HSSFWorkbook(fileStream);
                        Debug.WriteLine($"[DEBUG] HSSFWorkbook успешно создан для {filePath}");
                        return xlsWorkbook;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ERROR] Ошибка создания HSSFWorkbook: {ex.GetType().Name}: {ex.Message}");
                        throw;
                    }
                default:
                    throw new NotSupportedException($"Формат файла {extension} не поддерживается");
            }
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
