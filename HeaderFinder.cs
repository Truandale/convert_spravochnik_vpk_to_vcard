using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace convert_spravochnik_vpk_to_vcard
{
    /// <summary>
    /// Восстанавливает логику поиска заголовков как в коммите 18b487d.
    /// Ищет строку заголовков в первых 6 строках, находит колонки по имени.
    /// </summary>
    public static class HeaderFinder
    {
        /// <summary>
        /// Находит строку заголовков в первых 6 строках листа
        /// </summary>
        public static (int headerRowIndex, string[] headersRaw, string[] headersCanon) FindHeaderRow(ISheet sheet)
        {
            // Ищем в первых 6 строках ту, где больше всего заполненных колонок
            int bestRow = 0;
            int maxCols = 0;
            
            for (int r = 0; r < Math.Min(6, sheet.LastRowNum + 1); r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                
                int filledCols = 0;
                for (int c = 0; c < row.LastCellNum; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                        filledCols++;
                }
                
                if (filledCols > maxCols)
                {
                    maxCols = filledCols;
                    bestRow = r;
                }
            }
            
            var headerRow = sheet.GetRow(bestRow);
            if (headerRow == null)
                return (0, new string[0], new string[0]);
                
            var raw = new List<string>();
            var canon = new List<string>();
            
            for (int c = 0; c < headerRow.LastCellNum; c++)
            {
                var cellValue = headerRow.GetCell(c)?.ToString() ?? "";
                raw.Add(cellValue);
                canon.Add(NormalizeHeader(cellValue));
            }
            
            return (bestRow, raw.ToArray(), canon.ToArray());
        }
        
        /// <summary>
        /// Мапит колонки по именам для конкретной кнопки
        /// </summary>
        public static Dictionary<string, int> MapColumns(string button, string[] headersRaw, string[] headersCanon)
        {
            var result = new Dictionary<string, int>();
            
            // Универсальные варианты заголовков
            var variants = new Dictionary<string, string[]>
            {
                ["fio"] = new[] { "фио", "ф.и.о", "фио сотрудника" },
                ["title"] = new[] { "должность", "роль", "position", "title" },
                ["email"] = new[] { "электронный адрес", "email", "e-mail", "почта" },
                ["mobile"] = new[] { "мобильный номер", "мобильный", "сотовый", "телефон мобильный" },
                ["code"] = new[] { "код города", "городской код", "код" },
                ["city"] = new[] { "городской номер", "городской", "телефон городской" },
                ["ext"] = new[] { "внутренний телефон", "внутренний", "доб", "доб.", "внутр. номер телефона" },
                ["cell"] = new[] { "контактный телефон", "контактный", "телефон" },
                ["extra"] = new[] { "дополнительный номер/ e-mail", "дополнительный номер/e-mail", "дополнительный номер", "доп. номер", "доп номер" },
                ["org"] = new[] { "организация" },
                ["department"] = new[] { "структурное подразделение/ департамент", "структурное подразделение", "департамент", "отдел", "служба", "подразделение" }
            };
            
            // Для каждого поля ищем первое совпадение
            foreach (var field in variants.Keys)
            {
                var fieldVariants = variants[field];
                for (int i = 0; i < headersCanon.Length; i++)
                {
                    if (fieldVariants.Any(v => NormalizeHeader(v) == headersCanon[i]))
                    {
                        result[field] = i;
                        break;
                    }
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// Нормализация заголовка (как было в 18b487d)
        /// </summary>
        private static string NormalizeHeader(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            
            var result = s.ToLowerInvariant()
                         .Replace('ё', 'е')
                         .Replace('\u00A0', ' ')
                         .Replace('\r', ' ')
                         .Replace('\n', ' ')
                         .Replace("e-mail", "email")
                         .Replace("e mail", "email")
                         .Trim();
            
            // Убираем точки и схлопываем пробелы
            result = result.Replace(".", "");
            while (result.Contains("  "))
                result = result.Replace("  ", " ");
                
            return result;
        }
        
        /// <summary>
        /// Безопасное чтение ячейки
        /// </summary>
        public static string ReadCell(IRow row, int columnIndex)
        {
            if (row == null || columnIndex < 0 || columnIndex >= row.LastCellNum)
                return "";
                
            var cell = row.GetCell(columnIndex);
            return cell?.ToString()?.Trim() ?? "";
        }
    }
}
