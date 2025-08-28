using System;
using System.Collections.Generic;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace Converter.Parsing
{
    /// <summary>
    /// Строитель временных Excel файлов в формате, совместимом с VPKConverter
    /// </summary>
    public static class VpkNormalizedWorkbookBuilder
    {
        /// <summary>
        /// Создает временный Excel файл в формате VPK из нормализованных строк
        /// </summary>
        /// <param name="rows">Нормализованные строки контактов</param>
        /// <param name="prefix">Префикс для имени временного файла</param>
        /// <returns>Путь к созданному временному файлу</returns>
        public static string BuildTempWorkbook(List<NormalizedContactRow> rows, string prefix)
        {
            var tempPath = Path.Combine(Path.GetTempPath(), "VPK_temp");
            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }

            // Генерируем уникальное имя файла
            string suffix = GenerateRandomSuffix(6);
            string tempFileName = Path.Combine(tempPath, $"{prefix}{suffix}.xlsx");

            // Создаем новую книгу
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");

            // Создаем заголовки (в формате VPK)
            var headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Колонка 0"); // Не используется в VPK
            headerRow.CreateCell(1).SetCellValue("Локация");   // Колонка 1 - location
            headerRow.CreateCell(2).SetCellValue("Колонка 2"); // Не используется в VPK
            headerRow.CreateCell(3).SetCellValue("ФИО");       // Колонка 3 - name
            headerRow.CreateCell(4).SetCellValue("Должность"); // Колонка 4 - position
            headerRow.CreateCell(5).SetCellValue("Email");     // Колонка 5 - email
            headerRow.CreateCell(6).SetCellValue("Телефон");   // Колонка 6 - phone
            headerRow.CreateCell(7).SetCellValue("Внутренний");// Колонка 7 - internal phone

            // Заполняем данные
            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var dataRow = sheet.CreateRow(i + 1);

                dataRow.CreateCell(0).SetCellValue(""); // Пустая колонка 0
                dataRow.CreateCell(1).SetCellValue(row.Location);
                dataRow.CreateCell(2).SetCellValue(""); // Пустая колонка 2
                dataRow.CreateCell(3).SetCellValue(row.Name);
                dataRow.CreateCell(4).SetCellValue(row.Position);
                dataRow.CreateCell(5).SetCellValue(row.Email);
                dataRow.CreateCell(6).SetCellValue(row.Phone);
                dataRow.CreateCell(7).SetCellValue(row.InternalPhone);
            }

            // Сохраняем файл
            using (var fileStream = new FileStream(tempFileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            workbook.Dispose();
            return tempFileName;
        }

        private static string GenerateRandomSuffix(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            var result = new char[length];
            for (int i = 0; i < length; i++)
            {
                result[i] = chars[random.Next(chars.Length)];
            }
            return new string(result);
        }
    }
}
