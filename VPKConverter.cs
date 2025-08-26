using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace convert_spravochnik_vpk_to_vcard
{
    public static class VPKConverter
    {
        /// <summary>
        /// Экранирует специальные символы в значениях vCard согласно RFC 2426
        /// </summary>
        static string Esc(string? s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\")  // backslash
                    .Replace(";", "\\;")    // semicolon
                    .Replace(",", "\\,")    // comma
                    .Replace("\r\n", "\\n") // CRLF внутри значения
                    .Replace("\n", "\\n")   // LF -> \n
                    .Replace("\r", "");     // одиночный CR убрать
        }
        public static void Convert(string excelFilePath, string vCardFilePath)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "VPK_temp");

            // Создаем временную папку, если она не существует
            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }

            // Генерируем суффикс из случайного набора из шести символов (цифры и буквы)
            string suffix = GenerateRandomSuffix(6);
            // Создаем копию файла во временной папке с добавленным суффиксом
            string tempFileName = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(excelFilePath) + "_" + suffix + Path.GetExtension(excelFilePath));
            File.Copy(excelFilePath, tempFileName, true);

            // Читаем данные из временного Excel файла
            IWorkbook workbook;
            using (FileStream file = new FileStream(tempFileName, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
            }
            var sheet = workbook.GetSheetAt(0);
            var rowCount = sheet.LastRowNum;

            using (var vCardWriter = new StreamWriter(vCardFilePath, false, new UTF8Encoding(false)))
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    var currentRow = sheet.GetRow(row);
                    if (currentRow == null)
                    {
                        continue;
                    }

                    string location = currentRow.GetCell(1)?.ToString() ?? "";
                    string name = currentRow.GetCell(3)?.ToString() ?? "";
                    string position = currentRow.GetCell(4)?.ToString() ?? "";
                    string email = currentRow.GetCell(5)?.ToString() ?? "";
                    string phone = currentRow.GetCell(6)?.ToString() ?? "";
                    string internalPhone = currentRow.GetCell(7)?.ToString() ?? "";

                    // Пропускаем строки с пустыми обязательными полями
                    if (string.IsNullOrEmpty(name) || (string.IsNullOrEmpty(email) && string.IsNullOrEmpty(phone) && string.IsNullOrEmpty(internalPhone)))
                    {
                        continue;
                    }

                    // Очищаем номера телефонов от некорректных символов
                    phone = CleanPhoneNumber(phone);
                    internalPhone = CleanPhoneNumber(internalPhone);

                    // Удаляем переносы строк из FN и заменяем множественные пробелы на один пробел
                    name = Regex.Replace(name.Replace("\n", " ").Replace("\r", " "), @"\s+", " ");

                    // Создаем vCard вручную
                    var vCard = new List<string>
                    {
                        "BEGIN:VCARD",
                        "VERSION:3.0",
                        $"FN:{Esc(name)}",
                        $"ORG:{Esc(name)}",
                        $"TITLE:{Esc(position)}",
                        $"EMAIL;TYPE=INTERNET:{Esc(email)}",
                        $"TEL;WORK;VOICE:{phone}",
                        $"TEL;CELL;VOICE:{internalPhone}",
                        $"ADR;TYPE=WORK:;;{Esc(location)};;;;",
                        "NOTE:Дополнительная информация о контакте",
                        "END:VCARD"
                    };

                    foreach (var line in vCard)
                    {
                        vCardWriter.WriteLine(line);
                    }
                }
            }

            // Удаляем временный файл
            File.Delete(tempFileName);
        }

        private static string CleanPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrEmpty(phoneNumber))
            {
                return phoneNumber;
            }

            // Удаляем все символы, кроме цифр, +, -, и пробелов
            return new string(phoneNumber.Where(c => char.IsDigit(c) || c == '+' || c == '-' || c == ' ').ToArray());
        }

        private static string GenerateRandomSuffix(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
