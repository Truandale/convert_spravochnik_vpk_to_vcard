using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

class Program
{
    static void Main()
    {
        string folderPath = @"\\vpk-0-fs01.vpk.local\VPK";
        string searchPattern1 = "СПРАВОЧНИК ВПК*.xlsx";
        string searchPattern2 = "СПРАВОЧНИК ВПК*.xls";
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string tempPath = Path.Combine(Path.GetTempPath(), "VPK_temp");

        try
        {
            // Проверяем доступ к папке
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Папка не существует или недоступна: {folderPath}");
                return;
            }

            // Создаем временную папку, если она не существует
            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }

            // Найдем Excel файл с любым из двух расширений
            string excelFilePath = Directory.GetFiles(folderPath, searchPattern1).FirstOrDefault() ??
                                   Directory.GetFiles(folderPath, searchPattern2).FirstOrDefault();

            if (excelFilePath == null)
            {
                Console.WriteLine("Excel файл не найден.");
                return;
            }

            Console.WriteLine($"Найден Excel файл: {excelFilePath}");

            // Генерируем суффикс из случайного набора из шести символов (цифры и буквы)
            string suffix = GenerateRandomSuffix(6);
            // Создаем копию файла во временной папке с добавленным суффиксом
            string tempFileName = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(excelFilePath) + "_" + suffix + Path.GetExtension(excelFilePath));
            File.Copy(excelFilePath, tempFileName, true);
            Console.WriteLine($"Создан временный файл: {tempFileName}");

            // Читаем данные из временного Excel файла
            IWorkbook workbook;
            using (FileStream file = new FileStream(tempFileName, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
            }
            var sheet = workbook.GetSheetAt(0);
            var rowCount = sheet.LastRowNum;

            // Сохраняем vCard на рабочий стол без суффикса
            string vCardFileName = Path.Combine(desktopPath, Path.GetFileNameWithoutExtension(excelFilePath) + ".vcf");
            Console.WriteLine($"Создаем файл vCard: {vCardFileName}");

            using (var vCardWriter = new StreamWriter(vCardFileName, false))
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    var currentRow = sheet.GetRow(row);
                    if (currentRow == null)
                    {
                        continue;
                    }

                    string location = currentRow.GetCell(1)?.ToString();
                    string name = currentRow.GetCell(3)?.ToString();
                    string position = currentRow.GetCell(4)?.ToString();
                    string email = currentRow.GetCell(5)?.ToString();
                    string phone = currentRow.GetCell(6)?.ToString();
                    string internalPhone = currentRow.GetCell(7)?.ToString();

                    // Пропускаем строки с пустыми обязательными полями
                    if (string.IsNullOrEmpty(name) || (string.IsNullOrEmpty(email) && string.IsNullOrEmpty(phone) && string.IsNullOrEmpty(internalPhone)))
                    {
                        continue;
                    }

                    // Очищаем номера телефонов от некорректных символов
                    phone = CleanPhoneNumber(phone);
                    internalPhone = CleanPhoneNumber(internalPhone);

                    // Проверяем корректность мобильного номера (учитываем только цифры)
                    if (Regex.Replace(phone, @"\D", "").Length != 11)
                    {
                        Console.WriteLine($"Некорректный мобильный номер: FN: {name}, Mobile: {phone}");
                    }

                    // Удаляем переносы строк из FN и заменяем множественные пробелы на один пробел
                    name = Regex.Replace(name.Replace("\n", " ").Replace("\r", " "), @"\s+", " ");

                    // Создаем vCard вручную
                    var vCard = new List<string>
                    {
                        "BEGIN:VCARD",
                        "VERSION:3.0",
                        $"FN:{name}",
                        $"ORG:{name}",
                        $"TITLE:{position}",
                        $"EMAIL;TYPE=INTERNET:{email}",
                        $"TEL;WORK;VOICE:{phone}",
                        $"TEL;CELL;VOICE:{internalPhone}",
                        $"ADR;WORK;PARCEL:{location};{name}",
                        "NOTE:Дополнительная информация о контакте",
                        "END:VCARD"
                    };

                    foreach (var line in vCard)
                    {
                        vCardWriter.WriteLine(line);
                    }
                }
            }

            Console.WriteLine($"Файл vCard успешно создан: {vCardFileName}");

            // Удаляем временный файл
            File.Delete(tempFileName);
            Console.WriteLine($"Временный файл удален: {tempFileName}");

            Console.WriteLine("Конвертация завершена. Нажмите любую клавишу для выхода.");
            Console.ReadKey();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }

    static string CleanPhoneNumber(string phoneNumber)
    {
        if (string.IsNullOrEmpty(phoneNumber))
        {
            return phoneNumber;
        }

        // Удаляем все символы, кроме цифр, +, -, и пробелов
        return new string(phoneNumber.Where(c => char.IsDigit(c) || c == '+' || c == '-' || c == ' ').ToArray());
    }

    static string GenerateRandomSuffix(int length)
    {
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        var random = new Random();
        return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
    }
}
