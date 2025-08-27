using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Converter.Parsing;

namespace convert_spravochnik_vpk_to_vcard
{
    public static class VPKConverter
    {
        /// <summary>
        /// Обрабатывает телефонную строку: нормализует основной номер к E.164,
        /// извлекает добавочные номера и возвращает структурированные данные
        /// </summary>
        static (string mainPhone, string extension, List<string> additionalPhones) ProcessPhoneString(string? phoneStr)
        {
            if (string.IsNullOrWhiteSpace(phoneStr))
                return ("", "", new List<string>());

            // Проверяем, есть ли уже готовый ext в строке (от парсеров)
            if (phoneStr.Contains(";ext="))
            {
                var extParts = phoneStr.Split(new[] { ";ext=" }, StringSplitOptions.None);
                if (extParts.Length == 2)
                {
                    var phone = extParts[0].Trim();
                    var ext = extParts[1].Trim();
                    return (phone, ext, new List<string>());
                }
            }

            var phones = new List<string>();
            var extensions = new List<string>();
            
            // Разделяем по пробелам и обрабатываем каждую часть
            var parts = phoneStr.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var part in parts)
            {
                var clean = new string(part.Where(c => char.IsDigit(c) || c == '+').ToArray());
                if (string.IsNullOrEmpty(clean)) continue;
                
                // Если это короткий номер (3-5 цифр) - это добавочный
                if (clean.Length >= 3 && clean.Length <= 5 && !clean.StartsWith("+") && !clean.StartsWith("8") && !clean.StartsWith("7"))
                {
                    extensions.Add(clean);
                }
                else
                {
                    // Пытаемся нормализовать как российский номер
                    var normalized = RuPhone.NormalizeToE164RU(clean);
                    if (!string.IsNullOrEmpty(normalized))
                    {
                        phones.Add(normalized);
                    }
                }
            }
            
            var mainPhone = phones.FirstOrDefault() ?? "";
            var extension = extensions.FirstOrDefault() ?? "";
            var additionalPhones = phones.Skip(1).ToList();
            
            return (mainPhone, extension, additionalPhones);
        }

        /// <summary>
        /// Определяет, является ли номер мобильным по коду оператора
        /// </summary>
        static bool IsMobileNumber(string e164Phone)
        {
            if (string.IsNullOrEmpty(e164Phone) || !e164Phone.StartsWith("+7"))
                return false;
                
            if (e164Phone.Length != 12) return false;
            
            var code = e164Phone.Substring(2, 3); // Первые 3 цифры после +7
            
            // Основные мобильные коды России
            return code.StartsWith("9") || // 9XX - мобильные
                   new[] { "800", "801", "802", "803", "804", "805", "806", "807", "808", "809" }.Contains(code);
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

            var contacts = new List<AppleVCardWriter.Contact>();

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

                // Обрабатываем телефоны с улучшенной нормализацией
                var (mainPhone, mainExtension, additionalPhones) = ProcessPhoneString(phone);
                var (internalPhoneNorm, internalExtension, additionalInternal) = ProcessPhoneString(internalPhone);

                // Удаляем переносы строк из FN и заменяем множественные пробелы на один пробел
                name = Regex.Replace(name.Replace("\n", " ").Replace("\r", " "), @"\s+", " ");

                // Определяем лучший рабочий номер
                string workPhone = "";
                string workExtension = "";
                string mobilePhone = "";
                
                // Если основной телефон мобильный - используем его как мобильный
                if (!string.IsNullOrEmpty(mainPhone) && IsMobileNumber(mainPhone))
                {
                    mobilePhone = mainPhone;
                    // Если есть внутренний и он не мобильный - как рабочий
                    if (!string.IsNullOrEmpty(internalPhoneNorm) && !IsMobileNumber(internalPhoneNorm))
                    {
                        workPhone = internalPhoneNorm;
                        workExtension = internalExtension;
                    }
                }
                else if (!string.IsNullOrEmpty(mainPhone))
                {
                    // Основной не мобильный - используем как рабочий
                    workPhone = mainPhone;
                    workExtension = mainExtension;
                    
                    // Если внутренний мобильный - используем как мобильный
                    if (!string.IsNullOrEmpty(internalPhoneNorm) && IsMobileNumber(internalPhoneNorm))
                    {
                        mobilePhone = internalPhoneNorm;
                    }
                }
                else if (!string.IsNullOrEmpty(internalPhoneNorm))
                {
                    // Есть только внутренний
                    if (IsMobileNumber(internalPhoneNorm))
                    {
                        mobilePhone = internalPhoneNorm;
                    }
                    else
                    {
                        workPhone = internalPhoneNorm;
                        workExtension = internalExtension;
                    }
                }

                // Формируем NOTE для добавочных без основного номера
                string note = "";
                if (!string.IsNullOrEmpty(internalPhoneNorm) && string.IsNullOrEmpty(workPhone) && string.IsNullOrEmpty(mobilePhone))
                {
                    if (internalPhoneNorm.All(char.IsDigit) && internalPhoneNorm.Length >= 3 && internalPhoneNorm.Length <= 5)
                    {
                        note = $"Добавочный номер: {internalPhoneNorm}";
                    }
                }

                contacts.Add(new AppleVCardWriter.Contact
                {
                    FullName = name,
                    OrgOrDept = location, // Используем location как организацию
                    Title = position,
                    Email = email,
                    MobileE164 = mobilePhone,
                    WorkE164 = workPhone,
                    Ext = workExtension,
                    Note = note
                });

                // Добавляем дополнительные контакты для дополнительных телефонов
                foreach (var additionalPhone in additionalPhones)
                {
                    contacts.Add(new AppleVCardWriter.Contact
                    {
                        FullName = $"{name} (доп.)",
                        OrgOrDept = location,
                        Title = position,
                        Email = "",
                        MobileE164 = IsMobileNumber(additionalPhone) ? additionalPhone : "",
                        WorkE164 = !IsMobileNumber(additionalPhone) ? additionalPhone : "",
                        Ext = "",
                        Note = "Дополнительный номер"
                    });
                }
            }

            // Записываем Apple-совместимый vCard файл
            AppleVCardWriter.WriteVCardFile(vCardFilePath, contacts);

            // Удаляем временный файл
            File.Delete(tempFileName);
        }

        private static string GenerateRandomSuffix(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
