using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using Converter.Parsing;

// =================== НОВЫЙ ФАЙЛ VPKCONVERTERFIXED ===================
// ЭТОТ ФАЙЛ СОЗДАН ЗАНОВО ДЛЯ ОБХОДА ПРОБЛЕМ С КЭШЕМ
// ВЕРСИЯ: ИСПРАВЛЕННАЯ С EXCELUTILS.OPEN
// ================================================================

namespace convert_spravochnik_vpk_to_vcard
{
    public static class VPKConverterFixed
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

        static string GenerateRandomSuffix(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static void Convert(string excelFilePath, string vCardFilePath)
        {
            // =================== СУПЕР КРИТИЧЕСКИЙ МАРКЕР ===================
            System.Diagnostics.Debug.WriteLine("====== VPKCONVERTERFIXED.CONVERT ЗАПУЩЕНА ======");
            System.Diagnostics.Debug.WriteLine($"[DEBUG] VPKConverterFixed.Convert ЗАПУЩЕН!");
            System.Diagnostics.Debug.WriteLine($"[DEBUG] Входной файл: {excelFilePath}");
            System.Diagnostics.Debug.WriteLine($"[DEBUG] Расширение: {Path.GetExtension(excelFilePath)}");
            
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

            // =================== КРИТИЧЕСКИЙ БЛОК: ИСПОЛЬЗУЕМ ExcelUtils.Open ===================
            System.Diagnostics.Debug.WriteLine($"[DEBUG] Перед ExcelUtils.Open: {tempFileName}, ext={Path.GetExtension(tempFileName)}");
            
            IWorkbook workbook;
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Вызываем ExcelUtils.Open...");
                workbook = ExcelUtils.Open(tempFileName);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ExcelUtils.Open успешно выполнен для {tempFileName}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] Ошибка при открытии файла: {tempFileName}");
                System.Diagnostics.Debug.WriteLine($"[ERROR] Тип: {ex.GetType().Name}");
                System.Diagnostics.Debug.WriteLine($"[ERROR] Ошибка: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[ERROR] StackTrace: {ex.StackTrace}");
                throw;
            }
            
            using (workbook)
            {
                var contacts = new List<AppleVCardWriter.Contact>();
                bool anyValid = false;

                // Проверяем все листы книги с железобетонной валидацией
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    var sheet = workbook.GetSheetAt(s);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Обрабатываем лист {s}: {sheet.SheetName}");
                    
                    // Простая проверка - есть ли данные
                    if (sheet.LastRowNum < 1)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Лист {sheet.SheetName} пуст, пропускаем");
                        continue;
                    }

                    anyValid = true;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Лист {sheet.SheetName} содержит {sheet.LastRowNum} строк, обрабатываем...");

                    // Определяем тип файла по заголовкам или имени листа
                    bool isOriginalVPK = false;
                    var headerRow = sheet.GetRow(0);
                    if (headerRow != null)
                    {
                        var header1 = headerRow.GetCell(1)?.ToString()?.Trim() ?? "";
                        // Если заголовок колонки 1 содержит "Местонахождение" - это оригинальный ВПК
                        // Если "Локация" - это нормализованный файл от других парсеров
                        isOriginalVPK = header1.Contains("Местонахождение");
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Заголовок колонки 1: '{header1}', isOriginalVPK: {isOriginalVPK}");
                    }

                    for (int row = 1; row <= sheet.LastRowNum; row++)
                    {
                        var dataRow = sheet.GetRow(row);
                        if (dataRow == null) continue;

                        try
                        {
                            if (isOriginalVPK)
                            {
                                // Оригинальный ВПК файл
                                var location = dataRow.GetCell(1)?.ToString()?.Trim() ?? "";   // Колонка 1 - Местонахождение (город)
                                var name = dataRow.GetCell(3)?.ToString()?.Trim() ?? "";       // Колонка 3 - ФИО
                                var position = dataRow.GetCell(4)?.ToString()?.Trim() ?? "";   // Колонка 4 - Должность
                                var email = dataRow.GetCell(5)?.ToString()?.Trim() ?? "";      // Колонка 5 - Email
                                var phone = dataRow.GetCell(6)?.ToString()?.Trim() ?? "";      // Колонка 6 - Телефон
                                var internalPhone = dataRow.GetCell(7)?.ToString()?.Trim() ?? ""; // Колонка 7 - Внутренний

                                if (string.IsNullOrWhiteSpace(name))
                                    continue;

                                // Очищаем переносы строк из имени и должности
                                name = System.Text.RegularExpressions.Regex.Replace(name.Replace("\n", " ").Replace("\r", " "), @"\s+", " ").Trim();
                                position = System.Text.RegularExpressions.Regex.Replace(position.Replace("\n", " ").Replace("\r", " "), @"\s+", " ").Trim();

                                var contact = new AppleVCardWriter.Contact
                                {
                                    FullName = name,
                                    OrgOrDept = "", // ВПК не имеет поля организации, оставляем пустым
                                    Title = position,
                                    Email = !string.IsNullOrWhiteSpace(email) ? email : "",
                                    City = location, // Местонахождение (город) в поле ADR
                                    Note = ""
                                };
                                
                                // Обрабатываем телефоны для ВПК (НЕ извлекаем extension из телефона!)
                                if (!string.IsNullOrWhiteSpace(phone))
                                {
                                    // Для ВПК просто нормализуем телефон без извлечения extension
                                    var cleanPhone = new string(phone.Where(c => char.IsDigit(c) || c == '+').ToArray());
                                    var normalized = RuPhone.NormalizeToE164RU(cleanPhone);
                                    
                                    if (!string.IsNullOrEmpty(normalized))
                                    {
                                        if (IsMobileNumber(normalized))
                                            contact.MobileE164 = normalized;
                                        else
                                            contact.WorkE164 = normalized;
                                    }
                                }
                                
                                if (!string.IsNullOrWhiteSpace(internalPhone))
                                {
                                    // Для ВПК внутренний номер - это extension, не Note
                                    // Если несколько номеров, объединяем их все
                                    var allExtensions = internalPhone.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(ext => ext.Trim())
                                        .Where(ext => !string.IsNullOrEmpty(ext))
                                        .Select(ext => new string(ext.Where(char.IsDigit).ToArray()))
                                        .Where(ext => !string.IsNullOrEmpty(ext))
                                        .ToList();
                                    
                                    if (allExtensions.Any())
                                    {
                                        // Если один номер - в Ext, если несколько - объединяем через точку с запятой
                                        contact.Ext = string.Join("; ", allExtensions);
                                    }
                                }
                                
                                contacts.Add(contact);
                            }
                            else
                            {
                                // Нормализованный файл от других парсеров (ВИЦ, ВЗК, ЗЗГТ)
                                var organization = dataRow.GetCell(1)?.ToString()?.Trim() ?? "";   // Колонка 1 - Организация
                                var name = dataRow.GetCell(3)?.ToString()?.Trim() ?? "";           // Колонка 3 - ФИО
                                var position = dataRow.GetCell(4)?.ToString()?.Trim() ?? "";       // Колонка 4 - Должность
                                var email = dataRow.GetCell(5)?.ToString()?.Trim() ?? "";          // Колонка 5 - Email
                                var phone = dataRow.GetCell(6)?.ToString()?.Trim() ?? "";          // Колонка 6 - Телефон
                                var internalPhone = dataRow.GetCell(7)?.ToString()?.Trim() ?? "";   // Колонка 7 - Внутренний

                                if (string.IsNullOrWhiteSpace(name))
                                    continue;

                                // Очищаем переносы строк из имени и должности
                                name = System.Text.RegularExpressions.Regex.Replace(name.Replace("\n", " ").Replace("\r", " "), @"\s+", " ").Trim();
                                position = System.Text.RegularExpressions.Regex.Replace(position.Replace("\n", " ").Replace("\r", " "), @"\s+", " ").Trim();

                                var contact = new AppleVCardWriter.Contact
                                {
                                    FullName = name,
                                    OrgOrDept = organization, // Для других парсеров это действительно организация
                                    Title = position,
                                    Email = !string.IsNullOrWhiteSpace(email) ? email : "",
                                    Note = ""
                                };
                                
                                // Обрабатываем телефоны для других парсеров
                                if (!string.IsNullOrWhiteSpace(phone))
                                {
                                    var (mainPhone, extension, additionalPhones) = ProcessPhoneString(phone);
                                    if (!string.IsNullOrEmpty(mainPhone))
                                    {
                                        if (IsMobileNumber(mainPhone))
                                            contact.MobileE164 = mainPhone;
                                        else
                                            contact.WorkE164 = mainPhone;
                                    }
                                    
                                    if (!string.IsNullOrEmpty(extension))
                                    {
                                        contact.Ext = extension;
                                    }
                                }
                                
                                if (!string.IsNullOrWhiteSpace(internalPhone))
                                {
                                    // Для других парсеров внутренний номер тоже extension
                                    // Если несколько номеров, объединяем их все
                                    var allExtensions = internalPhone.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(ext => ext.Trim())
                                        .Where(ext => !string.IsNullOrEmpty(ext))
                                        .Select(ext => new string(ext.Where(char.IsDigit).ToArray()))
                                        .Where(ext => !string.IsNullOrEmpty(ext))
                                        .ToList();
                                    
                                    if (allExtensions.Any() && string.IsNullOrEmpty(contact.Ext))
                                    {
                                        // Если один номер - в Ext, если несколько - объединяем через точку с запятой
                                        contact.Ext = string.Join("; ", allExtensions);
                                    }
                                }
                                
                                contacts.Add(contact);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ERROR] Ошибка обработки строки {row}: {ex.Message}");
                        }
                    }
                }

                if (!anyValid)
                {
                    throw new InvalidOperationException("Не найдено ни одного валидного листа для обработки");
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] Всего обработано контактов: {contacts.Count}");

                // Записываем vCard файл
                AppleVCardWriter.WriteVCardFile(vCardFilePath, contacts);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] vCard файл записан: {vCardFilePath}");

                // Очищаем временный файл
                try
                {
                    File.Delete(tempFileName);
                }
                catch { }
            }
        }
    }
}
