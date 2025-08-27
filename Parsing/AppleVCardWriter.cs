using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Converter.Parsing
{
    /// <summary>
    /// Apple-совместимый генератор vCard 3.0 файлов
    /// Обеспечивает максимальную совместимость с iPhone/iOS Contacts
    /// </summary>
    public static class AppleVCardWriter
    {
        /// <summary>
        /// Модель контакта для Apple-совместимого vCard
        /// </summary>
        public class Contact
        {
            public string FullName { get; set; } = "";
            public string OrgOrDept { get; set; } = "";
            public string Title { get; set; } = "";
            public string Email { get; set; } = "";
            public string MobileE164 { get; set; } = "";
            public string WorkE164 { get; set; } = "";
            public string Ext { get; set; } = "";
            public string Note { get; set; } = "";
        }

        /// <summary>
        /// Экранирует специальные символы в значениях vCard согласно vCard 3.0
        /// </summary>
        public static string Esc(string? s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            
            // Убираем "сырые" переводы строк из Excel и хвосты
            var t = s.Replace("\r\n", "\n").Replace("\r", "\n").Trim();
            
            // Экранируем спецсимволы по vCard 3.0
            t = t.Replace("\\", "\\\\")   // backslash
                 .Replace(";", "\\;")     // semicolon  
                 .Replace(",", "\\,")     // comma
                 .Replace("\n", "\\n");   // newline
            
            return t;
        }

        /// <summary>
        /// Простейший разбор ФИО на компоненты для поля N
        /// </summary>
        public static (string last, string first, string middle) SplitFio(string? fn)
        {
            var parts = (fn ?? "").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length switch
            {
                >= 3 => (parts[0], parts[1], string.Join(" ", parts.Skip(2))),
                2 => (parts[0], parts[1], ""),
                1 => (parts[0], "", ""),
                _ => ("", "", "")
            };
        }

        /// <summary>
        /// Разбивает строку email на отдельные адреса
        /// </summary>
        public static IEnumerable<string> SplitEmails(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) yield break;
            
            foreach (var e in s.Split(new[] { ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = e.Trim();
                if (trimmed.Contains("@"))
                    yield return trimmed;
            }
        }

        /// <summary>
        /// Записывает один контакт в Apple-совместимом формате vCard 3.0
        /// </summary>
        public static void WriteContact(StreamWriter w, Contact c)
        {
            w.WriteLine("BEGIN:VCARD");
            w.WriteLine("VERSION:3.0");
            
            // FN - как на визитке
            w.WriteLine($"FN:{Esc(c.FullName)}");
            
            // N - структурированное имя (Фамилия;Имя;Отчество;;)
            var (lastName, firstName, middleName) = SplitFio(c.FullName);
            w.WriteLine($"N:{Esc(lastName)};{Esc(firstName)};{Esc(middleName)};;");

            // ORG - только если есть настоящая организация (не дублируем FN)
            if (!string.IsNullOrWhiteSpace(c.OrgOrDept))
            {
                w.WriteLine($"ORG:{Esc(c.OrgOrDept)}");
            }

            // TITLE - должность
            if (!string.IsNullOrWhiteSpace(c.Title))
            {
                w.WriteLine($"TITLE:{Esc(c.Title)}");
            }

            // EMAIL - разбиваем на отдельные строки
            foreach (var emailAddr in SplitEmails(c.Email))
            {
                w.WriteLine($"EMAIL;TYPE=INTERNET:{Esc(emailAddr)}");
            }

            // Мобильный телефон (умное определение типа + жёсткая нормализация)
            if (!string.IsNullOrWhiteSpace(c.MobileE164))
            {
                var strictPhone = RuPhone.StrictE164RU(c.MobileE164);
                if (!string.IsNullOrEmpty(strictPhone))
                {
                    // Если номер начинается с +79xx - это действительно мобильный (CELL)
                    // Иначе это городской номер, помеченный как мобильный (WORK)
                    var phoneType = strictPhone.StartsWith("+79") ? "CELL" : "WORK";
                    w.WriteLine($"TEL;TYPE={phoneType},VOICE:{strictPhone}");
                }
            }

            // Рабочий телефон с добавочным (жёсткая нормализация)
            if (!string.IsNullOrWhiteSpace(c.WorkE164))
            {
                var strictWork = RuPhone.StrictE164RU(c.WorkE164);
                if (!string.IsNullOrEmpty(strictWork))
                {
                    if (!string.IsNullOrWhiteSpace(c.Ext))
                    {
                        // RFC 3966 формат для добавочного номера (лучше для парсинга)
                        w.WriteLine($"TEL;TYPE=WORK,VOICE;VALUE=uri:tel:{strictWork};ext={c.Ext}");
                    }
                    else
                    {
                        w.WriteLine($"TEL;TYPE=WORK,VOICE:{strictWork}");
                    }
                }
            }
            else if (!string.IsNullOrWhiteSpace(c.Ext))
            {
                // Есть только добавочный номер без основного - выносим в NOTE
                var extNote = $"Внутренний номер: {c.Ext}";
                if (!string.IsNullOrWhiteSpace(c.Note))
                {
                    c.Note = $"{c.Note}; {extNote}";
                }
                else
                {
                    c.Note = extNote;
                }
            }

            // NOTE для дополнительной информации
            if (!string.IsNullOrWhiteSpace(c.Note))
            {
                w.WriteLine($"NOTE:{Esc(c.Note)}");
            }

            w.WriteLine("END:VCARD");
        }

        /// <summary>
        /// Создает Apple-совместимый vCard файл с объединением дублей
        /// </summary>
        public static void WriteVCardFile(string filePath, IEnumerable<Contact> contacts)
        {
            // Объединяем дубли по ФИО перед записью
            var mergedContacts = MergeDuplicates(contacts);
            
            using var w = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            w.NewLine = "\r\n"; // CRLF для максимальной совместимости с Apple/iOS
            
            foreach (var contact in mergedContacts)
            {
                WriteContact(w, contact);
            }
        }
        
        /// <summary>
        /// Объединяет дубли контактов по одинаковому ФИО
        /// Склеивает email'ы, заметки; оставляет первый телефон
        /// </summary>
        private static List<Contact> MergeDuplicates(IEnumerable<Contact> contacts)
        {
            var grouped = contacts
                .Where(c => !string.IsNullOrWhiteSpace(c.FullName))
                .GroupBy(c => c.FullName.Trim().ToLowerInvariant())
                .ToList();
                
            var result = new List<Contact>();
            
            foreach (var group in grouped)
            {
                if (group.Count() == 1)
                {
                    result.Add(group.First());
                }
                else
                {
                    // Объединяем дубли
                    var first = group.First();
                    var merged = new Contact
                    {
                        FullName = first.FullName,
                        OrgOrDept = first.OrgOrDept,
                        Title = CombineNonEmpty(group.Select(c => c.Title), " / "),
                        Email = CombineNonEmpty(group.Select(c => c.Email).Where(e => !string.IsNullOrWhiteSpace(e)), ", "),
                        MobileE164 = first.MobileE164, // Первый телефон
                        WorkE164 = first.WorkE164,
                        Ext = first.Ext,
                        Note = CombineNotesNicely(group.Select(c => c.Note))
                    };
                    result.Add(merged);
                }
            }
            
            return result;
        }
        
        private static string CombineNonEmpty(IEnumerable<string> values, string separator)
        {
            var nonEmpty = values.Where(v => !string.IsNullOrWhiteSpace(v)).Distinct();
            return string.Join(separator, nonEmpty);
        }
        
        /// <summary>
        /// Умное объединение заметок: выделяет добавочные номера и объединяет их красиво
        /// </summary>
        private static string CombineNotesNicely(IEnumerable<string> notes)
        {
            var notesList = notes.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            if (!notesList.Any()) return "";
            
            var extensions = new List<string>();
            var otherNotes = new List<string>();
            
            foreach (var note in notesList)
            {
                // Ищем добавочные номера
                if (note.Contains("Внутренний номер:"))
                {
                    var extMatch = System.Text.RegularExpressions.Regex.Match(note, @"Внутренний номер:\s*(\d+)");
                    if (extMatch.Success)
                    {
                        extensions.Add(extMatch.Groups[1].Value);
                    }
                    else
                    {
                        otherNotes.Add(note); // Не смогли извлечь номер, оставляем как есть
                    }
                }
                else
                {
                    otherNotes.Add(note);
                }
            }
            
            var result = new List<string>();
            
            // Объединяем добавочные в одну строку
            if (extensions.Any())
            {
                var uniqueExts = extensions.Distinct().ToList();
                if (uniqueExts.Count == 1)
                    result.Add($"Внутренний номер: {uniqueExts[0]}");
                else
                    result.Add($"Внутренние номера: {string.Join(", ", uniqueExts)}");
            }
            
            // Добавляем остальные заметки
            result.AddRange(otherNotes.Distinct());
            
            return string.Join(" | ", result);
        }
    }
}
