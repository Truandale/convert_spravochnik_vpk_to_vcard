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

            // Мобильный телефон
            if (!string.IsNullOrWhiteSpace(c.MobileE164))
            {
                w.WriteLine($"TEL;TYPE=CELL,VOICE:{c.MobileE164}");
            }

            // Рабочий телефон с добавочным
            if (!string.IsNullOrWhiteSpace(c.WorkE164))
            {
                if (!string.IsNullOrWhiteSpace(c.Ext))
                {
                    // RFC 3966 формат для добавочного номера (лучше для парсинга)
                    w.WriteLine($"TEL;TYPE=WORK,VOICE;VALUE=uri:tel:{c.WorkE164};ext={c.Ext}");
                }
                else
                {
                    w.WriteLine($"TEL;TYPE=WORK,VOICE:{c.WorkE164}");
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
        /// Создает Apple-совместимый vCard файл
        /// </summary>
        public static void WriteVCardFile(string filePath, IEnumerable<Contact> contacts)
        {
            using var w = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            w.NewLine = "\r\n"; // CRLF для максимальной совместимости с Apple/iOS
            
            foreach (var contact in contacts)
            {
                WriteContact(w, contact);
            }
        }
    }
}
