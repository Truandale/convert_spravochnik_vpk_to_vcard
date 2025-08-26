using System;
using System.Linq;

namespace Converter.Parsing
{
    /// <summary>Нормализация телефонных номеров под российский формат E.164 (+7).</summary>
    public static class RuPhone
    {
        /// <summary>
        /// Нормализует произвольную строку с номером к +7XXXXXXXXXX, если это российский 10/11-значный номер.
        /// Поддержка:
        ///  - 8XXXXXXXXXX -> +7XXXXXXXXXX
        ///  - 7XXXXXXXXXX -> +7XXXXXXXXXX
        ///  - XXXXXXXXXX  -> +7XXXXXXXXXX (если 10 цифр)
        ///  - +7XXXXXXXXXX (оставляем как есть)
        /// Если не удалось однозначно привести — вернёт пустую строку.
        /// </summary>
        public static string NormalizeToE164RU(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            var digits = new string(raw.Where(char.IsDigit).ToArray());

            if (raw.Trim().StartsWith("+"))
            {
                // уже международный; если не +7 — оставим как есть (может быть внешний номер)
                if (raw.Trim().StartsWith("+7")) return "+7" + TakeLast(digits, 10);
                return raw.Trim(); // не RU, не трогаем
            }

            if (digits.Length == 11 && (digits[0] == '7' || digits[0] == '8'))
                return "+7" + digits.Substring(1, 10);

            if (digits.Length == 10)
                return "+7" + digits;

            // Частые кейсы: мобильный на 10, городской код+номер на 10.
            // Если длина не 10/11 — считаем, что не можем привести надёжно.
            return "";
        }

        /// <summary>
        /// Склеивает "код города" + "городской номер" и нормализует к +7XXXXXXXXXX.
        /// Пример: code="83161", number="2-14-01" -> "+78316121401".
        /// </summary>
        public static string ComposeCityToE164RU(string? cityCode, string? cityNumber)
        {
            var cc = new string((cityCode ?? "").Where(char.IsDigit).ToArray());
            var cn = new string((cityNumber ?? "").Where(char.IsDigit).ToArray());
            var joined = cc + cn;

            if (joined.Length == 11 && (joined[0] == '7' || joined[0] == '8'))
                return "+7" + joined.Substring(1, 10);

            if (joined.Length == 10)
                return "+7" + joined;

            // Некоторые код+номер бывают диковатых длин, но для надёжного набора с мобильного нужен НННННННННН (10).
            return "";
        }

        private static string TakeLast(string s, int n) =>
            s.Length <= n ? s : s.Substring(s.Length - n, n);
    }
}
