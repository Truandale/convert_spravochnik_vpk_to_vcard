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
            
            // Сначала убираем все символы кроме цифр и плюса
            var cleaned = new string(raw.Where(c => char.IsDigit(c) || c == '+').ToArray());
            var digits = new string(cleaned.Where(char.IsDigit).ToArray());

            if (cleaned.StartsWith("+"))
            {
                // уже международный; если не +7 — оставим как есть (может быть внешний номер)
                if (cleaned.StartsWith("+7")) return "+7" + TakeLast(digits, 10);
                return cleaned; // не RU, не трогаем
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
        /// Добавлена валидация длины для ЗЗГТ: отклоняем странные номера.
        /// </summary>
        public static string ComposeCityToE164RU(string? cityCode, string? cityNumber)
        {
            var cc = new string((cityCode ?? "").Where(char.IsDigit).ToArray());
            var cn = new string((cityNumber ?? "").Where(char.IsDigit).ToArray());
            var joined = cc + cn;

            if (joined.Length == 11 && (joined[0] == '7' || joined[0] == '8'))
            {
                var result = "+7" + joined.Substring(1, 10);
                // Дополнительная проверка: результат должен быть ровно 12 символов (+7 + 10 цифр)
                if (result.Length == 12) return result;
            }

            if (joined.Length == 10)
            {
                var result = "+7" + joined;
                // Проверка: итоговая длина должна быть 12 символов
                if (result.Length == 12) return result;
            }

            // Валидация для российских номеров: после +7 должно быть ровно 10 цифр
            // Странные комбинации типа +7587... или +7289... отклоняем
            return "";
        }

        private static string TakeLast(string s, int n) =>
            s.Length <= n ? s : s.Substring(s.Length - n, n);

        /// <summary>
        /// Жёсткая нормализация для финального vCard: убирает ВСЁ кроме +7 и 10 цифр.
        /// Результат строго +7XXXXXXXXXX или пустая строка.
        /// Использовать перед записью TEL для гарантии качества.
        /// </summary>
        public static string StrictE164RU(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            
            // Убираем всё кроме цифр и плюса
            var cleaned = new string(raw.Where(c => char.IsDigit(c) || c == '+').ToArray());
            
            // Проверяем финальный паттерн: +7 + ровно 10 цифр
            if (System.Text.RegularExpressions.Regex.IsMatch(cleaned, @"^\+7\d{10}$"))
                return cleaned;
                
            return "";
        }
    }
}
