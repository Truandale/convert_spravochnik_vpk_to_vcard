using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;

namespace Converter.Parsing
{
    /// <summary>
    /// Строгий валидатор схем - каждая кнопка знает точную последовательность колонок
    /// </summary>
    public static class StrictSchemaValidator
    {
        // Канонизация: убираем регистр/знаки/переносы, склеиваем дефисы, e-mail → email
        private static string Canon(string? s)
        {
            s ??= "";
            s = s.ToLowerInvariant()
                 .Replace('ё','е')
                 .Replace("\r"," ").Replace("\n"," ")
                 .Replace("e-mail","email").Replace("e mail","email");
            s = Regex.Replace(s, @"[^a-zа-я0-9]+", " "); // всё, кроме букв/цифр → пробел
            s = Regex.Replace(s, @"\s+", " ").Trim();
            return s;
        }

        // Сигнатуры (строгая последовательность ключевых колонок — как в твоих Excel)
        // Сигнатуры пишем уже в канон-форме, чтобы сравнение было стабильным
        private static readonly Dictionary<string, string[]> Signatures = new()
        {
            // ВПК.xlsx: … | ФИО | Должность | E-mail: | Контактный телефон | Внутр. номер телефона
            ["ВПК"] = new[]
            {
                Canon("ФИО"),
                Canon("Должность"),
                Canon("E-mail"),
                Canon("Контактный телефон"),
                Canon("Внутр. номер телефона"),
            },

            // ВЗК.xlsx: Организация | Структурное подразделение | ФИО | Должность | Электронный адрес | Код города | Городской номер | Мобильный номер | Внутренний телефон
            ["ВЗК"] = new[]
            {
                Canon("Организация"),
                Canon("Структурное подразделение"),
                Canon("ФИО"),
                Canon("Должность"),
                Canon("Электронный адрес"),
                Canon("Код города"),
                Canon("Городской номер"),
                Canon("Мобильный номер"),
                Canon("Внутренний телефон"),
            },

            // ВИЦ.xlsx: ФИО | Должность | E-mail | Контактный телефон | Внутр. номер | Подразделение
            ["ВИЦ"] = new[]
            {
                Canon("Организация"),
                Canon("Структурное подразделение/ департамент"),
                Canon("ФИО"),
                Canon("Должность"),
                Canon("Электронный адрес"),
                Canon("Код города"),
                Canon("Городской номер"),
                Canon("Мобильный номер"),
                Canon("Дополнительный номер/ e-mail"),
                Canon("Внутренний телефон"),
            },

            // ЗЗГТ.xlsx: Организация | Структурное подразделение | ФИО | Должность | Электронный адрес | Код города | Городской номер | Мобильный номер | Внутренний телефон
            ["ЗЗГТ"] = new[]
            {
                Canon("Организация"),
                Canon("Структурное подразделение/ департамент"),
                Canon("ФИО"),
                Canon("Должность"),
                Canon("Электронный адрес"),
                Canon("Код города"),
                Canon("Городской номер"),
                Canon("Мобильный номер"),
                Canon("Внутренний телефон"),
            },
        };

        // Ищем строку заголовков среди первых 6 строк (часто бывает шапка/титул)
        private static (int rowIndex, List<string> headersCanon, List<string> headersRaw) ExtractHeaders(ISheet sh)
        {
            for (int r = 0; r < Math.Min(6, sh.LastRowNum + 1); r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;

                var raw = new List<string>();
                var canon = new List<string>();
                for (int c = 0; c < row.LastCellNum; c++)
                {
                    var v = row.GetCell(c)?.ToString() ?? "";
                    raw.Add(v);
                    canon.Add(Canon(v));
                }

                // эвристика: наличие «фио» в этой строке
                if (canon.Any(h => h == Canon("ФИО") || h.Contains(Canon("ФИО"))))
                    return (r, canon, raw);
            }

            // fallback: первая строка
            var r0 = sh.GetRow(0);
            var raw0 = new List<string>(); var canon0 = new List<string>();
            if (r0 != null)
            {
                for (int c = 0; c < r0.LastCellNum; c++)
                {
                    var v = r0.GetCell(c)?.ToString() ?? "";
                    raw0.Add(v); canon0.Add(Canon(v));
                }
            }
            return (0, canon0, raw0);
        }

        // Строго: сигнатура должна встретиться как КОНТАГИУЗНЫЙ блок (подряд) в строке заголовков.
        // Допускаем лишние столбцы слева/справа, но НЕ между элементами сигнатуры.
        private static bool ContainsContiguousSlice(IReadOnlyList<string> row, IReadOnlyList<string> signature, out int startIndex)
        {
            startIndex = -1;
            if (row.Count == 0 || signature.Count == 0) return false;

            for (int i = 0; i <= row.Count - signature.Count; i++)
            {
                bool ok = true;
                for (int j = 0; j < signature.Count; j++)
                {
                    if (row[i + j] != signature[j]) { ok = false; break; }
                }
                if (ok) { startIndex = i; return true; }
            }
            return false;
        }

        public sealed record ValidationResult(bool IsValid, string Reason, int HeaderRowIndex, int SliceStart, int SliceLength);

        public static ValidationResult ValidateStrict(string button, ISheet sheet)
        {
            if (!Signatures.TryGetValue(button, out var signature))
                return new(false, $"Неизвестная кнопка «{button}» (нет сигнатуры).", -1, -1, 0);

            var (hdrRow, canonHeaders, rawHeaders) = ExtractHeaders(sheet);
            // подрезаем хвостовые пустые колонки
            while (canonHeaders.Count > 0 && string.IsNullOrWhiteSpace(canonHeaders[^1])) canonHeaders.RemoveAt(canonHeaders.Count - 1);

            if (canonHeaders.Count == 0)
                return new(false, $"Не нашёл заголовков в листе.", hdrRow, -1, 0);

            if (!ContainsContiguousSlice(canonHeaders, signature, out var start))
            {
                // соберём подсказку: покажем канон-заголовки листа
                var sample = string.Join(" | ", canonHeaders);
                var expected = string.Join(" | ", signature);
                var why = $"Заголовки не соответствуют сигнатуре.\n" +
                          $"Ожидалось подряд: [{expected}]\n" +
                          $"В файле:          [{sample}]";
                return new(false, why, hdrRow, -1, 0);
            }

            return new(true, "", hdrRow, start, signature.Length);
        }

        /// <summary>
        /// Валидация для ВИЦ
        /// </summary>
        public static ValidationResult ValidateVIC(ISheet sheet)
        {
            return ValidateStrict("ВИЦ", sheet);
        }

        /// <summary>
        /// Валидация для ВЗК
        /// </summary>
        public static ValidationResult ValidateVZK(ISheet sheet)
        {
            return ValidateStrict("ВЗК", sheet);
        }

        /// <summary>
        /// Валидация для ЗЗГТ
        /// </summary>
        public static ValidationResult ValidateZZGT(ISheet sheet)
        {
            return ValidateStrict("ЗЗГТ", sheet);
        }

        /// <summary>
        /// Валидация для ВПК
        /// </summary>
        public static ValidationResult ValidateVPK(ISheet sheet)
        {
            return ValidateStrict("ВПК", sheet);
        }
    }
}
