using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace convert_spravochnik_vpk_to_vcard
{
    public static class StrictSchemaValidator
    {
        private static string Canon(string? s)
        {
            s ??= "";
            s = s.ToLowerInvariant().Replace('ё','е')
                 .Replace("\r"," ").Replace("\n"," ")
                 .Replace("e-mail","email").Replace("e mail","email");
            s = Regex.Replace(s, @"[^a-zа-я0-9]+", " ");
            s = Regex.Replace(s, @"\s+", " ").Trim();
            return s;
        }

        // Эталонные последовательности (подряд, на первой строке).
        // ВАЖНО: строки НЕ должны быть пустыми – иначе совпадёт «всё со всем».
        // Сигнатуры основаны на РЕАЛЬНЫХ заголовках из Excel файлов!
        private static readonly Dictionary<string, string[]> Sig = new()
        {
            ["ВПК"]  = new[]{ Canon("ФИО"), Canon("Должность"), Canon("E-mail"), Canon("Контактный телефон"), Canon("Внутр. номер телефона") },
            ["ВЗК"]  = new[]{ Canon("Код города"), Canon("Городской номер"), Canon("Мобильный номер"), Canon("Внутренний телефон") },
            ["ВИЦ"]  = new[]{ Canon("Мобильный номер"), Canon("Дополнительный номер/ e-mail"), Canon("Внутренний телефон") },
            ["ЗЗГТ"] = new[]{ Canon("Код города"), Canon("Городской номер"), Canon("Мобильный номер"), Canon("Внутренний телефон") },
        };

        // (опционально) различаем листы с одинаковой шапкой по имени листа
        private static readonly Dictionary<string, Regex> SheetNameGuard = new()
        {
            ["ВПК"]  = new(@"впк", RegexOptions.IgnoreCase),
            ["ВЗК"]  = new(@"взк|завод|корпусов", RegexOptions.IgnoreCase),
            ["ВИЦ"]  = new(@"виц", RegexOptions.IgnoreCase),
            ["ЗЗГТ"] = new(@"ззгт", RegexOptions.IgnoreCase),
        };

        public static (bool ok, int start, string why) ValidateFirstRowExact(string button, ISheet sh)
        {
            if (!Sig.TryGetValue(button, out var sig) || sig.Length == 0)
                return (false, -1, $"Неизвестная кнопка «{button}» (нет сигнатуры).");

            var row0 = sh.GetRow(0);
            if (row0 == null) return (false, -1, $"Лист «{sh.SheetName}»: первая строка пустая.");

            var hdr = new List<string>();
            for (int c = 0; c < row0.LastCellNum; c++)
                hdr.Add(Canon(row0.GetCell(c)?.ToString() ?? ""));

            // обрезаем пустой хвост
            while (hdr.Count > 0 && hdr[^1].Length == 0) hdr.RemoveAt(hdr.Count - 1);

            // ЖЁСТКО: без contains, только равенство, только подряд
            for (int i = 0; i <= hdr.Count - sig.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < sig.Length; j++)
                {
                    if (hdr[i + j] != sig[j]) { match = false; break; }
                    if (sig[j].Length == 0)   { match = false; break; } // safety
                }
                if (match)
                {
                    // доп. якорь по имени листа (снимите, если не нужно)
                    if (SheetNameGuard.TryGetValue(button, out var re) && !re.IsMatch(sh.SheetName))
                        return (false, -1, $"Лист «{sh.SheetName}»: имя листа не соответствует кнопке «{button}».");
                    return (true, i, "");
                }
            }

            var got  = string.Join(" | ", hdr);
            var need = string.Join(" | ", sig);
            return (false, -1, $"Лист «{sh.SheetName}»: заголовки не совпадают.\nОжидалось: [{need}]\nПолучено:  [{got}]");
        }

        // Индексы полей от начала совпавшего блока (чтобы парсеры не «искали» сами)
        // Основано на РЕАЛЬНЫХ позициях в Excel файлах после валидации
        public static Dictionary<string,int> GetHeaderIndexes(string button, int start) =>
            button switch
            {
                "ВПК"  => new() { ["fio"]=start+0, ["title"]=start+1, ["email"]=start+2, ["cell"]=start+3, ["ext"]=start+4 },
                "ВЗК"  => new() { ["code"]=start+0, ["city"]=start+1, ["mobile"]=start+2, ["ext"]=start+3 },
                "ВИЦ"  => new() { ["mobile"]=start+0, ["extra"]=start+1, ["ext"]=start+2 },
                "ЗЗГТ" => new() { ["code"]=start+0, ["city"]=start+1, ["mobile"]=start+2, ["ext"]=start+3 },
                _      => throw new InvalidOperationException($"Неизвестная кнопка «{button}».")
            };
    }
}