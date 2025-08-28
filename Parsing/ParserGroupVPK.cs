using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using convert_spravochnik_vpk_to_vcard;

namespace Converter.Parsing
{
    /// <summary>
    /// Парсер для кнопки «ВИЦ».
    /// Обходит все листы книги, читает шапку, нормализует строки под формат VPK
    /// (колонки Location, Name, Position, Email, Phone, InternalPhone),
    /// затем пишет временный .xlsx и отдаёт путь для VPKConverter.Convert(...).
    /// </summary>
    public sealed class ParserGroupVPK : IExcelParser
    {
        public string Name => "ВИЦ";

        // Варианты заголовков (расширяем при необходимости)
        private static readonly string[] H_Org        = { "организация" };
        private static readonly string[] H_Department = { "структурное подразделение/ департамент", "структурное подразделение", "департамент", "отдел", "служба", "подразделение" };
        private static readonly string[] H_Name       = { "фио", "ф.и.о", "фио сотрудника" };
        private static readonly string[] H_Position   = { "должность", "роль", "position", "title" };
        private static readonly string[] H_Email      = { "электронный адрес", "email", "e-mail", "почта" };
        private static readonly string[] H_CityCode   = { "код города", "городской код", "код" };
        private static readonly string[] H_CityNumber = { "городской номер", "городской", "телефон городской" };
        private static readonly string[] H_Mobile     = { "мобильный номер", "мобильный", "сотовый", "телефон мобильный" };
        private static readonly string[] H_Internal   = { "внутренний телефон", "внутренний", "доб", "доб." };
        private static readonly string[] H_ExtraNumOrMail = { "дополнительный номер/ e-mail", "дополнительный номер/e-mail", "дополнительный номер", "доп. номер", "доп номер", "дополнительный e-mail", "доп e-mail" };

        public string CreateVpkCompatibleWorkbook(string sourceExcelPath)
        {
            using var wb = WorkbookHelper.OpenWorkbook(sourceExcelPath);
            
            bool anyValid = false;
            var rows = new List<NormalizedContactRow>();

            // Проверяем все листы книги с жёсткой валидацией
            for (int s = 0; s < wb.NumberOfSheets; s++)
            {
                var sh = wb.GetSheetAt(s);
                var (ok, why) = StrictSchemaValidator.ValidateSheetFirstRow("ВИЦ", sh);
                
                if (!ok) 
                { 
                    Console.WriteLine($"[ВИЦ] Пропуск «{sh.SheetName}»: {why}"); 
                    continue; 
                }

                anyValid = true;
                Console.WriteLine($"[ВИЦ] Обрабатываем лист «{sh.SheetName}»");

                // Парсинг как в 18b487d: находим заголовки и колонки по имени
                var (headerRowIndex, headersRaw, headersCanon) = HeaderFinder.FindHeaderRow(sh);
                var cols = HeaderFinder.MapColumns("ВИЦ", headersRaw, headersCanon);

                Console.WriteLine($"[ВИЦ] Найдена строка заголовков: {headerRowIndex}, колонки: {string.Join(", ", cols.Keys)}");

                for (int r = headerRowIndex + 1; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    if (row == null) continue;

                    // Читаем поля через HeaderFinder
                    string name = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("fio", -1));
                    string position = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("title", -1));
                    string email = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("email", -1));
                    string mobile = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("mobile", -1));
                    string extra = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("extra", -1));
                    string internalNo = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("ext", -1));
                    string org = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("org", -1));
                    string dep = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("department", -1));

                    // Пропускаем строки без основных данных (телефонов)
                    if (string.IsNullOrWhiteSpace(mobile) && string.IsNullOrWhiteSpace(extra))
                        continue;

                    // Организация для ORG поля
                    string organization = CombineNonEmpty(CleanTextQuality(org), CleanTextQuality(dep), " / ");

                    // Если основного email нет, но в "Дополнительный номер/ e-mail" лежит почта — берём её.
                    if (string.IsNullOrWhiteSpace(email) && IsEmail(extra))
                        email = extra;

                    // Убираем пустые email
                    email = string.IsNullOrWhiteSpace(email) ? "" : email.Trim();

                    // Телефон: приоритет мобильному; если нет — берём из "Доп. номер", но только если это НЕ e-mail (бывает смешанное поле).
                    string phone = ChoosePhone(mobile, "", "", extra);

                    // Обработка добавочного номера
                    string cleanInternal = CleanSpaces(internalNo);
                    bool isExtension = !string.IsNullOrEmpty(cleanInternal) && 
                                      cleanInternal.All(char.IsDigit) && 
                                      cleanInternal.Length >= 3 && 
                                      cleanInternal.Length <= 5;

                    string finalInternalPhone = "";
                    
                    if (isExtension)
                    {
                        if (!string.IsNullOrEmpty(phone))
                        {
                            // Есть основной номер - добавляем extension
                            phone = $"{phone};ext={cleanInternal}";
                        }
                        else
                        {
                            // Нет основного номера - добавочный в InternalPhone для NOTE
                            finalInternalPhone = cleanInternal;
                        }
                    }
                    else if (!string.IsNullOrEmpty(cleanInternal))
                    {
                        // Это не добавочный (длинный номер) - пытаемся нормализовать
                        var normalized = RuPhone.NormalizeToE164RU(cleanInternal);
                        if (!string.IsNullOrEmpty(normalized))
                        {
                            finalInternalPhone = normalized;
                        }
                    }

                    // Пустые полностью строки не тащим
                    if (string.IsNullOrWhiteSpace(name) &&
                        string.IsNullOrWhiteSpace(email) &&
                        string.IsNullOrWhiteSpace(phone) &&
                        string.IsNullOrWhiteSpace(finalInternalPhone))
                        continue;

                    rows.Add(new NormalizedContactRow
                    {
                        Location      = organization,  // Теперь это правильная организация для ORG
                        Name          = name,
                        Position      = position,
                        Email         = email,
                        Phone         = phone,
                        InternalPhone = finalInternalPhone
                    });
                }
            }

            return VpkNormalizedWorkbookBuilder.BuildTempWorkbook(rows, "GroupVPK_");
        }

        // ---------- helpers ----------

        private static Dictionary<string, int> BuildHeaderMap(IRow headerRow)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
            {
                var raw = ExcelUtils.GetCellString(headerRow, i);
                if (string.IsNullOrWhiteSpace(raw)) continue;

                var key = NormalizeHeader(raw);
                if (!map.ContainsKey(key))
                    map[key] = i;
            }
            return map;
        }

        private static int FindIndex(Dictionary<string, int> map, IEnumerable<string> variants)
        {
            foreach (var v in variants)
            {
                var key = NormalizeHeader(v);
                if (map.TryGetValue(key, out var idx))
                    return idx;
            }
            return -1;
        }

        private static string NormalizeHeader(string s)
        {
            // нижний регистр, NBSP -> пробел, убрать переносы/точки, схлопнуть пробелы
            var x = (s ?? "")
                .Replace('\u00A0', ' ')
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Trim()
                .ToLowerInvariant();

            while (x.Contains("  ")) x = x.Replace("  ", " ");
            x = x.Replace(".", "");
            return x;
        }

        private static string Read(IRow row, int index) =>
            index >= 0 ? ExcelUtils.GetCellString(row, index) : "";

        private static string CombineNonEmpty(string a, string b, string sep)
        {
            a = a?.Trim() ?? "";
            b = b?.Trim() ?? "";
            if (a.Length > 0 && b.Length > 0) return a + sep + b;
            return a.Length > 0 ? a : b;
        }

        private static string CleanSpaces(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var x = s.Replace('\u00A0', ' ').Trim();
            while (x.Contains("  ")) x = x.Replace("  ", " ");
            return x;
        }

        /// <summary>
        /// Улучшенная очистка текста: убирает двойные пробелы, исправляет склейки
        /// </summary>
        private static string CleanTextQuality(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            
            var result = s.Replace('\u00A0', ' ')
                          .Replace('\r', ' ')
                          .Replace('\n', ' ')
                          .Trim();
            
            // Убираем множественные пробелы
            while (result.Contains("  ")) 
                result = result.Replace("  ", " ");
            
            // Исправляем типичные склейки (эвристика)
            result = System.Text.RegularExpressions.Regex.Replace(result, @"([а-яё])([А-ЯЁ])", "$1 $2");
            
            // Исправляем типичные опечатки в ВИЦ  
            result = result.Replace("Финансово- экономический", "Финансово-экономический");
            
            // Косметика: убираем лишние пробелы вокруг дефисов
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\s*-\s*", "-");
            
            return result;
        }

        private static bool IsEmail(string s)
        {
            s = s?.Trim() ?? "";
            if (s.Length == 0) return false;
            // очень простая евристика
            return s.Contains("@") && s.IndexOf('@') > 0 && s.IndexOf('@') < s.Length - 1;
        }

        private static string ChoosePhone(string mobile, string cityCode, string cityNumber, string extra)
        {
            // 1) мобильный
            var m = RuPhone.NormalizeToE164RU(CleanSpaces(mobile));
            if (!string.IsNullOrEmpty(m)) return m;

            // 2) городской (код + номер)
            var city = RuPhone.ComposeCityToE164RU(CleanSpaces(cityCode), CleanSpaces(cityNumber));
            if (!string.IsNullOrEmpty(city)) return city;

            // 3) "Доп. номер/ e-mail" — берём, только если это больше похоже на номер, и нормализуем к +7
            var ex = CleanSpaces(extra);
            if (!IsEmail(ex))
            {
                var exNorm = RuPhone.NormalizeToE164RU(ex);
                if (!string.IsNullOrEmpty(exNorm)) return exNorm;
            }

            return "";
        }
    }
}
