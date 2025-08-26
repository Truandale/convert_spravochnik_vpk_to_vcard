using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Converter.Parsing
{
    /// <summary>
    /// Парсер для кнопки «Группа ВПК».
    /// Обходит все листы книги, читает шапку, нормализует строки под формат VPK
    /// (колонки Location, Name, Position, Email, Phone, InternalPhone),
    /// затем пишет временный .xlsx и отдаёт путь для VPKConverter.Convert(...).
    /// </summary>
    public sealed class ParserGroupVPK : IExcelParser
    {
        public string Name => "Группа ВПК";

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
            using var wb = ExcelUtils.Open(sourceExcelPath);

            var rows = new List<NormalizedContactRow>();

            for (int s = 0; s < wb.NumberOfSheets; s++)
            {
                var sh = wb.GetSheetAt(s);
                if (sh == null) continue;

                // Заголовок — первая строка листа
                var header = sh.GetRow(sh.FirstRowNum);
                if (header == null) continue;

                var headers = BuildHeaderMap(header);

                int idxOrg        = FindIndex(headers, H_Org);
                int idxDep        = FindIndex(headers, H_Department);
                int idxName       = FindIndex(headers, H_Name);
                int idxPos        = FindIndex(headers, H_Position);
                int idxEmail      = FindIndex(headers, H_Email);
                int idxCityCode   = FindIndex(headers, H_CityCode);
                int idxCityNumber = FindIndex(headers, H_CityNumber);
                int idxMobile     = FindIndex(headers, H_Mobile);
                int idxInternal   = FindIndex(headers, H_Internal);
                int idxExtra      = FindIndex(headers, H_ExtraNumOrMail);

                for (int r = sh.FirstRowNum + 1; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    if (row == null) continue;

                    string org        = Read(row, idxOrg);
                    string dep        = Read(row, idxDep);
                    string name       = Read(row, idxName);
                    string position   = Read(row, idxPos);
                    string email      = Read(row, idxEmail);
                    string cityCode   = Read(row, idxCityCode);
                    string cityNumber = Read(row, idxCityNumber);
                    string mobile     = Read(row, idxMobile);
                    string internalNo = Read(row, idxInternal);
                    string extra      = Read(row, idxExtra);

                    // Локация: "Организация / Подразделение"
                    string location = CombineNonEmpty(org, dep, " / ");

                    // Если основного email нет, но в "Дополнительный номер/ e-mail" лежит почта — берём её.
                    if (string.IsNullOrWhiteSpace(email) && IsEmail(extra))
                        email = extra;

                    // Телефон: приоритет мобильному; если нет — городской (код + номер);
                    // если и этого нет — берём из "Доп. номер", но только если это НЕ e-mail (бывает смешанное поле).
                    string phone = ChoosePhone(mobile, cityCode, cityNumber, extra);

                    // Пустые полностью строки не тащим
                    if (string.IsNullOrWhiteSpace(name) &&
                        string.IsNullOrWhiteSpace(email) &&
                        string.IsNullOrWhiteSpace(phone) &&
                        string.IsNullOrWhiteSpace(internalNo))
                        continue;

                    rows.Add(new NormalizedContactRow
                    {
                        Location      = location,
                        Name          = name,
                        Position      = position,
                        Email         = email,
                        Phone         = phone,
                        InternalPhone = internalNo
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
