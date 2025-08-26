using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Converter.Parsing
{
    /// <summary>
    /// Парсер для книги «ЗЗГТ».
    /// Берёт нужные колонки по названиям шапок и приводит к формату,
    /// который уже умеет конвертер кнопки ВПК.
    /// </summary>
    public sealed class ParserZZGT : IExcelParser
    {
        public string Name => "ЗЗГТ";

        // Варианты заголовков (на всякий случай — с запасом)
        private static readonly string[] H_Org        = { "организация" };
        private static readonly string[] H_Department = { "структурное подразделение/ департамент", "структурное подразделение", "департамент", "отдел", "служба", "подразделение" };
        private static readonly string[] H_Name       = { "фио", "ф.и.о" };
        private static readonly string[] H_Position   = { "должность", "роль", "position", "title" };
        private static readonly string[] H_Email      = { "электронный адрес", "email", "e-mail", "почта" };
        private static readonly string[] H_CityCode   = { "код города", "городской код", "код" };
        private static readonly string[] H_CityNumber = { "городской номер", "городской", "телефон городской" };
        private static readonly string[] H_Mobile     = { "мобильный номер", "мобильный", "сотовый", "телефон мобильный" };
        private static readonly string[] H_Internal   = { "внутренний телефон", "внутренний", "доб", "доб." };

        public string CreateVpkCompatibleWorkbook(string sourceExcelPath)
        {
            using var wb = ExcelUtils.Open(sourceExcelPath);
            // В книге один лист «ЗЗГТ», но на всякий дефолтимся к первому
            var sh = wb.GetSheet("ЗЗГТ") ?? wb.GetSheetAt(0);
            if (sh == null) throw new InvalidOperationException("Не найден лист Excel.");

            var header = sh.GetRow(sh.FirstRowNum);
            if (header == null) throw new InvalidOperationException("Не найдена строка заголовков.");

            // Карта "нормализованный заголовок" -> индекс
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

            var rows = new List<NormalizedContactRow>();

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

                // Собираем локацию
                string location = CombineNonEmpty(org, dep, sep: " / ");

                // Выбираем телефон: приоритет мобильному, иначе городской (код + номер)
                string phone = ChoosePhone(mobile, cityCode, cityNumber);

                // Пустые полностью строки пропускаем
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

            // Пишем временный .xlsx с нужными индексами колонок для уже готового VPK-конвертера
            return VpkNormalizedWorkbookBuilder.BuildTempWorkbook(rows, "ZZGT_");
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
            return -1; // не нашли — поле необязательное
        }

        private static string NormalizeHeader(string s)
        {
            // в нижний регистр; убрать переводы строк/неразрывные пробелы; схлопнуть пробелы; убрать точки
            var chars = s.Replace('\u00A0', ' ')  // NBSP -> обычный пробел
                         .Replace('\r', ' ')
                         .Replace('\n', ' ')
                         .Trim()
                         .ToLowerInvariant();

            // схлопнем множественные пробелы
            while (chars.Contains("  ")) chars = chars.Replace("  ", " ");

            // уберём точки
            chars = chars.Replace(".", "");
            return chars;
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

        private static string ChoosePhone(string mobile, string cityCode, string cityNumber)
        {
            string m = CleanSpaces(mobile);
            if (!string.IsNullOrWhiteSpace(m)) return m;

            string cc = CleanSpaces(cityCode);
            string cn = CleanSpaces(cityNumber);

            if (!string.IsNullOrWhiteSpace(cc) && !string.IsNullOrWhiteSpace(cn))
                return $"{cc} {cn}".Trim();

            // если один из них пуст — вернём то, что есть
            return (!string.IsNullOrWhiteSpace(cn) ? cn : cc) ?? "";
        }

        private static string CleanSpaces(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            // убираем неразрывные пробелы и схлопываем повторяющиеся
            var x = s.Replace('\u00A0', ' ').Trim();
            while (x.Contains("  ")) x = x.Replace("  ", " ");
            return x;
        }
    }
}
