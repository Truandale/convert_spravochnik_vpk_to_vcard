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

                string org        = CleanTrailingLiteralNewlines(Read(row, idxOrg));
                string dep        = CleanTrailingLiteralNewlines(Read(row, idxDep));
                string name       = CleanTrailingLiteralNewlines(Read(row, idxName));
                string position   = CleanTrailingLiteralNewlines(Read(row, idxPos));
                string email      = CleanTrailingLiteralNewlines(Read(row, idxEmail));
                string cityCode   = Read(row, idxCityCode);
                string cityNumber = Read(row, idxCityNumber);
                string mobile     = Read(row, idxMobile);
                string internalNo = Read(row, idxInternal);

                // Собираем локацию для поля ORG (организация + подразделение)
                string location = CombineNonEmpty(org, dep, sep: " / ");

                // Выбираем телефон: приоритет мобильному, иначе городской (код + номер)
                string phone = ChoosePhone(mobile, cityCode, cityNumber);

                // Добавочный номер: если он есть, и есть основной номер - сохраняем для обработки как ext
                // Если основного нет, добавочный пойдет в InternalPhone для записи в NOTE
                string cleanInternal = CleanSpaces(internalNo);
                
                // Определяем, является ли внутренний номер добавочным (3-5 цифр)
                bool isExtension = !string.IsNullOrEmpty(cleanInternal) && 
                                  cleanInternal.All(char.IsDigit) && 
                                  cleanInternal.Length >= 3 && 
                                  cleanInternal.Length <= 5;

                string finalInternalPhone = "";
                
                // Если есть основной номер и добавочный - добавочный будет обработан как ext
                // Если нет основного номера, но есть добавочный - записываем его для NOTE
                if (isExtension)
                {
                    if (string.IsNullOrEmpty(phone))
                    {
                        // Нет основного номера - добавочный в NOTE
                        finalInternalPhone = cleanInternal;
                    }
                    // Если есть основной номер, добавочный будет в phone как ext, InternalPhone остается пустым
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

                // Пустые полностью строки пропускаем
                if (string.IsNullOrWhiteSpace(name) &&
                    string.IsNullOrWhiteSpace(email) &&
                    string.IsNullOrWhiteSpace(phone) &&
                    string.IsNullOrWhiteSpace(finalInternalPhone))
                    continue;

                rows.Add(new NormalizedContactRow
                {
                    Location      = location,  // Теперь это "Организация / Подразделение"
                    Name          = name,
                    Position      = position,
                    Email         = email,
                    Phone         = !string.IsNullOrEmpty(phone) && isExtension && !string.IsNullOrEmpty(cleanInternal) 
                                   ? $"{phone};ext={cleanInternal}" : phone,  // Основной номер с добавочным
                    InternalPhone = finalInternalPhone  // Либо пустой, либо номер для NOTE
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
            // 1) приоритет мобильному
            var m = RuPhone.NormalizeToE164RU(CleanSpaces(mobile));
            if (!string.IsNullOrEmpty(m)) return m;

            // 2) иначе городской: код + номер -> +7XXXXXXXXXX
            var city = RuPhone.ComposeCityToE164RU(CleanSpaces(cityCode), CleanSpaces(cityNumber));
            if (!string.IsNullOrEmpty(city)) return city;

            return "";
        }

        private static string CleanSpaces(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            // убираем неразрывные пробелы и схлопываем повторяющиеся
            var x = s.Replace('\u00A0', ' ').Trim();
            while (x.Contains("  ")) x = x.Replace("  ", " ");
            return x;
        }

        /// <summary>
        /// Убирает завершающие литералы \n (обратный слэш + n) из строки
        /// </summary>
        private static string CleanTrailingLiteralNewlines(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            
            // Убираем завершающие "\n" (как литеральные символы \ и n)
            while (s.EndsWith("\\n"))
            {
                s = s.Substring(0, s.Length - 2);
            }
            
            return s.Trim();
        }
    }
}
