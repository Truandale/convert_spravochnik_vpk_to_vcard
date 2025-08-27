using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Converter.Parsing
{
    /// <summary>
    /// Парсер для кнопки «ВЗК».
    /// Приводит Excel к формату, который ожидает VPKConverter (колонки: Location, Name, Position, Email, Phone, InternalPhone).
    /// </summary>
    public sealed class ParserVZK : IExcelParser
    {
        public string Name => "ВЗК";

        // Варианты заголовков (с запасом)
        private static readonly string[] H_Org        = { "организация" };
        private static readonly string[] H_Department = {
            "структурное подразделение/ департамент", "структурное подразделение",
            "департамент", "отдел", "служба", "подразделение"
        };
        private static readonly string[] H_Name       = { "фио", "ф.и.о", "фио сотрудника" };
        private static readonly string[] H_Position   = { "должность", "роль", "position", "title" };
        private static readonly string[] H_Email      = { "электронный адрес", "email", "e-mail", "почта" };
        private static readonly string[] H_CityCode   = { "код города", "городской код", "код" };
        private static readonly string[] H_CityNumber = { "городской номер", "городской", "телефон городской" };
        private static readonly string[] H_Mobile     = { "мобильный номер", "мобильный", "сотовый", "телефон мобильный" };
        private static readonly string[] H_Internal   = { "внутренний телефон", "внутренний", "доб", "доб." };

        public string CreateVpkCompatibleWorkbook(string sourceExcelPath)
        {
            using var wb = ExcelUtils.Open(sourceExcelPath);

            // В твоём файле лист «АО Завод Корпусов», но берём первый на всякий случай
            var sh = wb.GetSheet("АО Завод Корпусов") ?? wb.GetSheetAt(0);
            if (sh == null) throw new InvalidOperationException("Не найден лист Excel.");

            var header = sh.GetRow(sh.FirstRowNum);
            if (header == null) throw new InvalidOperationException("Не найдена строка заголовков.");

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

                string org        = CleanTextQuality(Read(row, idxOrg));
                string dep        = CleanTextQuality(Read(row, idxDep));
                string name       = CleanTextQuality(Read(row, idxName));
                string position   = CleanTextQuality(Read(row, idxPos));
                string email      = Read(row, idxEmail);
                string cityCode   = Read(row, idxCityCode);
                string cityNumber = Read(row, idxCityNumber);
                string mobile     = Read(row, idxMobile);
                string internalNo = Read(row, idxInternal);

                // Организация для ORG поля (а не Location)
                string organization = CombineNonEmpty(org, dep, " / ");

                // Телефон по приоритету + нормализация к +7
                string phone = ChoosePhone(mobile, cityCode, cityNumber);

                // Обработка email: разбиваем на отдельные адреса
                var emails = SplitEmails(email);
                string primaryEmail = emails.FirstOrDefault() ?? "";

                // Обработка добавочного номера
                string cleanInternal = Clean(internalNo);
                bool isExtension = !string.IsNullOrEmpty(cleanInternal) && 
                                  cleanInternal.All(char.IsDigit) && 
                                  cleanInternal.Length >= 3 && 
                                  cleanInternal.Length <= 5;

                string finalInternalPhone = "";
                
                if (isExtension)
                {
                    if (!string.IsNullOrEmpty(phone))
                    {
                        // Есть основной номер - добавляем extension в правильном RFC 3966 формате
                        // Очищаем основной номер от скобок и пробелов перед ext
                        string cleanPhone = RuPhone.NormalizeToE164RU(phone);
                        if (!string.IsNullOrEmpty(cleanPhone))
                        {
                            phone = cleanPhone; // Используем нормализованный номер без ext
                            finalInternalPhone = cleanInternal; // Добавочный пойдет в NOTE или как ext в AppleVCardWriter
                        }
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

                // Объединяем все email адреса в одну строку через точку с запятой
                var allEmails = string.Join("; ", emails);

                // Пустые полностью строки не тянем
                if (string.IsNullOrWhiteSpace(name) &&
                    string.IsNullOrWhiteSpace(allEmails) &&
                    string.IsNullOrWhiteSpace(phone) &&
                    string.IsNullOrWhiteSpace(finalInternalPhone))
                    continue;

                var contactRow = new NormalizedContactRow
                {
                    Location      = organization,  // Теперь это правильная организация для ORG
                    Name          = name,
                    Position      = position,
                    Email         = allEmails,  // Все email объединены через точку с запятой
                    Phone         = phone,
                    InternalPhone = finalInternalPhone
                };

                rows.Add(contactRow);
            }

            return VpkNormalizedWorkbookBuilder.BuildTempWorkbook(rows, "VZK_");
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

        private static string Clean(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var x = s.Replace('\u00A0', ' ').Trim();
            while (x.Contains("  ")) x = x.Replace("  ", " ");
            return x;
        }

        /// <summary>
        /// Улучшенная очистка текста: убирает двойные пробелы, исправляет склейки и типичные опечатки
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
            
            // Исправляем типичные опечатки в ВЗК
            result = result.Replace("режмно-секретный", "режимно-секретный");
            result = result.Replace("Режмно-секретный", "Режимно-секретный");
            result = result.Replace("режмно-секретного", "режимно-секретного");
            result = result.Replace("Режмно-секретного", "Режимно-секретного");
            result = result.Replace("отел", "отдел");
            result = result.Replace("Отел", "Отдел");
            result = result.Replace("иремонта", "и ремонта");
            result = result.Replace("обслуживания иремонта", "обслуживания и ремонта");
            
            return result;
        }

        /// <summary>
        /// Разбивает строку email на отдельные адреса и фильтрует пустые
        /// </summary>
        private static List<string> SplitEmails(string emailStr)
        {
            if (string.IsNullOrWhiteSpace(emailStr)) return new List<string>();
            
            var emails = emailStr.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Select(e => e.Trim())
                                 .Where(e => !string.IsNullOrEmpty(e) && e.Contains("@"))
                                 .ToList();
            
            return emails;
        }

        private static string ChoosePhone(string mobile, string cityCode, string cityNumber)
        {
            // 1) мобильный → +7XXXXXXXXXX
            var m = RuPhone.NormalizeToE164RU(Clean(mobile));
            if (!string.IsNullOrEmpty(m)) return m;

            // 2) городской (код + номер) → +7XXXXXXXXXX
            var city = RuPhone.ComposeCityToE164RU(Clean(cityCode), Clean(cityNumber));
            if (!string.IsNullOrEmpty(city)) return city;

            return "";
        }
    }
}
