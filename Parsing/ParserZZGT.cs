using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using convert_spravochnik_vpk_to_vcard;

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
            using var wb = WorkbookHelper.OpenWorkbook(sourceExcelPath);
            
            bool anyValid = false;
            var rows = new List<NormalizedContactRow>();

            // Проверяем все листы книги с жёсткой валидацией
            for (int s = 0; s < wb.NumberOfSheets; s++)
            {
                var sh = wb.GetSheetAt(s);
                var validation = StrictSchemaValidator.ValidateZZGT(sh);
                
                if (!validation.IsValid) 
                { 
                    Console.WriteLine($"[ЗЗГТ] Пропуск «{sh.SheetName}»: {validation.Reason}"); 
                    continue; 
                }

                anyValid = true;
                Console.WriteLine($"[ЗЗГТ] Обрабатываем лист «{sh.SheetName}»");

                // Парсинг как в 18b487d: находим заголовки и колонки по имени
                var (headerRowIndex, headersRaw, headersCanon) = HeaderFinder.FindHeaderRow(sh);
                var cols = HeaderFinder.MapColumns("ЗЗГТ", headersRaw, headersCanon);

                Console.WriteLine($"[ЗЗГТ] Найдена строка заголовков: {headerRowIndex}, колонки: {string.Join(", ", cols.Keys)}");

                for (int r = headerRowIndex + 1; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    if (row == null) continue;

                    // Читаем поля через HeaderFinder
                    string fio = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("fio", -1));
                    string title = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("title", -1));
                    string email = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("email", -1));
                    string cityCode = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("code", -1));
                    string cityNumber = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("city", -1));
                    string mobile = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("mobile", -1));
                    string internalNo = HeaderFinder.ReadCell(row, cols.GetValueOrDefault("ext", -1));

                    // Пропускаем строки без основных данных (телефонов)
                    if (string.IsNullOrWhiteSpace(mobile) && (string.IsNullOrWhiteSpace(cityCode) || string.IsNullOrWhiteSpace(cityNumber)))
                        continue;

                    // Выбираем телефон: приоритет мобильному, иначе городской (код + номер)
                    string phone = ChoosePhone(mobile, cityCode, cityNumber);

                    // Добавочный номер: если он есть, и есть основной номер - сохраняем для обработки как ext
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
                    if (string.IsNullOrWhiteSpace(fio) &&
                        string.IsNullOrWhiteSpace(email) &&
                        string.IsNullOrWhiteSpace(phone) &&
                        string.IsNullOrWhiteSpace(finalInternalPhone))
                        continue;

                    rows.Add(new NormalizedContactRow
                    {
                        Location      = "ЗЗГТ",  // Организация
                        Name          = fio,
                        Position      = title,
                        Email         = email,
                        Phone         = !string.IsNullOrEmpty(phone) && isExtension && !string.IsNullOrEmpty(cleanInternal) 
                                       ? $"{phone};ext={cleanInternal}" : phone,  // Основной номер с добавочным
                        InternalPhone = finalInternalPhone  // Либо пустой, либо номер для NOTE
                    });
                }
            }

            if (!anyValid)
                throw new InvalidOperationException("Справочник не соответствует кнопке «ЗЗГТ». Выберите корректный файл.");

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

            // 2) городской с валидацией склейки (специально для ЗЗГТ)
            var city = ComposeCityWithValidation(CleanSpaces(cityCode), CleanSpaces(cityNumber));
            if (!string.IsNullOrEmpty(city)) return city;

            return "";
        }

        /// <summary>
        /// Строгая валидация склейки городского номера для ЗЗГТ
        /// Проверяет что код+номер дают корректную российскую нумерацию
        /// </summary>
        private static string ComposeCityWithValidation(string cityCode, string cityNumber)
        {
            if (string.IsNullOrWhiteSpace(cityCode) || string.IsNullOrWhiteSpace(cityNumber))
                return "";

            var cc = new string(cityCode.Where(char.IsDigit).ToArray());
            var cn = new string(cityNumber.Where(char.IsDigit).ToArray());
            var joined = cc + cn;

            // Валидация: должно быть 9-10 цифр для нормальной российской нумерации
            if (joined.Length < 9 || joined.Length > 10)
            {
                // Логируем аномалию для отладки
                System.Diagnostics.Debug.WriteLine($"ЗЗГТ: Странная длина городского номера: код='{cityCode}' номер='{cityNumber}' -> {joined.Length} цифр");
                return "";
            }

            var result = "+7" + joined;
            
            // Финальная проверка: результат должен быть +7 + ровно 10 цифр (итого 12 символов)
            if (result.Length != 12)
            {
                System.Diagnostics.Debug.WriteLine($"ЗЗГТ: Некорректная финальная длина: {result} ({result.Length} символов)");
                return "";
            }

            return result;
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
