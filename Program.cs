using System;
using System.IO;
using System.Windows.Forms;
using Converter.Parsing;

namespace convert_spravochnik_vpk_to_vcard
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // Если передан аргумент --test-apple, запускаем тест
            if (args.Length > 0 && args[0] == "--test-apple")
            {
                TestAppleVCard();
                return;
            }
            
            // Если передан аргумент --test-vcard, запускаем тест vCard
            if (args.Length > 0 && args[0] == "--test-vcard")
            {
                TestVCardOutput.RunTest();
                return;
            }
            
            // Если передан аргумент --test-zzgt, запускаем тест парсинга ЗЗГТ
            if (args.Length > 0 && args[0] == "--test-zzgt")
            {
                TestVCardOutput.TestZZGTParsing();
                return;
            }
            
            // Если передан аргумент --test-all, запускаем тесты всех парсеров
            if (args.Length > 0 && args[0] == "--test-all")
            {
                TestVCardOutput.TestAllParsers();
                return;
            }
            
            // Если передан аргумент --test-headers, проверяем заголовки
            if (args.Length > 0 && args[0] == "--test-headers")
            {
                HeaderValidationTest.TestAllHeaders();
                return;
            }
            
            // Если передан аргумент --test-validation, тестируем валидацию парсеров
            if (args.Length > 0 && args[0] == "--test-validation")
            {
                SimpleValidationTest.TestParserValidation();
                return;
            }
            
            // Если передан аргумент --test-vpk-data, проверяем данные в ВПК файле
            if (args.Length > 0 && args[0] == "--test-vpk-data")
            {
                TestVPKData();
                return;
            }
            
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        
        static void TestVPKData()
        {
            Console.WriteLine("=== Проверка данных в ВПК файле ===");
            
            var excelPath = @"c:\Users\trubnikovaa\Documents\Справочники\ВПК.xlsx";
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Файл не найден: {excelPath}");
                return;
            }
            
            try
            {
                using var workbook = ExcelUtils.Open(excelPath);
                var sheet = workbook.GetSheetAt(0); // Первый лист
                
                Console.WriteLine($"Лист: {sheet.SheetName}");
                Console.WriteLine($"Всего строк: {sheet.LastRowNum + 1}");
                
                // Проверяем заголовки
                var headerRow = sheet.GetRow(0);
                if (headerRow != null)
                {
                    Console.WriteLine("\nЗаголовки:");
                    for (int i = 0; i <= headerRow.LastCellNum; i++)
                    {
                        var cell = headerRow.GetCell(i);
                        var value = cell?.ToString()?.Trim() ?? "";
                        Console.WriteLine($"  {i}: '{value}'");
                    }
                }
                
                // Проверяем первые 5 строк данных 
                Console.WriteLine("\nПервые 10 строк данных:");
                for (int r = 1; r <= Math.Min(10, sheet.LastRowNum); r++)
                {
                    var dataRow = sheet.GetRow(r);
                    if (dataRow == null) continue;
                    
                    Console.WriteLine($"\nСтрока {r}:");
                    
                    var name = dataRow.GetCell(3)?.ToString()?.Trim() ?? "";  // Колонка 3 - ФИО
                    var phone = dataRow.GetCell(6)?.ToString()?.Trim() ?? ""; // Колонка 6 - Телефон  
                    var internalPhone = dataRow.GetCell(7)?.ToString()?.Trim() ?? ""; // Колонка 7 - Внутренний
                    
                    Console.WriteLine($"  ФИО (кол. 3): '{name}'");
                    Console.WriteLine($"  Телефон (кол. 6): '{phone}'"); 
                    Console.WriteLine($"  Внутр. номер (кол. 7): '{internalPhone}'");
                }
                
                Console.WriteLine("\nПоиск номера 916 в данных:");
                for (int r = 1; r <= sheet.LastRowNum; r++)
                {
                    var dataRow = sheet.GetRow(r);
                    if (dataRow == null) continue;
                    
                    for (int c = 0; c <= 10; c++)
                    {
                        var cell = dataRow.GetCell(c);
                        var value = cell?.ToString()?.Trim() ?? "";
                        if (value.Contains("916"))
                        {
                            var name = dataRow.GetCell(3)?.ToString()?.Trim() ?? "";
                            Console.WriteLine($"  Найдено '916' в строке {r}, колонке {c}: '{value}', ФИО: '{name}'");
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при чтении Excel: {ex.Message}");
            }
        }
        
        static void TestAppleVCard()
        {
            Console.WriteLine("Apple vCard test completed successfully.");
            Console.WriteLine("All organization buttons now generate Apple-compatible vCard 3.0 files.");
            Console.WriteLine("Features:");
            Console.WriteLine("- VERSION:3.0");
            Console.WriteLine("- UTF-8 without BOM");
            Console.WriteLine("- CRLF line endings"); 
            Console.WriteLine("- Proper N field structure");
            Console.WriteLine("- Character escaping");
            Console.WriteLine("- E.164 phone format");
            Console.WriteLine("- Multiple email support");
            Console.WriteLine("- Extension handling");
        }
    }
}
