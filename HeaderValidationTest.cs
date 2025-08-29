using System;
using System.Collections.Generic;
using Converter.Parsing;
using static Converter.Parsing.StrictSchemaValidator;

namespace convert_spravochnik_vpk_to_vcard
{
    public static class HeaderValidationTest
    {
        public static void TestAllHeaders()
        {
            Console.WriteLine("=== Проверка структуры заголовков всех справочников ===\n");
            
            var files = new[]
            {
                ("ВПК", @"c:\Users\trubnikovaa\Documents\Справочники\ВПК.xlsx"),
                ("ВИЦ", @"c:\Users\trubnikovaa\Documents\Справочники\ВИЦ.xlsx"),
                ("ВЗК", @"c:\Users\trubnikovaa\Documents\Справочники\ВЗК.xlsx"),
                ("ЗЗГТ", @"c:\Users\trubnikovaa\Documents\Справочники\ЗЗГТ.xlsx")
            };
            
            foreach (var (name, path) in files)
            {
                Console.WriteLine($"=== {name} ===");
                try
                {
                    using var wb = WorkbookHelper.OpenWorkbook(path);
                    for (int s = 0; s < wb.NumberOfSheets; s++)
                    {
                        var sheet = wb.GetSheetAt(s);
                        Console.WriteLine($"Лист: '{sheet.SheetName}'");
                        
                        // Найдем строку заголовков
                        var (headerRowIndex, headersRaw, headersCanon) = HeaderFinder.FindHeaderRow(sheet);
                        
                        Console.WriteLine($"Строка заголовков: {headerRowIndex}");
                        Console.WriteLine("Сырые заголовки:");
                        for (int i = 0; i < headersRaw.Length; i++)
                        {
                            Console.WriteLine($"  {i}: '{headersRaw[i]}'");
                        }
                        
                        Console.WriteLine("Канонические заголовки:");
                        for (int i = 0; i < headersCanon.Length; i++)
                        {
                            Console.WriteLine($"  {i}: '{headersCanon[i]}'");
                        }
                        
                        Console.WriteLine($"Количество колонок: {headersCanon.Length}");
                        Console.WriteLine();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Ошибка: {ex.Message}");
                }
                Console.WriteLine(new string('-', 50));
            }
        }
    }
}
