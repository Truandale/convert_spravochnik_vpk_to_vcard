using System;
using System.Collections.Generic;
using Converter.Parsing;

namespace convert_spravochnik_vpk_to_vcard
{
    public static class TestVCardOutput
    {
        public static void RunTest()
        {
            Console.WriteLine("=== Тест создания vCard ===");
            
            var testContact = new AppleVCardWriter.Contact
            {
                FullName = "Иванов Иван Иванович",
                OrgOrDept = "ВПК",
                Title = "Инженер",
                Email = "ivanov@vpk.ru",
                MobileE164 = "+79012345678",
                WorkE164 = "+74951234567",
                Ext = "123",
                Note = "Тестовый контакт"
            };

            var contacts = new List<AppleVCardWriter.Contact> { testContact };
            string outputPath = @"c:\Users\trubnikovaa\Documents\Справочники\test_output.vcf";
            
            try
            {
                AppleVCardWriter.WriteVCardFile(outputPath, contacts);
                Console.WriteLine($"✓ Тестовый vCard файл создан: {outputPath}");
                
                // Выводим содержимое на консоль
                var content = System.IO.File.ReadAllText(outputPath);
                Console.WriteLine("Содержимое файла:");
                Console.WriteLine(content);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Ошибка: {ex.Message}");
            }
        }
        
        public static void TestZZGTParsing()
        {
            Console.WriteLine("=== Тест парсинга ЗЗГТ ===");
            
            try
            {
                var parser = new ParserZZGT();
                string sourceFile = @"c:\Users\trubnikovaa\Documents\Справочники\ЗЗГТ.xlsx";
                string outputFile = @"c:\Users\trubnikovaa\Documents\Справочники\test_zzgt.vcf";
                
                Console.WriteLine($"Исходный файл: {sourceFile}");
                
                // Давайте сначала посмотрим на структуру файла
                using (var wb = WorkbookHelper.OpenWorkbook(sourceFile))
                {
                    var sheet = wb.GetSheetAt(0);
                    Console.WriteLine($"Лист: {sheet.SheetName}, LastRowNum: {sheet.LastRowNum}");
                    
                    // Покажем первые 5 строк
                    for (int r = 0; r <= Math.Min(5, sheet.LastRowNum); r++)
                    {
                        var row = sheet.GetRow(r);
                        if (row == null) 
                        {
                            Console.WriteLine($"Строка {r}: null");
                            continue;
                        }
                        
                        var cells = new List<string>();
                        for (int c = 0; c < Math.Min(10, (int)row.LastCellNum); c++)
                        {
                            var cell = row.GetCell(c);
                            cells.Add(cell?.ToString() ?? "");
                        }
                        Console.WriteLine($"Строка {r}: [{string.Join(" | ", cells)}]");
                    }
                }
                
                Console.WriteLine("Создаем временный VPK файл...");
                
                string tempVpkFile = parser.CreateVpkCompatibleWorkbook(sourceFile);
                Console.WriteLine($"Временный VPK файл: {tempVpkFile}");
                
                Console.WriteLine("Конвертируем в vCard...");
                
                // Давайте сначала посмотрим на временный VPK файл
                Console.WriteLine("=== Содержимое временного VPK файла ===");
                using (var tempWb = WorkbookHelper.OpenWorkbook(tempVpkFile))
                {
                    var tempSheet = tempWb.GetSheetAt(0);
                    Console.WriteLine($"Временный лист: {tempSheet.SheetName}, LastRowNum: {tempSheet.LastRowNum}");
                    
                    for (int r = 0; r <= Math.Min(5, tempSheet.LastRowNum); r++)
                    {
                        var row = tempSheet.GetRow(r);
                        if (row == null) 
                        {
                            Console.WriteLine($"Строка {r}: null");
                            continue;
                        }
                        
                        var cells = new List<string>();
                        for (int c = 0; c < Math.Min(10, (int)row.LastCellNum); c++)
                        {
                            var cell = row.GetCell(c);
                            cells.Add(cell?.ToString() ?? "");
                        }
                        Console.WriteLine($"Строка {r}: [{string.Join(" | ", cells)}]");
                    }
                }
                Console.WriteLine("=== Конец содержимого временного файла ===");
                
                VPKConverterFixed.Convert(tempVpkFile, outputFile);
                
                Console.WriteLine($"✓ vCard файл создан: {outputFile}");
                
                // Показываем первые несколько строк результата
                var lines = System.IO.File.ReadAllLines(outputFile);
                Console.WriteLine($"Файл содержит {lines.Length} строк. Первые 20 строк:");
                for (int i = 0; i < Math.Min(20, lines.Length); i++)
                {
                    Console.WriteLine($"{i+1:D2}: {lines[i]}");
                }
                
                // Удаляем временный файл
                if (System.IO.File.Exists(tempVpkFile))
                {
                    System.IO.File.Delete(tempVpkFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Ошибка парсинга ЗЗГТ: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
            }
        }
    }
}
