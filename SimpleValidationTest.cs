using System;
using Converter.Parsing;

namespace convert_spravochnik_vpk_to_vcard
{
    public static class SimpleValidationTest
    {
        public static void TestParserValidation()
        {
            Console.WriteLine("=== Тест валидации парсеров (каждый должен принимать только свой файл) ===\n");
            
            var files = new[]
            {
                ("ВИЦ", @"c:\Users\trubnikovaa\Documents\Справочники\ВИЦ.xlsx"),
                ("ВЗК", @"c:\Users\trubnikovaa\Documents\Справочники\ВЗК.xlsx"),
                ("ЗЗГТ", @"c:\Users\trubnikovaa\Documents\Справочники\ЗЗГТ.xlsx")
            };
            
            // Тестируем ВИЦ парсер против всех файлов
            Console.WriteLine("=== ParserGroupVPK (для ВИЦ) ===");
            foreach (var (fileName, filePath) in files)
            {
                Console.Write($"  {fileName}: ");
                try
                {
                    var parser = new ParserGroupVPK();
                    string tempFile = parser.CreateVpkCompatibleWorkbook(filePath);
                    
                    // Если дошли сюда - файл принят
                    bool shouldAccept = fileName == "ВИЦ";
                    Console.WriteLine(shouldAccept ? "✅ Правильно принял" : "❌ Ошибочно принял");
                    
                    if (System.IO.File.Exists(tempFile))
                        System.IO.File.Delete(tempFile);
                }
                catch (Exception ex)
                {
                    // Файл отклонен
                    bool shouldReject = fileName != "ВИЦ";
                    Console.WriteLine(shouldReject ? "✅ Правильно отклонил" : $"❌ Ошибочно отклонил: {ex.Message}");
                }
            }
            
            // Тестируем ВЗК парсер против всех файлов
            Console.WriteLine("\n=== ParserVZK (для ВЗК) ===");
            foreach (var (fileName, filePath) in files)
            {
                Console.Write($"  {fileName}: ");
                try
                {
                    var parser = new ParserVZK();
                    string tempFile = parser.CreateVpkCompatibleWorkbook(filePath);
                    
                    // Если дошли сюда - файл принят
                    bool shouldAccept = fileName == "ВЗК";
                    Console.WriteLine(shouldAccept ? "✅ Правильно принял" : "❌ Ошибочно принял");
                    
                    if (System.IO.File.Exists(tempFile))
                        System.IO.File.Delete(tempFile);
                }
                catch (Exception ex)
                {
                    // Файл отклонен
                    bool shouldReject = fileName != "ВЗК";
                    Console.WriteLine(shouldReject ? "✅ Правильно отклонил" : $"❌ Ошибочно отклонил: {ex.Message}");
                }
            }
            
            // Тестируем ЗЗГТ парсер против всех файлов
            Console.WriteLine("\n=== ParserZZGT (для ЗЗГТ) ===");
            foreach (var (fileName, filePath) in files)
            {
                Console.Write($"  {fileName}: ");
                try
                {
                    var parser = new ParserZZGT();
                    string tempFile = parser.CreateVpkCompatibleWorkbook(filePath);
                    
                    // Если дошли сюда - файл принят
                    bool shouldAccept = fileName == "ЗЗГТ";
                    Console.WriteLine(shouldAccept ? "✅ Правильно принял" : "❌ Ошибочно принял");
                    
                    if (System.IO.File.Exists(tempFile))
                        System.IO.File.Delete(tempFile);
                }
                catch (Exception ex)
                {
                    // Файл отклонен
                    bool shouldReject = fileName != "ЗЗГТ";
                    Console.WriteLine(shouldReject ? "✅ Правильно отклонил" : $"❌ Ошибочно отклонил: {ex.Message}");
                }
            }
        }
    }
}
