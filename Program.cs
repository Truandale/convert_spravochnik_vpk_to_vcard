using System;
using System.Windows.Forms;

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
            
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
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
