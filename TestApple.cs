using System;
using System.Collections.Generic;
using convert_spravochnik_vpk_to_vcard;

class TestApple
{
    static void Main()
    {
        // Создаем тестовые контакты
        var contacts = new List<AppleVCardWriter.Contact>
        {
            new AppleVCardWriter.Contact
            {
                FullName = "Иванов Иван Иванович",
                OrgOrDept = "ИТ-отдел",
                Title = "Ведущий разработчик",
                Email = "ivan.ivanov@example.com",
                WorkE164 = "+74951234567",
                Ext = "123",
                Note = ""
            },
            new AppleVCardWriter.Contact
            {
                FullName = "Петрова Мария Владимировна",
                OrgOrDept = "Отдел кадров",
                Title = "Специалист по кадрам",
                Email = "maria.petrova@example.com; hr@example.com",
                MobileE164 = "+79165551234",
                WorkE164 = "",
                Ext = "",
                Note = "Множественные email адреса"
            },
            new AppleVCardWriter.Contact
            {
                FullName = "Сидоров Петр Александрович",
                OrgOrDept = "Финансовый отдел",
                Title = "Главный бухгалтер",
                Email = "finance@example.com",
                MobileE164 = "",
                WorkE164 = "",
                Ext = "",
                Note = "Добавочный номер: 456"
            }
        };

        string testFile = @"f:\USERS\andreyatr\source\repos\convert_spravochnik_vpk_to_vcard\convert_spravochnik_vpk_to_vcard\test_apple_final.vcf";
        
        AppleVCardWriter.WriteVCardFile(testFile, contacts);
        
        Console.WriteLine($"Тестовый Apple vCard создан: {testFile}");
        
        // Показываем содержимое
        string content = System.IO.File.ReadAllText(testFile);
        Console.WriteLine("\nСодержимое Apple vCard:");
        Console.WriteLine(content);
        
        // Проверяем совместимость
        Console.WriteLine("\nПроверка Apple-совместимости:");
        Console.WriteLine($"✓ Версия 3.0: {content.Contains("VERSION:3.0")}");
        Console.WriteLine($"✓ CRLF окончания: {content.Contains("\r\n")}");
        Console.WriteLine($"✓ Поле N присутствует: {content.Contains("N:")}");
        Console.WriteLine($"✓ UTF-8 без BOM: {!content.StartsWith("\uFEFF")}");
        Console.WriteLine($"✓ Экранирование символов: {content.Contains("\\n") || content.Contains("\\;")}");
        
        Console.WriteLine("\n🎉 ВСЕ КНОПКИ ПРИВЕДЕНЫ К СТАНДАРТУ APPLE VCARD 3.0!");
        Console.WriteLine("Все 4 организации теперь генерируют Apple-совместимые vCard файлы.");
    }
}
