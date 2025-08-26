namespace Converter.Parsing
{
    /// <summary>
    /// Нормализованная строка контакта для создания VPK-совместимого файла
    /// </summary>
    public class NormalizedContactRow
    {
        public string Location { get; set; } = "";
        public string Name { get; set; } = "";
        public string Position { get; set; } = "";
        public string Email { get; set; } = "";
        public string Phone { get; set; } = "";
        public string InternalPhone { get; set; } = "";
    }
}
