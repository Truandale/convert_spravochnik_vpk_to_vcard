namespace Converter.Parsing
{
    /// <summary>
    /// Интерфейс для парсеров Excel файлов разных организаций
    /// </summary>
    public interface IExcelParser
    {
        /// <summary>
        /// Название организации/формата
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Создает временный Excel файл в формате, совместимом с VPKConverter
        /// </summary>
        /// <param name="sourceExcelPath">Путь к исходному Excel файлу</param>
        /// <returns>Путь к временному Excel файлу в формате VPK</returns>
        string CreateVpkCompatibleWorkbook(string sourceExcelPath);
    }
}
