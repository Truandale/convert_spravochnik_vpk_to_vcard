using NPOI.SS.UserModel;
using System.IO;

namespace convert_spravochnik_vpk_to_vcard
{
    /// <summary>
    /// Универсальный открыватель Excel файлов (.xls и .xlsx)
    /// </summary>
    public static class WorkbookHelper
    {
        /// <summary>
        /// Открывает Excel файл любого формата через WorkbookFactory
        /// </summary>
        public static IWorkbook OpenWorkbook(string path)
        {
            using var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return WorkbookFactory.Create(fs); // сам определит XLS/XLSX
        }
    }
}
