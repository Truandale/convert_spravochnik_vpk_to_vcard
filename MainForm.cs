using System;
using System.IO;
using System.Windows.Forms;
using Converter.Parsing;

namespace convert_spravochnik_vpk_to_vcard
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void BtnVPK_Click(object? sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] Кнопка ВПК нажата!");
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Выбран файл: {openFileDialog.FileName}");
                System.Diagnostics.Debug.WriteLine("[DEBUG] Показываем saveFileDialog...");
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Выбран выходной файл: {saveFileDialog.FileName}");
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("[DEBUG] Вызываем VPKConverterFixed.Convert...");
                        VPKConverterFixed.Convert(openFileDialog.FileName, saveFileDialog.FileName);
                        System.Diagnostics.Debug.WriteLine("[DEBUG] VPKConverterFixed.Convert завершен успешно");
                        MessageBox.Show("Конвертация завершена успешно!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ERROR] Исключение в MainForm: {ex.GetType().Name}: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"[ERROR] StackTrace: {ex.StackTrace}");
                        MessageBox.Show($"Ошибка при конвертации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] saveFileDialog отменен пользователем");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[DEBUG] openFileDialog отменен пользователем");
            }
        }

        private void BtnZZGT_Click(object? sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] Кнопка ЗЗГТ нажата!");
            RunParser(new ParserZZGT());
        }

        private void BtnGroupVPK_Click(object? sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] Кнопка GroupVPK (ВИЦ) нажата!");
            RunParser(new ParserGroupVPK());
        }

        private void BtnZavodKorpusov_Click(object? sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] Кнопка ВЗК нажата!");
            RunParser(new ParserVZK());
        }

        private void RunParser(IExcelParser parser)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] RunParser вызван для парсера: {parser.Name}");
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] RunParser - выбран файл: {openFileDialog.FileName}");
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] RunParser - выбран выходной файл: {saveFileDialog.FileName}");
                    string? tempVpkFile = null;
                    try
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] RunParser - создаем временный файл...");
                        // Создаем временный файл в формате VPK
                        tempVpkFile = parser.CreateVpkCompatibleWorkbook(openFileDialog.FileName);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] RunParser - временный файл создан: {tempVpkFile}");
                        
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] RunParser - вызываем VPKConverterFixed.Convert...");
                        // Используем исправленный VPKConverterFixed
                        VPKConverterFixed.Convert(tempVpkFile, saveFileDialog.FileName);
                        
                        MessageBox.Show($"Конвертация {parser.Name} завершена успешно!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при конвертации {parser.Name}: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        // Удаляем временный файл
                        if (tempVpkFile != null && File.Exists(tempVpkFile))
                        {
                            try
                            {
                                File.Delete(tempVpkFile);
                            }
                            catch
                            {
                                // Игнорируем ошибки удаления временного файла
                            }
                        }
                    }
                }
            }
        }
    }
}