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
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        VPKConverter.Convert(openFileDialog.FileName, saveFileDialog.FileName);
                        MessageBox.Show("Конвертация завершена успешно!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при конвертации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnZZGT_Click(object? sender, EventArgs e)
        {
            RunParser(new ParserZZGT());
        }

        private void BtnGroupVPK_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("Группа ВПК: Алгоритм в разработке", "Информация");
        }

        private void BtnZavodKorpusov_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("Завод Корпусов: Алгоритм в разработке", "Информация");
        }

        private void RunParser(IExcelParser parser)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string? tempVpkFile = null;
                    try
                    {
                        // Создаем временный файл в формате VPK
                        tempVpkFile = parser.CreateVpkCompatibleWorkbook(openFileDialog.FileName);
                        
                        // Используем существующий VPKConverter
                        VPKConverter.Convert(tempVpkFile, saveFileDialog.FileName);
                        
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