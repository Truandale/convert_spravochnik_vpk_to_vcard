using System;
using System.Windows.Forms;

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
            MessageBox.Show("ЗЗГТ: Алгоритм в разработке", "Информация");
        }

        private void BtnGroupVPK_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("Группа ВПК: Алгоритм в разработке", "Информация");
        }

        private void BtnZavodKorpusov_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("Завод Корпусов: Алгоритм в разработке", "Информация");
        }
    }
}