namespace convert_spravochnik_vpk_to_vcard
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnVPK = new System.Windows.Forms.Button();
            this.btnZZGT = new System.Windows.Forms.Button();
            this.btnGroupVPK = new System.Windows.Forms.Button();
            this.btnZavodKorpusov = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // btnVPK
            // 
            this.btnVPK.Location = new System.Drawing.Point(30, 30);
            this.btnVPK.Name = "btnVPK";
            this.btnVPK.Size = new System.Drawing.Size(150, 40);
            this.btnVPK.TabIndex = 0;
            this.btnVPK.Text = "ВПК";
            this.btnVPK.UseVisualStyleBackColor = true;
            this.btnVPK.Click += new System.EventHandler(this.BtnVPK_Click);
            // 
            // btnZZGT
            // 
            this.btnZZGT.Location = new System.Drawing.Point(200, 30);
            this.btnZZGT.Name = "btnZZGT";
            this.btnZZGT.Size = new System.Drawing.Size(150, 40);
            this.btnZZGT.TabIndex = 1;
            this.btnZZGT.Text = "ЗЗГТ";
            this.btnZZGT.UseVisualStyleBackColor = true;
            this.btnZZGT.Click += new System.EventHandler(this.BtnZZGT_Click);
            // 
            // btnGroupVPK
            // 
            this.btnGroupVPK.Location = new System.Drawing.Point(30, 90);
            this.btnGroupVPK.Name = "btnGroupVPK";
            this.btnGroupVPK.Size = new System.Drawing.Size(150, 40);
            this.btnGroupVPK.TabIndex = 2;
            this.btnGroupVPK.Text = "Группа ВПК";
            this.btnGroupVPK.UseVisualStyleBackColor = true;
            this.btnGroupVPK.Click += new System.EventHandler(this.BtnGroupVPK_Click);
            // 
            // btnZavodKorpusov
            // 
            this.btnZavodKorpusov.Location = new System.Drawing.Point(200, 90);
            this.btnZavodKorpusov.Name = "btnZavodKorpusov";
            this.btnZavodKorpusov.Size = new System.Drawing.Size(150, 40);
            this.btnZavodKorpusov.TabIndex = 3;
            this.btnZavodKorpusov.Text = "ВЗК";
            this.btnZavodKorpusov.UseVisualStyleBackColor = true;
            this.btnZavodKorpusov.Click += new System.EventHandler(this.BtnZavodKorpusov_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";
            this.openFileDialog.Title = "Выберите Excel файл справочника";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "vcf";
            this.saveFileDialog.Filter = "vCard (*.vcf)|*.vcf";
            this.saveFileDialog.Title = "Куда сохранить vCard файл";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 161);
            this.Controls.Add(this.btnZavodKorpusov);
            this.Controls.Add(this.btnGroupVPK);
            this.Controls.Add(this.btnZZGT);
            this.Controls.Add(this.btnVPK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертер справочников в vCard";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnVPK;
        private System.Windows.Forms.Button btnZZGT;
        private System.Windows.Forms.Button btnGroupVPK;
        private System.Windows.Forms.Button btnZavodKorpusov;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }
}
