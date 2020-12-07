namespace UtnCreateUpdateMaxnettoScript
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
            this.btnLoadFromTable = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lvTable = new System.Windows.Forms.ListView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbScript = new System.Windows.Forms.ListBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnSaveToFile = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnLoadFromTable
            // 
            this.btnLoadFromTable.Location = new System.Drawing.Point(14, 14);
            this.btnLoadFromTable.Name = "btnLoadFromTable";
            this.btnLoadFromTable.Size = new System.Drawing.Size(142, 26);
            this.btnLoadFromTable.TabIndex = 0;
            this.btnLoadFromTable.Text = "Загрузить таблицу...";
            this.btnLoadFromTable.UseVisualStyleBackColor = true;
            this.btnLoadFromTable.Click += new System.EventHandler(this.btnLoadFromTable_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(14, 46);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(825, 395);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lvTable);
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(817, 367);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Исходная таблица";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // lvTable
            // 
            this.lvTable.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvTable.FullRowSelect = true;
            this.lvTable.Location = new System.Drawing.Point(7, 7);
            this.lvTable.Name = "lvTable";
            this.lvTable.ShowItemToolTips = true;
            this.lvTable.Size = new System.Drawing.Size(801, 349);
            this.lvTable.TabIndex = 0;
            this.lvTable.UseCompatibleStateImageBehavior = false;
            this.lvTable.View = System.Windows.Forms.View.Details;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lbScript);
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(817, 367);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Результирующий скрипт";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lbScript
            // 
            this.lbScript.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbScript.FormattingEnabled = true;
            this.lbScript.ItemHeight = 15;
            this.lbScript.Location = new System.Drawing.Point(6, 6);
            this.lbScript.Name = "lbScript";
            this.lbScript.Size = new System.Drawing.Size(805, 349);
            this.lbScript.TabIndex = 0;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "xls";
            this.openFileDialog1.Filter = "*.xls|*.xls|*.xlsx|*.xlsх";
            // 
            // btnSaveToFile
            // 
            this.btnSaveToFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSaveToFile.Enabled = false;
            this.btnSaveToFile.Location = new System.Drawing.Point(699, 14);
            this.btnSaveToFile.Name = "btnSaveToFile";
            this.btnSaveToFile.Size = new System.Drawing.Size(142, 26);
            this.btnSaveToFile.TabIndex = 0;
            this.btnSaveToFile.Text = "Выгрузить в файл...";
            this.btnSaveToFile.UseVisualStyleBackColor = true;
            this.btnSaveToFile.Click += new System.EventHandler(this.btnSaveToFile_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "sql";
            this.saveFileDialog1.FileName = "UpdateMaxnettoScript.sql";
            this.saveFileDialog1.Filter = "*.sql|*.sql";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(853, 455);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnSaveToFile);
            this.Controls.Add(this.btnLoadFromTable);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds;
            this.Text = "Создание скрипта на изменение MAXNETTO для вагонов";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnLoadFromTable;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ListView lvTable;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ListBox lbScript;
        private System.Windows.Forms.Button btnSaveToFile;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

