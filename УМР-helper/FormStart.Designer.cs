
namespace УМР_helper
{
    partial class FormStart
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormStart));
            this.btnDownloadFile = new System.Windows.Forms.Button();
            this.lbDownloadFile = new System.Windows.Forms.Label();
            this.dgvExcel = new System.Windows.Forms.DataGridView();
            this.lblWord = new System.Windows.Forms.Label();
            this.btnWord = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnWriteOnExcelFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lblMonth = new System.Windows.Forms.Label();
            this.lblGroup = new System.Windows.Forms.Label();
            this.dgvPropusk = new System.Windows.Forms.DataGridView();
            this.btnLoadFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.progressPropusk = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcel)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPropusk)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDownloadFile
            // 
            this.btnDownloadFile.Location = new System.Drawing.Point(143, 8);
            this.btnDownloadFile.Name = "btnDownloadFile";
            this.btnDownloadFile.Size = new System.Drawing.Size(98, 23);
            this.btnDownloadFile.TabIndex = 0;
            this.btnDownloadFile.Text = "Загрузить";
            this.btnDownloadFile.UseVisualStyleBackColor = true;
            this.btnDownloadFile.Click += new System.EventHandler(this.btnDownloadFile_Click);
            // 
            // lbDownloadFile
            // 
            this.lbDownloadFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbDownloadFile.AutoSize = true;
            this.lbDownloadFile.Location = new System.Drawing.Point(6, 12);
            this.lbDownloadFile.Name = "lbDownloadFile";
            this.lbDownloadFile.Size = new System.Drawing.Size(129, 15);
            this.lbDownloadFile.TabIndex = 1;
            this.lbDownloadFile.Text = "Загрузить EXCEL файл";
            // 
            // dgvExcel
            // 
            this.dgvExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvExcel.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgvExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvExcel.GridColor = System.Drawing.SystemColors.ControlDarkDark;
            this.dgvExcel.Location = new System.Drawing.Point(392, 6);
            this.dgvExcel.Name = "dgvExcel";
            this.dgvExcel.RowTemplate.Height = 25;
            this.dgvExcel.Size = new System.Drawing.Size(852, 573);
            this.dgvExcel.TabIndex = 2;
            // 
            // lblWord
            // 
            this.lblWord.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblWord.AutoSize = true;
            this.lblWord.Location = new System.Drawing.Point(6, 46);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(154, 15);
            this.lblWord.TabIndex = 3;
            this.lblWord.Text = "Сгенерировать word-файл";
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(166, 42);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(75, 23);
            this.btnWord.TabIndex = 4;
            this.btnWord.Text = "Генерация";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(6, 558);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(380, 21);
            this.progressBar.TabIndex = 5;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1258, 613);
            this.tabControl1.TabIndex = 6;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.lbDownloadFile);
            this.tabPage1.Controls.Add(this.dgvExcel);
            this.tabPage1.Controls.Add(this.progressBar);
            this.tabPage1.Controls.Add(this.btnDownloadFile);
            this.tabPage1.Controls.Add(this.btnWord);
            this.tabPage1.Controls.Add(this.lblWord);
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1250, 585);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Студенческий билет";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage2.Controls.Add(this.progressPropusk);
            this.tabPage2.Controls.Add(this.btnWriteOnExcelFile);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.lblMonth);
            this.tabPage2.Controls.Add(this.lblGroup);
            this.tabPage2.Controls.Add(this.dgvPropusk);
            this.tabPage2.Controls.Add(this.btnLoadFile);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1250, 585);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Пропуски";
            // 
            // btnWriteOnExcelFile
            // 
            this.btnWriteOnExcelFile.Location = new System.Drawing.Point(140, 49);
            this.btnWriteOnExcelFile.Name = "btnWriteOnExcelFile";
            this.btnWriteOnExcelFile.Size = new System.Drawing.Size(75, 23);
            this.btnWriteOnExcelFile.TabIndex = 6;
            this.btnWriteOnExcelFile.Text = "Записать";
            this.btnWriteOnExcelFile.UseVisualStyleBackColor = true;
            this.btnWriteOnExcelFile.Click += new System.EventHandler(this.btnWriteOnExcelFile_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 15);
            this.label2.TabIndex = 5;
            this.label2.Text = "Записать в Excel файл";
            // 
            // lblMonth
            // 
            this.lblMonth.AutoSize = true;
            this.lblMonth.Location = new System.Drawing.Point(529, 15);
            this.lblMonth.Name = "lblMonth";
            this.lblMonth.Size = new System.Drawing.Size(46, 15);
            this.lblMonth.TabIndex = 4;
            this.lblMonth.Text = "Месяц:";
            // 
            // lblGroup
            // 
            this.lblGroup.AutoSize = true;
            this.lblGroup.Location = new System.Drawing.Point(402, 15);
            this.lblGroup.Name = "lblGroup";
            this.lblGroup.Size = new System.Drawing.Size(49, 15);
            this.lblGroup.TabIndex = 3;
            this.lblGroup.Text = "Группа:";
            // 
            // dgvPropusk
            // 
            this.dgvPropusk.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgvPropusk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPropusk.Location = new System.Drawing.Point(402, 41);
            this.dgvPropusk.Name = "dgvPropusk";
            this.dgvPropusk.RowTemplate.Height = 25;
            this.dgvPropusk.Size = new System.Drawing.Size(842, 538);
            this.dgvPropusk.TabIndex = 2;
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.Location = new System.Drawing.Point(105, 11);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.Size = new System.Drawing.Size(75, 23);
            this.btnLoadFile.TabIndex = 1;
            this.btnLoadFile.Text = "Загрузить";
            this.btnLoadFile.UseVisualStyleBackColor = true;
            this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Загрузить файл";
            // 
            // progressPropusk
            // 
            this.progressPropusk.Location = new System.Drawing.Point(6, 556);
            this.progressPropusk.Name = "progressPropusk";
            this.progressPropusk.Size = new System.Drawing.Size(390, 23);
            this.progressPropusk.TabIndex = 7;
            // 
            // FormStart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1282, 637);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormStart";
            this.Text = "УМР-helper";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormStart_FormClosing);
            this.Load += new System.EventHandler(this.FormStart_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcel)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPropusk)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDownloadFile;
        private System.Windows.Forms.Label lbDownloadFile;
        private System.Windows.Forms.DataGridView dgvExcel;
        private System.Windows.Forms.Label lblWord;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnLoadFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgvPropusk;
        private System.Windows.Forms.Label lblMonth;
        private System.Windows.Forms.Label lblGroup;
        private System.Windows.Forms.Button btnWriteOnExcelFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressPropusk;
    }
}

