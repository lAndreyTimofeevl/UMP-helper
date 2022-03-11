
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
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcel)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDownloadFile
            // 
            this.btnDownloadFile.Location = new System.Drawing.Point(149, 12);
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
            this.lbDownloadFile.Location = new System.Drawing.Point(12, 16);
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
            this.dgvExcel.Location = new System.Drawing.Point(453, 16);
            this.dgvExcel.Name = "dgvExcel";
            this.dgvExcel.RowTemplate.Height = 25;
            this.dgvExcel.Size = new System.Drawing.Size(817, 609);
            this.dgvExcel.TabIndex = 2;
            // 
            // lblWord
            // 
            this.lblWord.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblWord.AutoSize = true;
            this.lblWord.Location = new System.Drawing.Point(12, 50);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(154, 15);
            this.lblWord.TabIndex = 3;
            this.lblWord.Text = "Сгенерировать word-файл";
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(172, 46);
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
            this.progressBar.Location = new System.Drawing.Point(12, 602);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(435, 23);
            this.progressBar.TabIndex = 5;
            // 
            // FormStart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1282, 637);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.lblWord);
            this.Controls.Add(this.dgvExcel);
            this.Controls.Add(this.lbDownloadFile);
            this.Controls.Add(this.btnDownloadFile);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormStart";
            this.Text = "УМР-helper";
            this.Load += new System.EventHandler(this.FormStart_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcel)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDownloadFile;
        private System.Windows.Forms.Label lbDownloadFile;
        private System.Windows.Forms.DataGridView dgvExcel;
        private System.Windows.Forms.Label lblWord;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

