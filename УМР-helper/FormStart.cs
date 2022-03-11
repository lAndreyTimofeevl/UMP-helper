
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelObj = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace УМР_helper
{
    public partial class FormStart : Form
    {
        Word._Application oWord = new Word.Application();
        Table table = new Table();
        
        public FormStart()
        {
            InitializeComponent();
        }

        private void btnDownloadFile_Click(object sender, EventArgs e)
        {
            table.openTable(dgvExcel);
            lblWord.Visible = true;
            btnWord.Visible = true;
            progressBar.Maximum = dgvExcel.RowCount - 2;
            progressBar.Minimum = 0;
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            
            opendocument();
        }

        private void FormStart_Load(object sender, EventArgs e)
        {
            lblWord.Visible = false;
            btnWord.Visible = false;
            progressBar.Visible = false;
        }
        public void opendocument()
        {
            progressBar.Visible = true;
            Document oDoc;
            Document tmp;

            for (int i = 1; i < dgvExcel.RowCount-1; i++)
            {
                tmp = oWord.Documents.Add(Environment.CurrentDirectory + "\\Итог\\Итог.docx");
                int j = 2;
                oDoc = GetDoc(Environment.CurrentDirectory + "\\Шаблон\\Шаблон.docx",
                    dgvExcel[j, i].Value.ToString(),
                    dgvExcel[j + 1, i].Value.ToString(),
                    dgvExcel[j + 2, i].Value.ToString(),
                    dgvExcel[j + 3, i].Value.ToString(),
                    dgvExcel[j + 4, i].Value.ToString(),
                    dgvExcel[j + 5, i].Value.ToString());
                oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\temp.docx");
                oDoc.Close();
                //oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\Itu-119 " + dgvExcel[j, i].Value.ToString() + ".docx");
                //oWord.MergeDocuments(oDoc, tmp);
                tmp.Merge(Environment.CurrentDirectory + "\\temp.docx", tmp);
                //tmp.Merge(Environment.CurrentDirectory + "\\temp.docx", tmp);
                //oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\temp.docx");
                //tmp.SaveAs(FileName: Environment.CurrentDirectory + "\\Итог\\Итог.docx");
                tmp.Close();
                progressBar.Value ++;
                if(progressBar.Value == progressBar.Maximum)
                {
                    MessageBox.Show("Файл успешно сгенерирован!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    progressBar.Visible = false;
                    progressBar.Value = 0;
                }
            }
            

        }
        private Document GetDoc(string path, string midname, string fname, string lname, string date, string number, string prikaz)
        {
            Document oDoc = oWord.Documents.Add(path);
            SetTemplate(oDoc, midname, fname, lname, date, number, prikaz);
            return oDoc;
        }
        // Замена закладок на данные из dgv
        private void SetTemplate(Word.Document oDoc, string midname, string fname, string lname, string number, string prikaz, string date)
        {
            oDoc.Bookmarks["midname"].Range.Text = midname;
            oDoc.Bookmarks["fname"].Range.Text = fname;
            oDoc.Bookmarks["lname"].Range.Text = lname;
            oDoc.Bookmarks["dateprikaz"].Range.Text = date;
            oDoc.Bookmarks["formob"].Range.Text = "очная";
            oDoc.Bookmarks["number_bilet"].Range.Text = number;
            oDoc.Bookmarks["numpri"].Range.Text = prikaz;
        }
        
    }

}

