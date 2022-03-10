
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
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            
            opendocument();
        }

        private void FormStart_Load(object sender, EventArgs e)
        {
            lblWord.Visible = false;
            btnWord.Visible = false;
        }
        public void opendocument()
        {
            _Document oDoc;
            for (int i = 1; i < dgvExcel.RowCount-1; i++)
            {
                int j = 2;
                oDoc = GetDoc(Environment.CurrentDirectory + "\\Шаблон.docx",
                    dgvExcel[j, i].Value.ToString(),
                    dgvExcel[j + 1, i].Value.ToString(),
                    dgvExcel[j + 2, i].Value.ToString(),
                    dgvExcel[j + 3, i].Value.ToString(),
                    dgvExcel[j + 4, i].Value.ToString(),
                    dgvExcel[j + 5, i].Value.ToString());
                oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\Itu-119" + dgvExcel[j, i].Value.ToString() + ".docx");
                oDoc.Close();
                
            }
            MessageBox.Show("Файл успешно сгенерирован!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }
        private _Document GetDoc(string path, string midname, string fname, string lname, string date, string number, string prikaz)
        {
            _Document oDoc = oWord.Documents.Add(path);
            SetTemplate(oDoc, midname, fname, lname, date, number, prikaz);
            return oDoc;
        }
        private _Document PlusDoc(string[] mas, string path)
        {
            _Document doc = oWord.Documents.Add(path);

            return doc;
        }
        // Замена закладок на данные из dgv
        private void SetTemplate(Word._Document oDoc, string midname, string fname, string lname, string number, string prikaz, string date)
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