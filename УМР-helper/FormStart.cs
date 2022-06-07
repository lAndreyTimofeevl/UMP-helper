using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
        string groupp = null;
        string month = null;
        int findMonthJ = 0;

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

            label2.Visible = false;
            btnWriteOnExcelFile.Visible = false;
        }
        public void opendocument()
        {
            progressBar.Visible = true;
            Document oDoc;

            for (int i = 1; i < dgvExcel.RowCount-1; i++)
            {
                int j = 2;
                oDoc = GetDoc(Environment.CurrentDirectory + "\\Шаблон\\Шаблон.docx",
                    dgvExcel[j, i].Value.ToString(),
                    dgvExcel[j + 1, i].Value.ToString(),
                    dgvExcel[j + 2, i].Value.ToString(),
                    dgvExcel[j + 3, i].Value.ToString(),
                    dgvExcel[j + 4, i].Value.ToString(),
                    dgvExcel[j + 5, i].Value.ToString());
                oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\Itu-119 " + dgvExcel[j, i].Value.ToString() + ".docx");
                oDoc.Close();
                Marshal.ReleaseComObject(oDoc);
                progressBar.Value ++;
                if(progressBar.Value == progressBar.Maximum)
                {
                    MessageBox.Show("Файл успешно сгенерирован!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    progressBar.Visible = false;
                    progressBar.Value = 0;
                    Marshal.ReleaseComObject(oDoc);
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

        private void FormStart_FormClosing(object sender, FormClosingEventArgs e)
        {
            Marshal.ReleaseComObject(oWord);
        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            label2.Visible = true;
            btnWriteOnExcelFile.Visible = true;
            OpenFileDialog excelFile = new OpenFileDialog();
            excelFile.DefaultExt = "*.xlsx;*xls";
            excelFile.Filter = "Excel(*.xlsx)|*.xlsx";
            excelFile.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet nwSheet;
            ExcelObj.Range shtRange;
            System.Data.DataTable dt = new System.Data.DataTable();
            if (excelFile.ShowDialog() == DialogResult.OK)
            {
                workbook = app.Workbooks.Open(excelFile.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);
                nwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                shtRange = nwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= shtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(new DataColumn((shtRange.Cells[2, Cnum] as ExcelObj.Range).Value2.ToString()));
                    groupp = (shtRange.Cells[1, 2] as ExcelObj.Range).Value2.ToString();
                    month = (shtRange.Cells[1, 4] as ExcelObj.Range).Value2.ToString();
                }
                dt.AcceptChanges();
                string[] columnNames = new string[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }
                for (int Rnum = 3; Rnum <= shtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= shtRange.Columns.Count; Cnum++)
                    {
                        if ((shtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] = (shtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                dgvPropusk.DataSource = dt;
                app.Quit();
                //Marshal.ReleaseComObject(workbook);
                //Marshal.ReleaseComObject(nwSheet);
                //Marshal.ReleaseComObject(shtRange);
            }
            app.Quit();
            oWord.Quit();
            //Marshal.ReleaseComObject(oWord);

            lblGroup.Text = "Группа: " + groupp;
            lblMonth.Text = "Месяц: " + month;
        }

        private void btnWriteOnExcelFile_Click(object sender, EventArgs e)
        {
            progressPropusk.Maximum = dgvPropusk.RowCount - 2;
            progressPropusk.Minimum = 0;
            ExcelObj.Application app1 = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet nwSheet;
            workbook = app1.Workbooks.Open(Environment.CurrentDirectory + "\\excel\\List.xlsx");
            try
            {
                nwSheet = workbook.Worksheets[groupp];
            }
            catch(Exception ex)
            {
                var newWS = (Worksheet)workbook.Sheets.Add(After: workbook.ActiveSheet);
                newWS.Name = groupp;
                nwSheet = newWS;
            }
            ExcelObj.Range cell1 = nwSheet.get_Range("A1", "A2").Cells;
            cell1.Merge(Type.Missing);
            nwSheet.Cells[1, 1] = "№";

            ExcelObj.Range cell2 = nwSheet.get_Range("B1", "B2").Cells;
            cell2.Merge(Type.Missing);
            nwSheet.Cells[1, 2] = "ФИО";

            ExcelObj.Range september = nwSheet.get_Range("C1", "E1").Cells;
            september.Merge(Type.Missing);
            nwSheet.Cells[1, 3] = "Сентябрь";

            ExcelObj.Range oktober = nwSheet.get_Range("F1", "H1").Cells;
            oktober.Merge(Type.Missing);
            nwSheet.Cells[1, 6] = "Октябрь";

            ExcelObj.Range november = nwSheet.get_Range("I1", "K1").Cells;
            november.Merge(Type.Missing);
            nwSheet.Cells[1, 9] = "Ноябрь";

            ExcelObj.Range december = nwSheet.get_Range("L1", "N1").Cells;
            december.Merge(Type.Missing);
            nwSheet.Cells[1, 12] = "Декабрь";

            ExcelObj.Range yanuar = nwSheet.get_Range("O1", "Q1").Cells;
            yanuar.Merge(Type.Missing);
            nwSheet.Cells[1, 15] = "Январь";

            ExcelObj.Range februar = nwSheet.get_Range("R1", "T1").Cells;
            februar.Merge(Type.Missing);
            nwSheet.Cells[1, 18] = "Февраль";

            ExcelObj.Range mart = nwSheet.get_Range("U1", "W1").Cells;
            mart.Merge(Type.Missing);
            nwSheet.Cells[1, 21] = "Март";

            ExcelObj.Range aprel = nwSheet.get_Range("X1", "Z1").Cells;
            aprel.Merge(Type.Missing);
            nwSheet.Cells[1, 24] = "Апрель";

            ExcelObj.Range may = nwSheet.get_Range("AA1", "AC1").Cells;
            may.Merge(Type.Missing);
            nwSheet.Cells[1, 27] = "Май";

            for (int i = 0; i < 27; i+=3)
            {
                nwSheet.Cells[2, i + 3] = "Всего";
                nwSheet.Cells[2, i + 4] = "Уваж.";
                nwSheet.Cells[2, i + 5] = "Неуваж.";
            }
            for(int i = 0; i < 29; i++)
            {
                if (nwSheet.Cells[1, i + 1].Value == month)
                {
                    findMonthJ = i;
                    break;
                }
                
            }
            for (int i = 0; i < dgvPropusk.Rows.Count; i++)
            {
                for (int j = 0; j < dgvPropusk.ColumnCount-3; j++)
                {
                    nwSheet.Cells[i + 3, j + 1] = dgvPropusk.Rows[i].Cells[j].Value;
                    
                }
                for (int j = 2; j < dgvPropusk.ColumnCount; j++)
                {
                    nwSheet.Cells[i + 3, j + findMonthJ - 1] = dgvPropusk.Rows[i].Cells[j].Value;

                }
                progressPropusk.Value++;
                if (progressPropusk.Value == progressPropusk.Maximum)
                {
                    MessageBox.Show("Данные успешно записаны!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    progressPropusk.Visible = false;
                    progressPropusk.Value = 0;
                }
            }
            //workbook.SaveAs();
            app1.Quit();
            //Marshal.ReleaseComObject(workbook);
            //Marshal.ReleaseComObject(nwSheet);
            //Marshal.ReleaseComObject(oWord);
        }
    }

}