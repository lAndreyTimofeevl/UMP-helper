using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelObj = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace УМР_helper
{
    class Table
    {
        Word._Application oWord = new Word.Application();
        public void openTable(DataGridView table)
        {
            
            OpenFileDialog excelFile = new OpenFileDialog();
            excelFile.DefaultExt = "*.xlsx;*xls";
            excelFile.Filter = "Excel(*.xlsx)|*.xlsx";
            excelFile.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet nwSheet;
            ExcelObj.Range shtRange;
            System.Data.DataTable dt = new System.Data.DataTable();
            if(excelFile.ShowDialog() == DialogResult.OK)
            {
                workbook = app.Workbooks.Open(excelFile.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);
                nwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                shtRange = nwSheet.UsedRange;
                for(int Cnum = 1; Cnum <=shtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(new DataColumn((shtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();
                string[] columnNames = new string[dt.Columns.Count];
                for(int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }
                for(int Rnum = 2; Rnum <= shtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for(int Cnum = 1; Cnum <= shtRange.Columns.Count; Cnum++)
                    {
                        if((shtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] = (shtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                table.DataSource = dt;
                app.Quit();
            }
            app.Quit();
            oWord.Quit();


        }
        
    }
}