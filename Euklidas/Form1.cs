using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using NPOI.HSSF.Model;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.Util;
using NPOI.SS.UserModel;

namespace Euklidas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        XSSFWorkbook wb;
        XSSFSheet sh;
        string fname = "";
        string[] rows_ = { "A", "B", "C", "D", "E", "F", "G", "H", "I" };
        Excelmodel[] excelmodels = new Excelmodel[10]; // Cia tiesiog.
        public ulong convertMe(string valueToConvert)
        {
            ulong convertedValue = 0;
            byte[] ascii = Encoding.ASCII.GetBytes(valueToConvert);
            foreach (Byte b in ascii)
            {
                convertedValue += ulong.Parse(b.ToString());
            }
            return convertedValue;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            duomenis();
        }

        public void duomenis()
        {

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
            }
            using (var fs = new FileStream(fname, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);


            }
            int i = 0;
            string sheetName = "Sheet1";
            sh = (XSSFSheet)wb.GetSheet(sheetName);
            while (sh.GetRow(i) != null)
            {
                if (dataGridView1.Columns.Count < sh.GetRow(i).Cells.Count)
                {
                    for (int j = 0; j < sh.GetRow(i).Cells.Count; j++)
                    {
                        dataGridView1.Columns.Add("", "");
                    }
                }
                dataGridView1.Rows.Add();

              
                for (int j = 0; j < sh.GetRow(i).Cells.Count; j++)
                {
                    var cell = sh.GetRow(i).GetCell(j);

                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case NPOI.SS.UserModel.CellType.Numeric:
                                dataGridView1[j, i].Value = sh.GetRow(i).GetCell(j).NumericCellValue;
                                break;
                            case NPOI.SS.UserModel.CellType.String:
                                dataGridView1[j, i].Value = sh.GetRow(i).GetCell(j).StringCellValue;
                                break;
                        }
                    }
                }

                i++;
            }
        }
        string[] userValues = new string[9];
        string[] value;
        public string[] readData()
        {
            var results = new List<string>();

            for (var row = 1; row < dataGridView1.Rows.Count - 1; row++)
            {
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    results.Add(dataGridView1.Rows[row].Cells[i].Value.ToString());
                }
            }

            return results.ToArray();

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] secondRow = get_Sheets(1);
            excelmodels[0] = new Excelmodel();
            double[] convertedValues = new double[secondRow.Length];
            for(int i = 0; i < convertedValues.Length; i++)
            {
                convertedValues[i] = convertMe(secondRow[i]);
            }
            double[] tryMe = new double[] {147.0, 47.0, 125.0, 25.0, 142.0, 74.0, 94.0, 121.0, 47.0};
            Euklidas1 euklidas1 = new Euklidas1();
            double mysteryResult = euklidas1.EuklidoAtstumas(tryMe, convertedValues);
            MessageBox.Show(mysteryResult.ToString());           
        }
       
        
        public string[] get_Sheets(int colNum)
        {
            CellReference[] cr = new CellReference[rows_.Length];
            ISheet sheet = wb.GetSheetAt(0);
            IRow[] row = new IRow[rows_.Length];
            ICell[] cell = new ICell[rows_.Length];
            string[] cellResult = new string[rows_.Length];
            for (int i = 0; i < cr.Length; i++)
            {
                cr[i] = new CellReference(rows_[i] + colNum);
                row[i] = sheet.GetRow(cr[i].Row);
                cell[i] = row[i].GetCell(cr[i].Col);
                cellResult[i] = cell[i].ToString();
            }
            return cellResult;
        }
    }
}