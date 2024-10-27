using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Notepad0._1
{
    public partial class Form3 : Form
    {
        public string[,] data;
        public DataGridView dataGridView;
        public static Form3 instance;
        public string[,] data_table;
   
        public Form3()
        {
            InitializeComponent();
            InitializeDataGridView();
            instance = this;
            dataGridView = dataGridView1;
            
        }
        private void InitializeDataGridView()
        {
            
            dataGridView1.ColumnCount = 10;
            dataGridView1.RowCount = 2;
           
            dataGridView1.AllowUserToAddRows = true;
            
        }
        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            Excel.Application excelApp = new Excel.Application();

            if (excelApp != null)
            {
                
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);

                
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "UserData";

               

                
                for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++) 
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                
                Excel.Range range = worksheet.UsedRange;

                
                object[,] valueArray = (object[,])range.Value2;

                int rowCount = valueArray.GetLength(0); 
                int colCount = valueArray.GetLength(1); 

                
                data = new string[rowCount, colCount];

                
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        
                        data[i - 1, j - 1] = valueArray[i, j]?.ToString();
                    }
                }
                data_table = data;
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                this.DialogResult = DialogResult.OK;
                this.Close();



            }
        }
    }
}
