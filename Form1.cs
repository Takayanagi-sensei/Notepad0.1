using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Notepad0._1
{
    public partial class Form1 : Form
    {
        public static Form1 Instance;
        public RichTextBox MainRichTextBox;
        public Form1()
        {
            InitializeComponent();
            //To access this forms functions in other forms
            Instance = this;
            MainRichTextBox = richTextBox1;
        }
        //delete all
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }
        //opening a file
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Open your file";
            ofd.Filter = "Text Document(*.txt)|*.txt|All files(*.txt)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.LoadFile(ofd.FileName, RichTextBoxStreamType.PlainText);
            }
            this.Text = ofd.FileName;
        }
        //saving a file
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Do you want to save this file?";
            sfd.Filter = "Text Document(*.txt)|*.txt|All files(*.txt)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SaveFile(sfd.FileName, RichTextBoxStreamType.PlainText);
            }
            this.Text = sfd.FileName;
        }


        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Do you want to save this file?";
            sfd.Filter = "Text Document(*.txt)|*.txt|All files(*.txt)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SaveFile(sfd.FileName, RichTextBoxStreamType.PlainText);
            }
            this.Text = sfd.FileName;
        }

        // exiting the application
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        //reversing any mistake
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }
        //redoing any text deleted recently
        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo(); 

        }
        // Event handler for Cut action
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }
        // Event handler for copy action
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }
        // Event handler for Paste action
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();
        }
        // Event handler to insert current time and date
        private void timeDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime dt= DateTime.Now;
            richTextBox1.Text += dt.ToString(); 
        }

        private void fontsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fnt = new FontDialog();
            if (fnt.ShowDialog()==DialogResult.OK)
            {
                richTextBox1.Font = fnt.Font;
            }
        }

        private void colorToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
                ColorDialog clr = new ColorDialog();
                if (clr.ShowDialog() == DialogResult.OK)
                {
                    richTextBox1.SelectionColor = clr.Color;
                }
            
        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionLength > 0)
            {

                Font boldFont = new Font(richTextBox1.SelectionFont, FontStyle.Bold);

                richTextBox1.SelectionFont = boldFont;
            }
        }
        // Variables for managing zoom factor
        private float zoomFactor = 1.0f;
        private void zoomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            zoomFactor += 0.25f;

            richTextBox1.ZoomFactor = zoomFactor;
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This notepad is created by Avijit Ghosh","Creator of this software", MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void zoomOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (zoomFactor > 1f)
            {
                zoomFactor -= 0.25f;
                richTextBox1.ZoomFactor = zoomFactor;
            }
        }

        // Event handler for Find feature
        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string wordToFind = richTextBox1.Text;

            if (!string.IsNullOrWhiteSpace(wordToFind))
            {
                int index = richTextBox1.Find(wordToFind);

                if (index != -1)
                {
                    
                    richTextBox1.Select(index, wordToFind.Length);
                    richTextBox1.ScrollToCaret();
                }
                else
                {
                    MessageBox.Show("Word not found.", "Find");
                }
            }
            else
            {
                MessageBox.Show("Please enter a word to find.", "Find");
            }
        }
        // Load event for Form1
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        // Event handler for toggling bullet points
        private void bulletToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionBullet = !richTextBox1.SelectionBullet;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void undoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void cutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void copyToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void deleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void selectAllToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();

        }

        //for searching using google using right click functionality
        private void searchWithGoogleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            string selectedText = richTextBox1.SelectedText;

            
            if (!string.IsNullOrEmpty(selectedText))
            {
               
                string googleSearchUrl = $"https://www.google.com/search?q={Uri.EscapeDataString(selectedText)}";

                Process.Start(new ProcessStartInfo
                {
                    FileName = googleSearchUrl,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Please select some text to search.", "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        //updating status information

        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            UpdateStatus();
        }
        //live update line and column status
        private void UpdateStatus()
        {
            int pos = richTextBox1.SelectionStart;
            int line = richTextBox1.GetLineFromCharIndex(pos)+1;
            int col = pos - richTextBox1.GetFirstCharIndexOfCurrentLine() + 1;
            status.Text = "Line -> " + line + " Column -> " + col;
        }

        private void viewToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        //template part

        //to do template
        private void toDoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(Environment.NewLine + "To Do:" + Environment.NewLine );
            richTextBox1.SelectionBullet = true;
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.SelectionBullet = false;
        }


        //Method to import data from an Excel file
        public string[,] ImportDataFromExcel(string filePath)
        {
           
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            string[,] data = new string[rowCount, colCount];

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    data[i - 1, j - 1] = Convert.ToString(range.Cells[i, j].Value2);
                }
            }

            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return data;
        }

        // Method to generate a table from imported data
        public void GenerateDashTable(string[,] data, RichTextBox richTextBox2)
        {
            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            int[] colWidths = new int[colCount];
            for (int col = 0; col < colCount; col++)
            {
                colWidths[col] = Enumerable.Range(0, rowCount)
                    .Select(row => data[row, col]?.Length ?? 0)
                    .Max();
            }

            StringBuilder tableBuilder = new StringBuilder();

            string separator = "+";
            for (int i = 0; i < colWidths.Length; i++)
            {
                separator += new string('-', colWidths[i] + 2) + "+";
            }
            tableBuilder.AppendLine(separator);

            for (int row = 0; row < rowCount; row++)
            {
                string rowLine = "|";
                for (int col = 0; col < colCount; col++)
                {
                    string cellData = data[row, col] ?? string.Empty;
                    rowLine += " " + cellData.PadRight(colWidths[col]) + " |";
                }
                tableBuilder.AppendLine(rowLine);
                tableBuilder.AppendLine(separator);
            }

            richTextBox2.Text += tableBuilder.ToString();
        }

        // Method for importing from an Excel file via Form3
        public void import()
        {
            Form3 form3 = new Form3();
            if (form3.ShowDialog() == DialogResult.OK)
            {
                

               
                GenerateDashTable(Form3.instance.data_table, richTextBox1);
                
            }
            
        }
        //import method for directly selecting an Excel file
        public void import_1()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Workbook|*.xlsx;*.xls";
            openFileDialog.Title = "Select an Excel File";

    
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                
                string filePath = openFileDialog.FileName;

              
                string[,] data = ImportDataFromExcel(filePath);


                GenerateDashTable(data, richTextBox1);
            }

        }
        private void tableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(Environment.NewLine);
            import();
             
        }
        private void importTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(Environment.NewLine);
            import_1();
        }
        private void dailyJournalToolStripMenuItem_Click(object sender, EventArgs e)
        {

            richTextBox1.AppendText(Environment.NewLine + "Daily Journal" + Environment.NewLine);
            richTextBox1.AppendText(                      "-------------" + Environment.NewLine);
            richTextBox1.AppendText("Date: " + Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("Describe the day:" + Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("Goals For Tomorrow: " + Environment.NewLine);
            richTextBox1.SelectionBullet = true;
            richTextBox1.AppendText("" + Environment.NewLine);
            richTextBox1.SelectionBullet = false;



        }
        public void HighlightText(string word)
        {
            richTextBox1.SelectAll();
            richTextBox1.SelectionBackColor = Color.White;
            richTextBox1.DeselectAll();

            
            int startIndex = 0;
            while (startIndex < richTextBox1.TextLength)
            {
                int wordStartIndex = richTextBox1.Find(word, startIndex, RichTextBoxFinds.None);
                if (wordStartIndex != -1)
                {
                    richTextBox1.SelectionStart += wordStartIndex;
                    richTextBox1.SelectionLength = word.Length;
                    richTextBox1.SelectionBackColor = Color.Yellow;
                    startIndex = wordStartIndex + word.Length;
                }
                else
                    break;
            }

        }
        public void normal()
        {
            richTextBox1.SelectAll();
            richTextBox1.SelectionBackColor = Color.White;
            richTextBox1.DeselectAll();
        }
            private void findToolStripMenuItem_Click_1(object sender, EventArgs e)
            {
            Highlighter findForm = new Highlighter();
            findForm.Show();
            }

        private void buttonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
        }
    }
}
