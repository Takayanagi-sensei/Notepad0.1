using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Notepad0._1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

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

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo(); 

        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

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

        private float zoomFactor = 1.0f;
        private void zoomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            zoomFactor += 1f;

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
                zoomFactor -= 1f;
                richTextBox1.ZoomFactor = zoomFactor;
            }
        }

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

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

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

        private void searchWithGoogleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Get the selected text from the TextBox
            string selectedText = richTextBox1.SelectedText;

            // Check if text is selected
            if (!string.IsNullOrEmpty(selectedText))
            {
                // Format the Google search URL
                string googleSearchUrl = $"https://www.google.com/search?q={Uri.EscapeDataString(selectedText)}";

                // Open the default browser with the Google search
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

        

        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            UpdateStatus();
        }

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

        private void toDoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(Environment.NewLine + "Title" + Environment.NewLine );
            richTextBox1.SelectionBullet = true;
            richTextBox1.AppendText("First Task" + Environment.NewLine);
            richTextBox1.AppendText("Second Task" + Environment.NewLine);
            richTextBox1.AppendText("Third Task" + Environment.NewLine);
            richTextBox1.SelectionBullet = false;
        }

        private void dailyJournalToolStripMenuItem_Click(object sender, EventArgs e)
        {

            richTextBox1.AppendText(Environment.NewLine + "Daily Journal" + Environment.NewLine);
            richTextBox1.AppendText(                      "~~~~~~~~~~" + Environment.NewLine);
            richTextBox1.AppendText("Date: " + Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("Describe the day:" + Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("Goals For Tomorrow: " + Environment.NewLine);
            richTextBox1.SelectionBullet = true;
            richTextBox1.AppendText("" + Environment.NewLine);
            richTextBox1.SelectionBullet = false;



        }
    }
}
