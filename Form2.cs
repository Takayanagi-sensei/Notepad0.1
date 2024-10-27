using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.Drawing;

namespace Notepad0._1
{
    public partial class Highlighter : Form
    {
        public static Highlighter Instance;
        public Highlighter()
        {
            InitializeComponent();
            Instance = this;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;

            this.Size = new Size(350, 115);
        }
        public String str;
        
        private string GenerateSeparator(int[] colWidths)
        {
            string separator = "+";
            foreach (var width in colWidths)
            {
                separator += new string('-', width + 2) + "+";
            }
            return separator;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            
            string wordToFind = textBox1.Text;
            if (!string.IsNullOrEmpty(wordToFind))
            {
                Form1.Instance.HighlightText(wordToFind);
            }

            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
                Form1.Instance.normal();
            

            
        }
    }
}
