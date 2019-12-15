using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;

namespace ConvertPDF2TXT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "Conversor";
            button1.Text = "Abrir PDF";
            button2.Text = "Converter";
            button3.Text = "Salvar";
            textBox1.Enabled = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Ficheiros PDF (*.pdf)|*.pdf";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string path = ofd.FileName.ToString();
                textBox1.Text = path;
            }
          
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            PDDocument doc = PDDocument.load(textBox1.Text);
            PDFTextStripper stripper = new PDFTextStripper();
            richTextBox1.Text = (stripper.getText(doc));
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.DefaultExt = "*.txt";
            saveFileDialog1.Filter = "Ficheiro TXT|*.txt";
            if(saveFileDialog1.ShowDialog()==System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0)
            {
                richTextBox1.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.PlainText);
            }
        }
    }
}


//Fonte de boa parte do códgio: https://www.youtube.com/watch?v=neynvzrPTbs
