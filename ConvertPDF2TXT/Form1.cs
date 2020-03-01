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
            label1.Text = "Cd. LER:";
            label2.Text = "Peso:";
            label3.Text = "Matricula:";
            label4.Text = "Cd Doc:";
            label5.Text = "Data:";
            //textBox1.Enabled = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Ficheiros PDF (*.pdf)|*.pdf";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach(string item in ofd.FileNames)
                {
                    listBox1.Items.Add(item);
                }
            }
          
        }

        public string anterior;
        private void Button2_Click(object sender, EventArgs e)
        {
            foreach(string item in listBox1.Items)
            {
                PDDocument doc = PDDocument.load(item);
                PDFTextStripper stripper = new PDFTextStripper();
                richTextBox1.Text = (stripper.getText(doc));
                string sPattern = "^\\d{6}$";
                int contador = 0;
                bool cod_ler;
                int matricula_str = 0;
                textBox2.Text = "Error";
                textBox3.Text = "Error";
                textBox4.Text = "Error";
                textBox5.Text = "Error";
                textBox6.Text = "Error";
                foreach (string s in richTextBox1.Lines)
                {

                    try
                    {
                        if (matricula_str == 1)
                        {
                            matricula_str++;
                            string matricula = new string(s.Reverse().ToArray());
                            matricula = new string(matricula.Substring(0, 25).Reverse().ToArray());
                            textBox4.Text = matricula.Substring(0, 8);
                            textBox6.Text = matricula.Substring(9, 10);

                        }

                        if (s.Contains("N.º ORDEM NIF/NIPC ORGANIZAÇÃO MATRÍCULA DATA INÍCIO TRANSPORTE HORA INÍCIO TRANSPORTE"))
                        {
                            matricula_str++;
                        }

                        if (s.Contains("CÓDIGO DOCUMENTO"))
                        {
                            textBox5.Text = s.Substring(17, 16);
                        }

                        if (contador == 0)
                        {
                            if (s.Contains("DADOS ORIGINAIS"))
                            {
                                contador++;
                            }
                        }
                        else
                        {
                            contador++;
                            if (contador >= 1)
                            {

                                try
                                {
                                    Console.WriteLine(s);
                                    cod_ler = true;
                                    string teste_str = s.Substring(0, 6);
                                    int teste = int.Parse(teste_str);
                                    textBox2.Text = s.Substring(0, 6);
                                    textBox3.Text = anterior.Substring(0, 7);
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        string teste_str = s.Substring(0, 7);
                                        float teste = float.Parse(teste_str);
                                        anterior = s;
                                    }
                                    catch { }

                                }
                            }
                        }
                    }
                    catch (Exception eg)
                    {
                        MessageBox.Show("Error", eg.ToString());
                    }
                }
                string[] row = new string[] { textBox2.Text, textBox5.Text, textBox3.Text, textBox6.Text, textBox4.Text };
                dataGridView1.Rows.Add(row);
            }
            listBox1.Items.Clear();
            
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

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}


//Fonte de boa parte do códgio: https://www.youtube.com/watch?v=neynvzrPTbs
