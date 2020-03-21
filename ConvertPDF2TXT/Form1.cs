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
            label6.Text = "Estabelecimento:";
            label7.Text = "Transportador:";
            label8.Text = "Operação:";
            //textBox1.Enabled = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Ficheiros PDF (*.pdf)|*.pdf";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string item in ofd.FileNames)
                {
                    listBox1.Items.Add(item);
                }
            }

        }

        public string anterior;
        private void Button2_Click(object sender, EventArgs e)
        {
            foreach (string item in listBox1.Items)
            {
                PDDocument doc = PDDocument.load(item);
                PDFTextStripper stripper = new PDFTextStripper();
                richTextBox1.Text = (stripper.getText(doc));
                string sPattern = "^\\d{6}$";
                int contador = 0;
                int contador_est = 0;
                bool cod_ler;
                int matricula_str = 0;
                textBox2.Text = "Error";
                textBox3.Text = "Error";
                textBox4.Text = "Error";
                textBox5.Text = "Error";
                textBox6.Text = "Error";
                textBox7.Text = "Error";
                textBox8.Text = "Error";
                textBox9.Text = "Error";
                foreach (string s in richTextBox1.Lines)
                {

                    try
                    {
                        if (matricula_str == 1)
                        {
                            matricula_str++;
                            string linha = new string(s.Reverse().ToArray());
                            string matricula = new string(linha.Substring(0, 25).Reverse().ToArray());
                            textBox4.Text = matricula.Substring(0, 8);
                            textBox6.Text = matricula.Substring(9, 10);
                            string transportador = new string(linha.Substring(27, 50).Reverse().ToArray());
                            int index = transportador.IndexOf(",");
                            textBox9.Text = transportador.Substring(0, index);

                        }

                        if (s.Contains("N.º ORDEM NIF/NIPC ORGANIZAÇÃO MATRÍCULA DATA INÍCIO TRANSPORTE HORA INÍCIO TRANSPORTE"))
                        {
                            matricula_str++;
                        }
                        if (s.Contains("ESTABELECIMENTO") && contador_est==0)
                        {
                            string est = s.Remove(0,16);
                            textBox7.Text = est;
                            contador_est++;
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
                string[] row = new string[] { textBox2.Text, textBox5.Text, textBox3.Text, textBox6.Text, textBox4.Text, textBox8.Text, textBox7.Text, textBox9.Text };
                dataGridView1.Rows.Add(row);
            }
            listBox1.Items.Clear();

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application tabealxcel = new Microsoft.Office.Interop.Excel.Application();
                tabealxcel.Application.Workbooks.Add(Type.Missing);
                for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                {
                    tabealxcel.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        tabealxcel.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                        //tabealxcel.Cells[i + 2, j + 1] = "batata";

                    }
                }
                tabealxcel.Columns.AutoFit();
                tabealxcel.Visible = true;
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}


//Fonte de boa parte do códgio: https://www.youtube.com/watch?v=neynvzrPTbs
