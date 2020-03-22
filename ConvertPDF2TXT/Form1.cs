﻿using System;
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

            this.MaximizeBox = false;
            this.MinimizeBox = false;
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
            dataGridView1.ReadOnly = false;
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


        private void Button2_Click(object sender, EventArgs e)
        {
            try {
                foreach (string item in listBox1.Items)
                {
                    PDDocument doc = PDDocument.load(item);
                    PDFTextStripper stripper = new PDFTextStripper();
                    richTextBox1.Text = (stripper.getText(doc));
                    string sPattern = "^\\d{6}$";
                    string error = "";
                    int contador = 0;
                    int contador_est = 0;
                    bool cod_ler;
                    string anterior = "Error";
                    string anterior_comp = "Error";
                    int index = 0;
                    int index2 = 0;
                    int matricula_str = 0;
                    int contador2 = 0;
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
                                if (s.Contains("-"))
                                {
                                    textBox4.Text = matricula.Substring(0, 8);
                                }
                                textBox6.Text = matricula.Substring(9, 10);
                                //string transportador = new string(linha.Substring(27, 50).Reverse().ToArray());
                                //index = transportador.IndexOf(",");
                                //textBox9.Text = transportador.Substring(0, index);
                                string transportador = s.Substring(s.IndexOf(" ") + 1);
                                transportador = transportador.Substring(transportador.IndexOf(" ")+1, transportador.IndexOf("-")-11);
                                textBox9.Text = transportador;

                            }

                            if (s.Contains("N.º ORDEM NIF/NIPC ORGANIZAÇÃO MATRÍCULA DATA INÍCIO TRANSPORTE HORA INÍCIO TRANSPORTE"))
                            {
                                matricula_str++;
                            }
                            if (s.Contains("ESTABELECIMENTO") && contador_est == 0)
                            {
                                string est = s.Remove(0, 16);
                                textBox7.Text = est;
                                contador_est++;
                            }

                            if (s.Contains("CÓDIGO DOCUMENTO"))
                            {
                                textBox5.Text = s.Substring(17, 16);
                            }
                            if (contador2 > 0)
                            {
                                if (s.Contains("R"))
                                {
                                    if (s.Contains(" - "))
                                    {
                                        index = s.IndexOf("-");
                                        textBox8.Text = s.Substring(0, index);
                                        break;
                                    }
                                }
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
                                        try
                                        {
                                            textBox3.Text = anterior.Substring(anterior.IndexOf(")") + 2, anterior.IndexOf(",") + 1);
                                        }
                                        catch
                                        {
                                            textBox3.Text = anterior.Substring(0, anterior_comp.IndexOf(",") + 1);
                                        }
                                        
                                    }
                                    catch (Exception ex)
                                    {
                                        try
                                        {
                                            index = s.IndexOf(",");
                                            string teste_str = s.Substring(0, index);
                                            float teste = float.Parse(teste_str);
                                            index2 = s.IndexOf(",");
                                            try
                                            {
                                                anterior_comp = s;
                                                anterior = s.Substring(s.IndexOf(")"));
                                                index2 = s.IndexOf(")");
                                            }
                                            catch {
                                                anterior = s;
                                            }
                                            
                                            contador2++;
                                        }
                                        catch { }

                                    }
                                }
                            }
                        }
                        catch (Exception eg)
                        {
                            //MessageBox.Show("Error", eg.ToString());
                        }
                    }
                    if (textBox2.Text == "Error")
                    {
                        error = error + " Cod. Ler";
                    }

                    if (textBox5.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Numero da guia";
                    }

                    if (textBox3.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Quantidade";
                    }

                    if (textBox6.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Data de entrada";
                    }

                    if (textBox4.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Matricula";
                    }

                    if (textBox8.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Operação";
                    }

                    if (textBox7.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Estabelecimento";
                    }

                    if (textBox9.Text == "Error")
                    {
                        if (error != "")
                        {
                            error = error + ",";
                        }
                        error = error + " Transportador";
                    }

                    if (error != "")
                    {
                        var res = MessageBox.Show("Error ao ler:  " + error + " \nDocumenro afetado: " + item + "\n Deseja abrir o documento e confirmar?", "Error ao ler", MessageBoxButtons.YesNoCancel);
                        if (res == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(item);
                            var res2 = MessageBox.Show("Deseja adicionar a lista o documento?", "Cod. Ler", MessageBoxButtons.YesNo);
                            if (res2 == DialogResult.Yes)
                            {
                                string[] row = new string[] { textBox2.Text, textBox5.Text, textBox3.Text, textBox6.Text, textBox4.Text, textBox8.Text, textBox7.Text, textBox9.Text };
                                dataGridView1.Rows.Add(row);
                            }

                        }
                        else if (res == DialogResult.No)
                        {
                            string[] row = new string[] { textBox2.Text, textBox5.Text, textBox3.Text, textBox6.Text, textBox4.Text, textBox8.Text, textBox7.Text, textBox9.Text };
                            dataGridView1.Rows.Add(row);
                        }
                    }
                    else
                    {
                        string[] row = new string[] { textBox2.Text, textBox5.Text, textBox3.Text, textBox6.Text, textBox4.Text, textBox8.Text, textBox7.Text, textBox9.Text };
                        dataGridView1.Rows.Add(row);
                    }


                }
                listBox1.Items.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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

        private void AjudaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Abrir - Abre um navegador onde o utilizador pode escolher os documentos a serem analisados\nConverter - Adiciona a informação dos documentos selecionados a tabela \nSalvar - Salva a tabela num documento excel ", "Ajuda");
        }

        private void CréditosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Programador: Duarte Cruz \nIcons: https://www.flaticon.com/br/autores/iconixar \nCodigo Base: https://www.youtube.com/watch?v=neynvzrPTbs", "Créditos");
        }
    }
}


//Fonte de boa parte do códgio: https://www.youtube.com/watch?v=neynvzrPTbs
