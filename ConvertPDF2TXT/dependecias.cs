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
using System.Data.OleDb;

namespace ConvertPDF2TXT
{
    public partial class dependecias : Form
    {
        public dependecias()
        {
            InitializeComponent();
        }

        private void Dependecias_Load(object sender, EventArgs e)
        {
            label2.Visible = false;
            label3.Visible = false;
            label5.Visible = false;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                textBox1.Text = "Analisando PDF";
                label2.Text = "Modulo do PDF Encontrado";
                PDDocument pdf = new PDDocument();
                progressBar1.Value = progressBar1.Value + 10;
                
                label2.Text = "Nenhum leitor PDF encontrado";
                string file = "new.pdf";
        
                System.Diagnostics.Process pdf_viewer = System.Diagnostics.Process.Start(file);
                progressBar1.Value = progressBar1.Value + 10;
                pdf_viewer.CloseMainWindow();
                pdf_viewer.Close();
                progressBar1.Value = progressBar1.Value + 10;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                label2.Show();
                label2.ForeColor = System.Drawing.Color.Red;
                progressBar1.ForeColor = System.Drawing.Color.Red;
            }
            finally
            {
                label2.Text = "PDF Encontrado";
                label2.ForeColor = System.Drawing.Color.Green;
                label2.Show();
                Cursor = Cursors.Default;
            }

            Cursor = Cursors.WaitCursor;
            try
            {
                textBox1.Text = "Analisando Conexão a DB";
                label3.Text = "Modulo de acesso a DB não encontrado";
                OleDbCommand cmd = new OleDbCommand();
                OleDbConnection cn = new OleDbConnection();
                OleDbConnection dr;
                progressBar1.Value = progressBar1.Value + 15;
                label3.Text = "Impossivel aceder a DB";
                cn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\duart\Documents\exemplo1.accdb";
                cmd.Connection = cn;
                cn.Open();
                string q = "Select * from Registos_Entradas";
                cmd.CommandText = q;
                cmd.ExecuteNonQuery();
                cn.Close();
                progressBar1.Value = progressBar1.Value + 25;
            }
            catch (Exception ex)
            {
                label3.Show();
                label3.ForeColor = System.Drawing.Color.Red;
                progressBar1.ForeColor = System.Drawing.Color.Red;
            }
            finally
            {
                label3.Text = "Acesso a DB bem sucedido";
                label3.ForeColor = System.Drawing.Color.Green;
                label3.Show();
                Cursor = Cursors.Default;
            }
            Cursor = Cursors.WaitCursor;

            try
            {
                textBox1.Text = "Analisando Excel";
                label5.Text = "Modulo Excel encontrado";
                Microsoft.Office.Interop.Excel.Application tabealxcel = new Microsoft.Office.Interop.Excel.Application();
                progressBar1.Value = progressBar1.Value + 20;
                label5.Text = "Nenhum leitor Excel encontrado";
                tabealxcel.Visible = true;
                tabealxcel.Visible = false;
                progressBar1.Value = progressBar1.Value + 10;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                label5.Show();
                label5.ForeColor = System.Drawing.Color.Red;
                progressBar1.ForeColor = System.Drawing.Color.Red;
            }
            finally
            {
                label5.Text = "Excel Encontrado";
                label5.ForeColor = System.Drawing.Color.Green;
                label5.Show();
                Cursor = Cursors.Default;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
            progressBar1.Value = 0;
            label2.Visible = false;
            label3.Visible = false;
            label5.Visible = false;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
