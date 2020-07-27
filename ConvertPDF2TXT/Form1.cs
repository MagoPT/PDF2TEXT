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
using MySql.Data.MySqlClient;
using System.Data.OleDb;

namespace ConvertPDF2TXT
{
    public partial class Form1 : Form
    {
        OleDbCommand cmd = new OleDbCommand();
        OleDbConnection cn = new OleDbConnection();
        OleDbConnection dr;
        string db_loc = "";
        string provider_Db = "";
        string security_info = "Persist Security Info = ;";
        string db_pass = "Jet OLEDB:Database Password = ;";
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
            button4.Text = "Apagar";
            button5.Text = "Localizar DB";
            label1.Text = "Cd. LER:";
            label2.Text = "Peso:";
            label3.Text = "Matricula:";
            label4.Text = "Cd Doc:";
            label5.Text = "Data:";
            label6.Text = "Estabelecimento:";
            label7.Text = "Transportador:";
            label8.Text = "Operação:";
            dataGridView1.ReadOnly = false;
            label9.Visible = false;
            progressBar1.Visible = false;
        }

        public float quantidade_get(string mat, int caract)
        {
            float matricula_final;
            try
            {
                matricula_final = float.Parse(mat.Substring(0, caract));
            }
            catch
            {
                caract--;
                matricula_final = quantidade_get(mat, caract);
            }

            return matricula_final;   
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
            int tamanho = listBox1.Items.Count;
            int load_size = 0;
            if (tamanho != 0)
            {
                load_size = 100 / tamanho;
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                //label9.Visible = true;
            }
            try {
                for (int counter = 0; counter < tamanho; counter++)
                {
                    progressBar1.Value += load_size;
                    label9.Text = "A carregar "+(counter+1)+"de" +tamanho;
                    string item = listBox1.Items[0].ToString();
                    PDDocument doc = PDDocument.load(item);
                    PDFTextStripper stripper = new PDFTextStripper();
                    richTextBox1.Text = (stripper.getText(doc));
                    string sPattern = "^\\d{6}$";
                    string error = "";
                    int contador = 0;
                    int contador_est = 0;
                    bool cod_ler = true;
                    bool justone = false;
                    string anterior = "Error";
                    string anterior_comp = "Error";
                    int index = 0;
                    int index2 = 0;
                    int matricula_str = 0;
                    string transport_orig = "";
                    int contador2 = 0;
                    var old_Strings = new List<string>();
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
                                int leng_mat = 0;
                                string matricula = new string(linha.Substring(0, 25).Reverse().ToArray());
                               
                                if (s.Contains("-"))
                                {
                                    textBox4.Text = matricula.Substring(0, 8).ToUpper();
                                    leng_mat = textBox4.Text.Length;
                                    textBox4.Text = textBox4.Text.Replace("-","");
                                    textBox4.Text = textBox4.Text.Replace(" ", "");
                                    if (textBox4.Text.Length == 4)
                                    {
                                        String matricula_v2 = new string(linha.Substring(0, 29).Reverse().ToArray());
                                        leng_mat = textBox4.Text.Length;
                                        textBox4.Text = matricula.Substring(0, 11).ToUpper();
                                        textBox4.Text = textBox4.Text.Replace("-", "");
                                        textBox4.Text = textBox4.Text.Replace(" ", "");
                                    }
                                }else if (matricula.Substring(0, 8).Replace(" ","").Length==6)
                                {
                                    textBox4.Text = matricula.Substring(0, 8).ToUpper();
                                    leng_mat = textBox4.Text.Length;
                                    textBox4.Text = textBox4.Text.Replace("-", "");
                                    textBox4.Text = textBox4.Text.Replace(" ", "");
                                }
                                textBox6.Text = matricula.Substring(9, 10);

                                string data = textBox6.Text[8] + "" + textBox6.Text[9] + "/" + textBox6.Text[5] + "" + textBox6.Text[6] + "/" + textBox6.Text[0] + "" + textBox6.Text[1] + textBox6.Text[2] + "" + textBox6.Text[3];
                                textBox6.Text = data;

                                string transportador = s.Substring(s.IndexOf(" ") + 1);
                                try
                                {
                                    bool matr = false;
                                    
                                    transportador = transportador.Substring(transportador.IndexOf(" ") + 1);
                                    string matricula_test = new string(linha.Substring(0, 17+leng_mat).Reverse().ToArray());
                                    if (matricula_test[1] == ' ')
                                    {
                                        matricula_test = new string(linha.Substring(0, 16 + leng_mat).Reverse().ToArray());
                                    }
                                    try
                                    {
                                        matricula_test = matricula_test.Substring(matricula_test.IndexOf(" . "));
                                    }
                                    catch{}

                                    if (textBox4.Text == "Error")
                                    {
                                        string matricula2 = new string(s.Reverse().ToArray());
                                        matricula2 = matricula2.Substring(17);
                                        matricula2 = new string(matricula2.Reverse().ToArray());
                                        matr = true;
                                    }
                                   
                                    transportador = transportador.Replace(matricula_test, "");
                                    if (matr)
                                    {
                                        if (transportador.ToLower().Contains("lda")){
                                            string teste_rand = new string(transportador.Reverse().ToArray());

                                            teste_rand = teste_rand.Substring(0, teste_rand.IndexOf('a'));
                                            if (teste_rand.Contains('.'))
                                            {
                                                teste_rand = teste_rand.Substring(0, teste_rand.IndexOf('.'));
                                            }
                                            teste_rand = new string(teste_rand.Reverse().ToArray());
                                            teste_rand.Remove(' ');
                                            textBox4.Text = teste_rand;
                                        }
                                    }                                   
                                }
                                catch (Exception exep)
                                {
                                    transportador = transportador.Substring(transportador.IndexOf(" ") + 1, transportador.IndexOf(","));
                                    
                                    transportador = transportador.Replace(transportador.Substring(transportador.IndexOf(",")), "");
                                }
                                finally
                                {
                                    textBox9.Text = transportador;
                                    transport_orig = s;
                                }
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
                                if (s.Contains("R") && !s.Contains("Res"))
                                {
                                    if (s.Contains(" - "))
                                    {
                                        index = s.IndexOf("-");
                                        string operat = s.Substring(0, index);
                                        if (operat.Length < 5)
                                        {
                                            textBox8.Text = operat;
                                            break;
                                        }
                                        else
                                        {
                                            textBox2.Text = operat;
                                            contador2--;
                                        }
                                    }
                                }
                            }

                            if (contador == 0)
                            {
                                if (s.Contains("DADOS ORIGINAIS"))
                                {
                                    contador++;
                                }
                            }else{
                                contador++;
                                
                                if (contador >= 1) {
                                    old_Strings.Add(s);
                                    try
                                    {
                                        cod_ler = true;
                                        string teste_str = s.Substring(0, 6);
                                        int teste = int.Parse(teste_str);
                                        textBox2.Text = s.Substring(0, 6);
                                        try
                                        {
                                            textBox3.Text = anterior.Substring(anterior.IndexOf(")") + 1, anterior.IndexOf(",") + 2);
                                            textBox3.Text=textBox3.Text.Replace(" ","");
                                        }
                                        catch
                                        {
                                            textBox3.Text = anterior.Substring(0, anterior_comp.IndexOf(","));
                                            textBox3.Text = textBox3.Text.Replace(" ", "");
                                        }                                      
                                    }
                                    catch (Exception ex)
                                    {
                                        try
                                        {
                                            index = s.IndexOf(",");
                                            if(cod_ler && !justone)
                                            {
                                                anterior_comp = s;
                                                anterior = s.Substring(s.IndexOf(")")+1);
                                                anterior = anterior.Substring(0, anterior.IndexOf("("));
                                                cod_ler = false;
                                                justone = true;
                                                textBox3.Text = anterior;
                                            }
                                            string teste_str = s.Substring(0, index);
                                            float teste = float.Parse(teste_str);

                                            contador2++;
                                        }
                                        catch(Exception expetion) { cod_ler = true; }
                                    }
                                }
                            }
                        }
                        catch (Exception eg)
                        {
                            //MessageBox.Show("Error", eg.ToString());
                        }
                    }
                    if(textBox4.Text == "Error" || textBox8.Text=="Error")
                    {
                        if (textBox8.Text == "Error")
                        {
                            for (int i = 0; i < 4; i++)
                            {
                                switch (i)
                                {
                                    case 1:
                                        try
                                        {
                                            textBox3.Text = quantidade_get(old_Strings[i], 10) + "";
                                        }
                                        catch { }
                                        break;
                                    case 3:
                                        try
                                        {
                                            textBox8.Text = old_Strings[i].Substring(0, old_Strings[i].IndexOf('-'));
                                        }
                                        catch { }
                                        break;
                                }
                            }
                        }

                        try
                        {
                            string reverse = new string(transport_orig.Reverse().ToArray());
                            string text_remove = reverse.Substring(0, 16);
                            reverse = reverse.Replace(text_remove, "");
                            string orig = new string(reverse.Reverse().ToArray());
                            text_remove = orig.Substring(0, 12);
                            orig = orig.Replace(text_remove, "");
                            string matricula_final = "";
                            string transportador_final = "";
                            string orig_rev = new string(orig.Reverse().ToArray());
                            matricula_final = new string(orig_rev.Substring(0, orig_rev.IndexOf(' ')).Reverse().ToArray());
                            if(matricula_final.Equals(""))
                            {
                                orig_rev = orig_rev.Substring(1, orig_rev.Length-1);
                                matricula_final = new string(orig_rev.Substring(0, orig_rev.IndexOf(' ')).Reverse().ToArray());
                            }
                            transportador_final = orig.Replace(matricula_final, "");
                            transportador_final = new string(transportador_final.Reverse().ToArray());
                            string check_last_mat = transportador_final.ToUpper().Substring(0, 8);
                            
                            if(check_last_mat.Contains("ADL") || check_last_mat.Contains(".ADL") || check_last_mat.Contains("AS") || check_last_mat.Contains(".A.S"))
                            {
                                //TODO se precisarem de gerir as extensões
                            }
                            else
                            {
                                transportador_final = transportador_final.Replace(check_last_mat, "");
                                matricula_final = new string(check_last_mat.Reverse().ToArray()) + matricula_final;
                            }

                            //MessageBox.Show(check_last_mat);
                            transportador_final = new string(transportador_final.Reverse().ToArray());
                            textBox4.Text = matricula_final;
                            textBox9.Text = transportador_final;
                        }
                        catch(Exception ai_a_minha_vida) {
                            MessageBox.Show(ai_a_minha_vida+"");
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
                                string[] row = new string[] { textBox7.Text, textBox2.Text, textBox3.Text, textBox6.Text, textBox5.Text, textBox4.Text, textBox9.Text, textBox8.Text, item };
                                dataGridView1.Rows.Add(row);
                            }
                        }
                        else if (res == DialogResult.No)
                        {
                            string[] row = new string[] { textBox7.Text, textBox2.Text, textBox3.Text, textBox6.Text, textBox5.Text, textBox4.Text, textBox9.Text, textBox8.Text, item };
                            dataGridView1.Rows.Add(row);
                        }
                        listBox1.Items.RemoveAt(0);
                    }
                    else
                    {
                        string[] row = new string[] { textBox7.Text, textBox2.Text, textBox3.Text, textBox6.Text, textBox5.Text, textBox4.Text, textBox9.Text, textBox8.Text, item };
                        dataGridView1.Rows.Add(row);
                        listBox1.Items.RemoveAt(0);
                    }
                }
                listBox1.Items.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            label9.Visible = false;
            progressBar1.Visible = false;
        }
    
        private void Button3_Click(object sender, EventArgs e)
        {
            String erros_log = "";
            int erros_num = 0;
            string path_error="";
            try
            {
                cn.Close();
                cn.ConnectionString = provider_Db + db_loc;
                cn.Open();
                cmd.Connection = cn;
                if (dataGridView1.Rows.Count > 0)
                {
                    path_error = dataGridView1.Rows[0].Cells[8].Value+"";
                    path_error = path_error.Replace(dataGridView1.Rows[0].Cells[4].Value + ".pdf", "");
                    string path_correcto = path_error + "Corretas\\";
                    path_error += "Erradas\\";
                    for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
                    {
                        try
                        {
                            string q = "insert into Registos_Entradas(Empresa, Cod_Produto,Quantidade,Data_entrada,Num_Guia,Matricula,Transportador)VALUES('" + dataGridView1.Rows[i].Cells[0].Value + "','" + dataGridView1.Rows[i].Cells[1].Value + "','" + dataGridView1.Rows[i].Cells[2].Value + "','" + dataGridView1.Rows[i].Cells[3].Value + "','" + dataGridView1.Rows[i].Cells[4].Value + "','" + dataGridView1.Rows[i].Cells[5].Value + "','" + dataGridView1.Rows[i].Cells[6].Value + "')";
                            cmd.CommandText = q;
                            cmd.ExecuteNonQuery();
                            string fileName = dataGridView1.Rows[i].Cells[4].Value + "";
                            System.IO.Directory.CreateDirectory(path_correcto);
                            string sourceFile = System.IO.Path.Combine(dataGridView1.Rows[i].Cells[8].Value + "", "");
                            string destFile = System.IO.Path.Combine(path_correcto, fileName + ".pdf");
                            System.IO.File.Copy(sourceFile, destFile, true);
                            System.IO.File.Delete(sourceFile);
                        }
                        catch (System.Data.OleDb.OleDbException error)
                        {
                            string fileName = dataGridView1.Rows[i].Cells[4].Value + "";
                            erros_log += fileName + ": " + error+"\n";
                            erros_num++;
                            System.IO.Directory.CreateDirectory(path_error);
                            string sourceFile = System.IO.Path.Combine(dataGridView1.Rows[i].Cells[8].Value+"","");
                            string destFile = System.IO.Path.Combine(path_error, fileName+".pdf");
                            System.IO.File.Copy(sourceFile, destFile, true);
                            System.IO.File.Delete(sourceFile);
                        }
                    }
                }

                cn.Close();
                DateTime date = DateTime.Now;
                string timestamp = date.Hour +"_"+ date.Minute+"_" + date.Second;
                string erros_relat = path_error + "\\log_" + timestamp + "_error.txt";
                switch (erros_num)
                {                     
                    case 0:
                        MessageBox.Show("Todas os dados foram bem inseridos na BD","Sucesso");
                        break;
                    case 1:
                        MessageBox.Show(" 1 guia não foi inserida corretamente. \nUm relátorio completo pode ser encontrado em: " + erros_relat, "Erro DB");
                        System.IO.File.WriteAllText(erros_relat, erros_log);
                        break;
                    default:
                        MessageBox.Show(erros_num + " guias não foram inseridas corretamente. \nUm relátorio completo pode ser encontrado em: " + erros_relat, "Erro DB");
                        System.IO.File.WriteAllText(erros_relat, erros_log);
                        break;
                }
                dataGridView1.Rows.Clear();
            }
            catch(Exception ex)
            {
                if(db_loc == "")
                {
                    MessageBox.Show("Selecione a BD","BD erro");
                }
                else
                {
                    MessageBox.Show(ex+"","Erro");
                }
            } 
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {}

        private void AjudaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Abrir - Abre um navegador onde o utilizador pode escolher os documentos a serem analisados\nConverter - Adiciona a informação dos documentos selecionados a tabela \nSalvar - Salva a tabela num documento excel \nApagar - Apaga a tabela atual ", "Ajuda");
        }

        private void CréditosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Programador: Duarte Cruz \nIcons: https://www.flaticon.com/br/autores/iconixar \nCodigo Base: https://www.youtube.com/watch?v=neynvzrPTbs", "Créditos");
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("Deseja apagar a tabela atual?","Apagar",MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                dataGridView1.Rows.Clear();
            }    
        }

        private void DependênciasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dependecias f2 = new dependecias();
            f2.Show();
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Ficheiros DB (*.accdb,*.mdb)|*.accdb;*.mdb";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                int dot = ofd.FileName.IndexOf('.');
                    string ext = ofd.FileName.Substring(dot+1);
                if(ext == "accdb")
                {
                    provider_Db = "Provider=Microsoft.ACE.OLEDB.12.0;";
                }else if(ext == "mdb")
                {
                    provider_Db = "Provider=Microsoft.JET.OLEDB.4.0;";
                }
                    db_loc = "Data Source ="+ ofd.FileName;
            }
        }
    }
}
//Fonte de boa parte do códgio: https://www.youtube.com/watch?v=neynvzrPTbs