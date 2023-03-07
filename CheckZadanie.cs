using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Polohov
{
    public partial class CheckZadanie : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        public CheckZadanie(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;
      
            
            InitializeComponent();

            //textBox1.Text = " ";

            listView1.Columns.Add("Дата", 120);
            listView1.Columns.Add("Смена", 60);
            listView1.Columns.Add("Станок", 100);
            listView1.Columns.Add("Рабочий", 100);            
            listView1.Columns.Add("Партия", 100);
            listView1.Columns.Add("Продукт", 120);
            listView1.Columns.Add("Вес", 50);       
            
            //listView1.Columns.Add("id", 0);

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;

            comboBox1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;

            comboBox2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;

            comboBox3.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;

            comboBox4.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;

            comboBox5.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;

            dateTimePicker1.Enabled = false;
            button11.Enabled = false;
            button12.Enabled = false;

            dateTimePicker2.Enabled = false;

            //заполняем комбобоксы
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select name from partiya where show!=0 and name!='Не определено'";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox1.Items.Add(r[0].ToString());
                }
            }
            conn.Close();

            conn.Open();            
            command.CommandText = "select name from prodykt";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox2.Items.Add(r[0].ToString());
                }
            }
            conn.Close();

            conn.Open();
            command.CommandText = "select name from stanok";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox4.Items.Add(r[0].ToString());
                }
            }
            conn.Close();

            conn.Open();
            command.CommandText = "select surname from rabotnik where working=1 and name!=''";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox3.Items.Add(r[0].ToString());
                }
            }
            conn.Close();

            comboBox5.Items.Add("День");
            comboBox5.Items.Add("Ночь");



        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                comboBox1.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = true;
            }
            if (!checkBox1.Checked)
            {
                comboBox1.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button2_Click(sender,e);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                comboBox2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
            }
            if (!checkBox2.Checked)
            {
                comboBox2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                button3_Click(sender, e);
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                comboBox3.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
            }
            if (!checkBox4.Checked)
            {
                comboBox3.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                button5_Click(sender, e);
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                comboBox4.Enabled = true;
                button7.Enabled = true;
                button8.Enabled = true;
            }
            if (!checkBox3.Checked)
            {
                comboBox4.Enabled = false;
                button7.Enabled = false;
                button8.Enabled = false;
                button7_Click(sender, e);
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                comboBox5.Enabled = true;
                button9.Enabled = true;
                button10.Enabled = true;
            }
            if (!checkBox5.Checked)
            {
                comboBox5.Enabled = false;
                button9.Enabled = false;
                button10.Enabled = false;
                button9_Click(sender, e);
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                dateTimePicker1.Enabled = true;
                button11.Enabled = true;
                button12.Enabled = true;

                if (checkBox7.Text == "Период") { checkBox7.Checked = false; dateTimePicker2.Enabled = false; }
            }
            if (!checkBox6.Checked)
            {
                dateTimePicker1.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
                //dateTimePicker2.Enabled = false;
                checkBox7.Checked = false;
                button11_Click(sender, e);
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                checkBox6.Text = "С";
                checkBox7.Text = "По";
                dateTimePicker1.Enabled = true;
                button11.Enabled = true;
                button12.Enabled = true;
                dateTimePicker2.Enabled = true;
                checkBox6.Checked = true;
            }
            if (!checkBox7.Checked)
            {
                checkBox6.Text = "Дата";
                checkBox7.Text = "Период";
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
                checkBox6.Checked = false;
                button11_Click(sender, e);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                if (textBox1.Text.Contains("ПАРТИЯ:"))
                {
                    if (textBox1.Text.Contains(comboBox1.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                    textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(";"),","+comboBox1.Text);
                    //textBox1.Text.IndexOf(";").ToString();
                }
                if (!textBox1.Text.Contains("ПАРТИЯ:"))
                {
                    textBox1.Text=textBox1.Text.Insert(0,"ПАРТИЯ:"+comboBox1.Text+";");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("ПАРТИЯ:"))
            {
                textBox1.Text=textBox1.Text.Remove(0, textBox1.Text.IndexOf(";",0)+1);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
            {
                string s="Продукт:";
                s = s.ToUpper();
                int start = 0;
                start = textBox1.Text.IndexOf(";") + 1;
                if (textBox1.Text.Contains(s))
                {
                    if (textBox1.Text.Contains(comboBox2.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                    textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(s)+s.Length,  comboBox2.Text+",");
                    //textBox1.Text.IndexOf(";").ToString();
                }
                if (!textBox1.Text.Contains(s))
                {                                        
                    textBox1.Text = textBox1.Text.Insert(start, s+comboBox2.Text + ";");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("ПРОДУКТ:"))
            {
                int start = textBox1.Text.IndexOf("ПРОДУКТ:");
                int end = textBox1.Text.IndexOf(";", start);
                int count = end - start;
                textBox1.Text = textBox1.Text.Remove(start,count+1);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (comboBox4.Text != "")
            {
                string s = "Станок:";
                s = s.ToUpper();
                int start = 0;
                start = textBox1.Text.IndexOf(";");
                if (textBox1.Text.Contains(s))
                {
                    if (textBox1.Text.Contains(comboBox4.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                    textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(s) + s.Length, comboBox4.Text + ",");
                    //textBox1.Text.IndexOf(";").ToString();
                }
                if (!textBox1.Text.Contains(s))
                {
                    textBox1.Text = textBox1.Text.Insert(start + 1, s + comboBox4.Text + ";");
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("СТАНОК:"))
            {
                int start = textBox1.Text.IndexOf("СТАНОК:");
                int end = textBox1.Text.IndexOf(";", start);
                int count = end - start;
                textBox1.Text = textBox1.Text.Remove(start, count + 1);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text != "")
            {
                string s = "Работник:";
                s = s.ToUpper();
                int start = 0;
                start = textBox1.Text.IndexOf(";");
                if (textBox1.Text.Contains(s))
                {
                    if (textBox1.Text.Contains(comboBox3.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                    textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(s) + s.Length, comboBox3.Text + ",");
                    //textBox1.Text.IndexOf(";").ToString();
                }
                if (!textBox1.Text.Contains(s))
                {
                    textBox1.Text = textBox1.Text.Insert(start + 1, s + comboBox3.Text + ";");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("РАБОТНИК:"))
            {
                int start = textBox1.Text.IndexOf("РАБОТНИК:");
                int end = textBox1.Text.IndexOf(";", start);
                int count = end - start;
                textBox1.Text = textBox1.Text.Remove(start, count + 1);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (comboBox5.Text != "")
            {
                string s = "Смена:";
                s = s.ToUpper();
                int start = 0;
                start = textBox1.Text.IndexOf(";");
                if (textBox1.Text.Contains(s))
                {
                    if (textBox1.Text.Contains(comboBox5.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                    textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(s) + s.Length, comboBox5.Text + ",");
                    //textBox1.Text.IndexOf(";").ToString();
                }
                if (!textBox1.Text.Contains(s))
                {
                    textBox1.Text = textBox1.Text.Insert(start + 1, s +comboBox5.Text + ";");
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("СМЕНА:"))
            {
                int start = textBox1.Text.IndexOf("СМЕНА:");
                int end = textBox1.Text.IndexOf(";", start);
                int count = end - start;
                textBox1.Text = textBox1.Text.Remove(start, count + 1);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (checkBox6.Checked && !checkBox7.Checked)
            {
                if (dateTimePicker1.Text != "")
                {
                    if (textBox1.Text.Contains("ПЕРИОД"))
                    {
                        MessageBox.Show("Возможно добавить или дату или период!");
                        return;
                    }
                    string s = "Дата:";
                    int start = 0;
                    s=s.ToUpper();
                    start = textBox1.Text.IndexOf(";");
                    if (textBox1.Text.Contains(s))
                    {
                        if (textBox1.Text.Contains(dateTimePicker1.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                        textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(s) + s.Length, dateTimePicker1.Text + ",");
                        //textBox1.Text.IndexOf(";").ToString();
                    }
                    if (!textBox1.Text.Contains(s))
                    {
                        textBox1.Text = textBox1.Text.Insert(start + 1, s + dateTimePicker1.Text + ";");
                    }
                }
            }
            if (checkBox6.Checked && checkBox7.Checked)
            {
                if (dateTimePicker1.Text != ""&&dateTimePicker2.Text!="")
                {
                    if (textBox1.Text.Contains("ДАТА"))
                    {
                        MessageBox.Show("Возможно добавить или дату или период!");
                        return;
                    }
                    if (Convert.ToDateTime(dateTimePicker1.Text) >= Convert.ToDateTime(dateTimePicker2.Text))
                    {
                        MessageBox.Show("Первая дата периода не должна превышать последнюю!");
                        return;
                    }
                    string s = "Период:";
                    int start = 0;
                    s = s.ToUpper();
                    start = textBox1.Text.IndexOf(";");
                    if (textBox1.Text.Contains(s))
                    {
                        if (textBox1.Text.Contains(dateTimePicker1.Text)&&textBox1.Text.Contains(dateTimePicker2.Text)) { MessageBox.Show("Такое наименование уже есть!"); return; }
                        MessageBox.Show("Возможно добавить только один период!");
                        return;
                        //textBox1.Text = textBox1.Text.Insert(textBox1.Text.IndexOf(s) + s.Length, " " + dateTimePicker1.Text + ",");
                        //textBox1.Text.IndexOf(";").ToString();
                    }
                    if (!textBox1.Text.Contains(s))
                    {
                        textBox1.Text = textBox1.Text.Insert(start + 1, s + dateTimePicker1.Text+"-" +dateTimePicker2.Text+ ";");
                    }
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains("ДАТА:"))
            {
                int start = textBox1.Text.IndexOf("ДАТА:");
                int end = textBox1.Text.IndexOf(";", start);
                int count = end - start;
                textBox1.Text = textBox1.Text.Remove(start, count + 1);
            }
            if (textBox1.Text.Contains("ПЕРИОД:"))
            {
                int start = textBox1.Text.IndexOf("ПЕРИОД:");
                int end = textBox1.Text.IndexOf(";", start);
                int count = end - start;
                textBox1.Text = textBox1.Text.Remove(start, count + 1);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "") { MessageBox.Show("Выберите что-нибудь!"); return; }
            string partiya, prodykt, stanok, rabotnik, smena, data1, data2;
            string partiya1, prodykt1, stanok1, rabotnik1, smena1, data11, data21;
            string partiya3=""; string prodykt3=""; string stanok3=""; string rabotnik3=""; string smena3=""; string data3=""; string data4="";
            ArrayList partiya2, prodykt2, stanok2, rabotnik2, smena2, data12, data22;
            prodykt = "ПРОДУКТ:"; partiya = "ПАРТИЯ:"; stanok = "СТАНОК:"; rabotnik = "РАБОТНИК:"; smena = "СМЕНА:"; data1 = "ДАТА:"; data2 = "ПЕРИОД:";
            
            
            
            prodykt2 = new ArrayList();
            if(textBox1.Text.Contains(prodykt))
            {
                
                int start = textBox1.Text.IndexOf(prodykt) + prodykt.Length;
                int end = textBox1.Text.IndexOf(";", start);
                prodykt1=textBox1.Text.Substring(start,end-start);
                if (prodykt1.Contains("|"))
                {
                    int old = 0;                   
                    for (int i = 0; i < prodykt1.Length; i++)
                    {
                        if (prodykt1[i] == ',')
                        {                            
                            prodykt2.Add(prodykt1.Substring(old,i-old));
                            old = i+1;
                        }
                    }
                    prodykt2.Add(prodykt1.Substring(old, prodykt1.Length-old));
                }
                else
                {
                    prodykt2.Add(prodykt1);
                }
                //цикл для продукта

                if (prodykt2.Count == 1) prodykt3 = " and prodykt.name='" + (string)prodykt2[0] + "'";
                //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                else
                {
                    prodykt3 = " and prodykt.name in (";
                    for (int i = 0; i < prodykt2.Count-1; i++)
                    {
                        prodykt3 += "'" + (string)prodykt2[i] + "'|";
                    }
                    prodykt3 += "'"+(string)prodykt2[prodykt2.Count-1]+"') ";
                }                

            }

            partiya2 = new ArrayList();
            if (textBox1.Text.Contains(partiya))
            {                
                int start = textBox1.Text.IndexOf(partiya) + partiya.Length;
                int end = textBox1.Text.IndexOf(";", start);
                partiya1 = textBox1.Text.Substring(start, end - start);
                if (partiya1.Contains(","))
                {
                    int old = 0;
                    for (int i = 0; i < partiya1.Length; i++)
                    {
                        if (partiya1[i] == ',')
                        {
                            partiya2.Add(partiya1.Substring(old, i - old));
                            old = i + 1;
                        }
                    }
                    partiya2.Add(partiya1.Substring(old, partiya1.Length - old));
                }
                else
                {
                    partiya2.Add(partiya1);
                }
                //цикл для партии

                if (partiya2.Count == 1) partiya3 = " and partiya.name='" + (string)partiya2[0] + "'";
                //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                else
                {
                    partiya3 = " and partiya.name in (";
                    for (int i = 0; i < partiya2.Count - 1; i++)
                    {
                        partiya3 += "'" + (string)partiya2[i] + "',";
                    }
                    partiya3 += "'" + (string)partiya2[partiya2.Count - 1] + "') ";
                }        
            }

            if (!textBox1.Text.Contains(partiya))
            {
                //записываем партии из кб
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    partiya2.Add(comboBox1.Items[i].ToString());
                }
                    //цикл для партии

                    if (partiya2.Count == 1) partiya3 = " and partiya.name='" + (string)partiya2[0] + "'";
                    //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                    else
                    {
                        partiya3 = " and partiya.name in (";
                        for (int i = 0; i < partiya2.Count - 1; i++)
                        {
                            partiya3 += "'" + (string)partiya2[i] + "',";
                        }
                        partiya3 += "'" + (string)partiya2[partiya2.Count - 1] + "') ";
                    }
            }

            stanok2 = new ArrayList();
            if (textBox1.Text.Contains(stanok))
            {                
                int start = textBox1.Text.IndexOf(stanok) + stanok.Length;
                int end = textBox1.Text.IndexOf(";", start);
                stanok1 = textBox1.Text.Substring(start, end - start);
                if (stanok1.Contains(","))
                {
                    int old = 0;
                    for (int i = 0; i < stanok1.Length; i++)
                    {
                        if (stanok1[i] == ',')
                        {
                            stanok2.Add(stanok1.Substring(old, i - old));
                            old = i + 1;
                        }
                    }
                    stanok2.Add(stanok1.Substring(old, stanok1.Length - old));
                }
                else
                {
                    stanok2.Add(stanok1);
                }

                //цикл для станка

                if (stanok2.Count == 1) stanok3 = " and stanok.name='" + (string)stanok2[0] + "'";
                //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                else
                {
                    stanok3 = " and stanok.name in (";
                    for (int i = 0; i < stanok2.Count - 1; i++)
                    {
                        stanok3 += "'" + (string)stanok2[i] + "',";
                    }
                    stanok3 += "'" + (string)stanok2[stanok2.Count - 1] + "') ";
                }
            }

            rabotnik2 = new ArrayList();
            if (textBox1.Text.Contains(rabotnik))
            {                
                int start = textBox1.Text.IndexOf(rabotnik) + rabotnik.Length;
                int end = textBox1.Text.IndexOf(";", start);
                rabotnik1 = textBox1.Text.Substring(start, end - start);
                if (rabotnik1.Contains(","))
                {
                    int old = 0;
                    for (int i = 0; i < rabotnik1.Length; i++)
                    {
                        if (rabotnik1[i] == ',')
                        {
                            rabotnik2.Add(rabotnik1.Substring(old, i - old));
                            old = i + 1;
                        }
                    }
                    rabotnik2.Add(rabotnik1.Substring(old, rabotnik1.Length - old));
                }
                else
                {
                    rabotnik2.Add(rabotnik1);
                }

                //цикл для работника

                if (rabotnik2.Count == 1) rabotnik3 = " and rabotnik.surname='" + (string)rabotnik2[0] + "'";
                //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                else
                {
                    rabotnik3 = " and rabotnik.surname in (";
                    for (int i = 0; i < rabotnik2.Count - 1; i++)
                    {
                        rabotnik3 += "'" + (string)rabotnik2[i] + "',";
                    }
                    rabotnik3 += "'" + (string)rabotnik2[rabotnik2.Count - 1] + "') ";
                }
                      
            }

            smena2 = new ArrayList();
            if (textBox1.Text.Contains(smena))
            {                
                int start = textBox1.Text.IndexOf(smena) + smena.Length;
                int end = textBox1.Text.IndexOf(";", start);
                smena1 = textBox1.Text.Substring(start, end - start);
                if (smena1.Contains(","))
                {
                    int old = 0;
                    for (int i = 0; i < smena1.Length; i++)
                    {
                        if (smena1[i] == ',')
                        {
                            smena2.Add(smena1.Substring(old, i - old));
                            old = i + 1;
                        }
                    }
                    smena2.Add(smena1.Substring(old, smena1.Length - old));
                }
                else
                {
                    smena2.Add(smena1);
                }
                //для смены                
                int sm1 = 0;
                if ((string)smena2[0] == "День") sm1 = 1;
                smena3 = " and sobitie.smena=" + sm1;                
            }


            data12 = new ArrayList();
            if (textBox1.Text.Contains(data1))
            {                
                int start = textBox1.Text.IndexOf(data1) + data1.Length;
                int end = textBox1.Text.IndexOf(";", start);
                data11 = textBox1.Text.Substring(start, end - start);
                if (data11.Contains(","))
                {
                    int old = 0;
                    for (int i = 0; i < data11.Length; i++)
                    {
                        if (data11[i] == ',')
                        {
                            data12.Add(data11.Substring(old, i - old));
                            old = i + 1;
                        }
                    }
                    data12.Add(data11.Substring(old, data11.Length - old));
                }
                else
                {
                    data12.Add(data11);
                }

                //цикл для даты                                                
                if (data12.Count == 1) data3 = " and sobitie.data='" + (string)data12[0] + "'";
                //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                else
                {
                    data3 = " and sobitie.data in (";
                    for (int i = 0; i < data12.Count - 1; i++)
                    {
                        data3 += "'" + (string)data12[i] + "',";
                    }
                    data3 += "'" + (string)data12[data12.Count - 1] + "') ";
                }                
            }


            data22 = new ArrayList();
            if (textBox1.Text.Contains(data2))
            {                
                int start = textBox1.Text.IndexOf(data2) + data2.Length;
                int end = textBox1.Text.IndexOf(";", start);
                data21 = textBox1.Text.Substring(start, end - start);
                string d1 = data21.Substring(0, data21.IndexOf("-"));
                start=data21.IndexOf("-")+1;
                end=data21.Length - data21.IndexOf("-");
                string d2 = data21.Substring(start,end-1 );
                DateTime dt1 = Convert.ToDateTime(d1);
                DateTime dt2 = Convert.ToDateTime(d2);                                
                TimeSpan dt4 = new TimeSpan(1, 0, 0, 0);
                while (dt1 <= dt2)
                {
                    data22.Add(dt1.ToLongDateString());
                    dt1 = dt1 + dt4;
                }

                //цикл для периода       
                if (data22.Count == 1) data4 = " and sobitie.data='" + (string)data22[0] + "'";
                //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
                else
                {
                    data4 = " and sobitie.data in (";
                    for (int i = 0; i < data22.Count - 1; i++)
                    {
                        data4 += "'" + (string)data22[i] + "',";
                    }
                    data4 += "'" + (string)data22[data22.Count - 1] + "') ";
                }                
            }
            //формирование строки запроса
            conn.Open();
            command.CommandText = "select sobitie.id from sobitie,vessklad,prodykt,partiya where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.iddvigfrom=2 and sobitie.iddvig=2 and sobitie.idbalans=2"+prodykt3+partiya3+smena3+data3+data4;
            r = command.ExecuteReader();
            ArrayList idlist = new ArrayList();
            if(r.HasRows)
                while (r.Read())
                {
                    idlist.Add(r[0]);
                }
            conn.Close();
            if (idlist.Count == 0) { MessageBox.Show("Ничего не найдено!"); return; }
            string idstr = "";
            //цикл для id
            if (idlist.Count == 1) idstr = " and sobitie.id=" + idlist[0].ToString();
            //prodykt3 = " and prodykt.name in ('" + (string)prodykt2[0] + "'";
            else
            {
                idstr = " and sobitie.id in (";
                int kk = 0;
                for (int i = 0; i < idlist.Count - 1; i++)
                {
                    kk = (int)idlist[i];
                    idstr += kk.ToString() + ",";
                }
                kk = (int)idlist[idlist.Count - 1];
                idstr += kk.ToString() + ") ";
            }
            
            conn.Open();
            command.CommandText = "select sobitie.id,sobitie.idproizv from sobitie,rabotnik,vessklad,proizv,stanok where sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id and rabotnik.id=vessklad.idrabotnik and stanok.id=vessklad.idstanok" + idstr + rabotnik3 + stanok3+" order by sobitie.idproizv asc";
            idlist = new ArrayList();
            ArrayList prlist;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    prlist = new ArrayList();
                    prlist.Add(r[0]);
                    prlist.Add(r[1]);

                    idlist.Add(prlist);
                }
            }
            conn.Close();
            if (idlist.Count == 0) { MessageBox.Show("Ничего не найдено!"); return; }

            //конец работы фильтра!!!


            listView1.Items.Clear();
            ListViewItem lvi;
            bool doit = true;
            int id = 0; string ves = ""; string idpr = "";
            partiya = ""; prodykt = ""; data1 = ""; smena = ""; rabotnik = ""; stanok = "";            
            for (int i = 0; i < idlist.Count; i++)
            {
                prlist = new ArrayList();
                prlist = (ArrayList)idlist[i];
                if (doit)
                {
                    id = (int)prlist[0];

                    conn.Open();
                    command.CommandText = "select sobitie.data,sobitie.smena,sobitie.idproizv from sobitie where sobitie.id=" + id.ToString();
                    r = command.ExecuteReader();
                    r.Read();
                    data1 = r[0].ToString();
                    smena = "Ночь";
                    try
                    {
                        if ((bool)r[1]) smena = "День";
                    }
                    catch (System.Exception)
                    {
                        smena = "Не указано";
                    }
                    idpr = r[2].ToString();
                    conn.Close();

                    conn.Open();
                    command.CommandText = "select rabotnik.surname,stanok.name from vessklad,proizv,stanok,rabotnik where rabotnik.id=vessklad.idrabotnik and stanok.id=vessklad.idstanok and proizv.idprodykt1=vessklad.id and proizv.id=" + idpr;
                    r = command.ExecuteReader();
                    r.Read();
                    rabotnik = r[0].ToString(); stanok = r[1].ToString();
                    conn.Close();

                    //поиск подсобников
                    conn.Open();
                    command.CommandText = "select rabotnik.surname from rabotnik,podsobniki where podsobniki.idpodsobnik=rabotnik.id and podsobniki.idsobitie=" + id.ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {
                            rabotnik += "," + (string)r[0];
                        }
                    }
                    conn.Close();

                    lvi = new ListViewItem(new string[] { data1, smena, stanok, rabotnik, "ВЗЯТО", "", "" });
                    lvi.BackColor = Color.Bisque;
                    listView1.Items.Add(lvi);
                    //lvi = new ListViewItem(new string[] { "", "", "", "", "", "", "" });
                    //listView1.Items.Add(lvi);

                    conn.Open();
                    command.CommandText = "select sobitie.id,partiya.name,prodykt.name, sobitie.ves from sobitie,vessklad,partiya,prodykt where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.idproizv=" + idpr + idstr;
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {
                            //if ((int)r[0] != id) idlist[idlist.IndexOf(r[0])] = 0;
                            partiya = r[1].ToString(); prodykt = r[2].ToString(); ves = r[3].ToString();
                            lvi = new ListViewItem(new string[] { "", "", "", "", partiya, prodykt, ves });
                            lvi.BackColor = Color.Bisque;
                            listView1.Items.Add(lvi);
                        }
                    }
                    conn.Close();

                    lvi = new ListViewItem(new string[] { "", "", "", "", "ПОЛУЧЕНО", "", "" });
                    lvi.BackColor = Color.FromArgb(220,234,250);
                    listView1.Items.Add(lvi);

                    conn.Open();
                    command.CommandText = "select sobitie.id,partiya.name,prodykt.name, sobitie.ves from sobitie,vessklad,partiya,prodykt where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.iddvigfrom=2 and sobitie.iddvig=2 and sobitie.idbalans=1  and sobitie.idproizv=" + idpr;
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {
                            //if ((int)r[0] != id) idlist[idlist.IndexOf(r[0])] = 0;
                            partiya = r[1].ToString(); prodykt = r[2].ToString(); ves = r[3].ToString();
                            lvi = new ListViewItem(new string[] { "", "", "", "", partiya, prodykt, ves });
                            lvi.BackColor = Color.FromArgb(220, 234, 250);
                            listView1.Items.Add(lvi);
                        }
                    }
                    conn.Close();
                    //ищем мусор
                    conn.Open();
                    command.CommandText = "select tipothoda.name, othodi.ves from tipothoda,othodi where othodi.idtipothoda=tipothoda.id and othodi.idproizv="+idpr;
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {                                                        
                            lvi = new ListViewItem(new string[] { "", "", "", "", "ОТХОДЫ", r[0].ToString(), r[1].ToString() });
                            lvi.BackColor = Color.FromArgb(174, 182, 191);
                            listView1.Items.Add(lvi);
                        }
                    }
                    conn.Close();

                    lvi = new ListViewItem(new string[] { "", "", "", "", "", "", "" });
                    listView1.Items.Add(lvi);
              
                }
                if (i < idlist.Count - 1)
                {
                    ArrayList temp = (ArrayList)idlist[i + 1];
                    if ((int)prlist[1] == (int)temp[1]) doit = false;
                    else doit = true;
                }
            }
            
       
        }

        private void button14_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Переработка";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();


            int r = 0;
            int c = 0;
            c = listView1.Columns.Count;
            r = listView1.Items.Count+1;
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, r, c, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;

            for (int i = 1; i <= c; i++)
            {
                oTable.Cell(1, i).Range.Text = listView1.Columns[i - 1].Text;
            }

            for (int i = 2; i <= r; i++)
            {
                for (int j = 1; j <= c; j++)
                {
                    if (listView1.Items[i - 2].SubItems[j - 1].Text == "Итого:") oTable.Rows[i].Range.Font.Shadow = 5;
                    oTable.Cell(i, j).Range.Text = listView1.Items[i - 2].SubItems[j - 1].Text;
                }
            }
            oTable.Borders.Enable = 1;
            oTable.Borders[WdBorderType.wdBorderTop].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Borders[WdBorderType.wdBorderLeft].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Borders[WdBorderType.wdBorderRight].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;
            oWord.Visible = true;            
        }
    }
}