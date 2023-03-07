using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;


namespace Polohov
{
    public partial class Rabochie : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        //int id;
        decimal sumvesa;
        //ArrayList arr;
        //ArrayList temp;
        ListViewItem li;
        public Rabochie(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;
            InitializeComponent();

            listView1.Columns.Add("Партия", 50);
            listView1.Columns.Add("Продукт", 150);
            listView1.Columns.Add("Станок", 100);
            listView1.Columns.Add("Взято(кг)", 80);
            listView1.Columns.Add("id", 0);            
            listView1.Columns.Add("Себест.(грн)", 80);
            //listView1.Columns.Add("idpr", 0);

            listView2.Columns.Add("Фамилия",100);
            listView2.Columns.Add("Смена", 50);
            listView2.Columns.Add("З\\п(грн)", 50);

            comboBox3.Items.Add("Ночь");
            comboBox3.Items.Add("День");

            conn.Open();
            command.CommandText = "select surname from rabotnik";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    if((string)r[0]!="")
                    comboBox2.Items.Add(r[0]);
                }
            }
            conn.Close();
            //ищем фамилии рабочих без з/п
            conn.Open();
            command.CommandText = "select rabotnik.surname from rabotnik,sobitie,proizv,vessklad where vessklad.idrabotnik=rabotnik.id and proizv.idprodykt1=vessklad.id and sobitie.idproizv=proizv.id and sobitie.price is null and iddvig=2 and idbalans=2";
            r = command.ExecuteReader();
            int x = 0;
            if (r.HasRows)
            {
                while (r.Read())
                {
                    x = 0;
                    for (int i = 0; i < comboBox1.Items.Count; i++)
                    {
                        if ((string)r[0] == (string)comboBox1.Items[i]) x = 1;
                    }
                    if(x!=1)comboBox1.Items.Add(r[0]);
                }
            }
            conn.Close();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox4.Items.Clear();
            conn.Open();
            command.CommandText = "select sobitie.data from sobitie,rabotnik,proizv,vessklad where sobitie.price is null and iddvig=2 and idbalans=2 and sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id and vessklad.idrabotnik=rabotnik.id and rabotnik.surname='"+comboBox1.Text+"'";
            r = command.ExecuteReader();
            int x = 0;
            if (r.HasRows)
            {
                while (r.Read())
                {
                    x = 0;
                    for (int i = 0; i < comboBox4.Items.Count; i++)
                    {
                        if ((string)r[0] == (string)comboBox4.Items[i]) x = 1;
                    }
                    if(x!=1)comboBox4.Items.Add(r[0]);
                }
            }
            conn.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox4.Text != "")
            {
                listView1.Items.Clear();
                //listView1.Items.Add("Взято в работу");
                conn.Open();
                string[] s = new string[6];
                sumvesa = 0;
                command.CommandText = "select partiya.name,stanok.name,sobitie.ves,sobitie.id from partiya,stanok,vessklad,sobitie,proizv,rabotnik where partiya.id=vessklad.idpartiya and  stanok.id=vessklad.idstanok and sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id  and sobitie.price is null and iddvig=2 and idbalans=2 and rabotnik.id=vessklad.idrabotnik and rabotnik.surname='" + comboBox1.Text + "' and sobitie.data='"+comboBox4.Text+"'";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        s[0] = r[0].ToString();
                        s[1] = "";
                        s[2] = r[1].ToString();
                        s[3] = r[2].ToString();
                        s[4] = r[3].ToString();
                        s[5] = "";
                        
                        li = new ListViewItem(s);
                        listView1.Items.Add(li);
                        sumvesa += Convert.ToDecimal(r[2].ToString().Replace('.',','));
                    }
                }
                conn.Close();
                conn.Open();
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    command.CommandText = "select prodykt.name from prodykt,sobitie,vessklad where vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id and sobitie.id="+listView1.Items[i].SubItems[4].Text;
                    //string prod = (string)command.ExecuteScalar();
                    listView1.Items[i].SubItems[1].Text = (string)command.ExecuteScalar();
                }
                conn.Close();
                listView2.Items.Clear();
                string[] s1 = new string[3];
                s1[0] = comboBox1.Text;
                s1[1] = "";
                s1[2] = "";
                li = new ListViewItem(s1);
                listView2.Items.Add(li);

                //Вес итого
                listView1.Items.Add("");
                s[0] = "Итого:"; s[1] = ""; s[2] = ""; s[3] = sumvesa.ToString(); s[4] = ""; s[5] = "";
                li = new ListViewItem(s);
                listView1.Items.Add(li);             
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listView2.SelectedIndices.Count != 0)
            {
                if (listView2.SelectedIndices[0] == 0) { MessageBox.Show("Удалять можно только подсобников!"); return; }
                listView2.Items.RemoveAt(listView2.SelectedIndices[0]);
                listView2.Refresh();
            }            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(comboBox2.Text!="")
            {
                for (int i = 0; i < listView2.Items.Count; i++)
                {
                    if (comboBox2.Text == listView2.Items[i].SubItems[0].Text)
                    {
                        MessageBox.Show("Этот рабочий уже есть в списке!");
                        return;
                    }
                }
                string[] s1 = new string[3];
                s1[0] = comboBox2.Text;
                s1[1] = "";
                s1[2] = "";
                ListViewItem li = new ListViewItem(s1);
                listView2.Items.Add(li);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text == "" || textBox1.Text == "")
            {
                MessageBox.Show("Для продолжения необходимо выбрать смену и указать з/п!");
                return;
            }
            if (listView2.SelectedIndices.Count == 0 && listView2.Items.Count > 1)
            {
                MessageBox.Show("Выберите строку куда хотите добавить!");
                return;
            }
            if (listView2.Items.Count == 1) listView2.Items[0].Selected = true;
            try
            {
                decimal x=Convert.ToDecimal(textBox1.Text.Replace('.',','));
                if (x <= 0) { MessageBox.Show("Зарплата должна быть больше 0!"); return; }
            }
            catch (System.Exception)
            {
                MessageBox.Show("В поле З/П необходимо ввести число!");
                return;
            }


            for (int i = 0; i < listView2.Items.Count; i++)
            {
                listView2.Items[i].SubItems[1].Text = comboBox3.Text;
            }
            listView2.Items[listView2.SelectedIndices[0]].SubItems[2].Text = textBox1.Text;
           
            
            
            decimal zpvsya = 0;
            for (int i = 0; i < listView2.Items.Count; i++)
            {
                if(listView2.Items[i].SubItems[2].Text!="")
                zpvsya += Convert.ToDecimal(listView2.Items[i].SubItems[2].Text.Replace('.',','));
            }
            decimal sumrez = 0;
            for (int i = 0; i < listView1.Items.Count - 2; i++)
            {
                decimal rez = 0;
                rez = zpvsya * Convert.ToDecimal(listView1.Items[i].SubItems[3].Text.Replace('.',',')) / sumvesa;
                listView1.Items[i].SubItems[5].Text = rez.ToString();
                sumrez += rez;
            }
            listView1.Items[listView1.Items.Count - 1].SubItems[5].Text = sumrez.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //проверка на одинаковость смены у подсобника и у рабочего
            for (int i = 0; i < listView2.Items.Count; i++)
            {
                for (int k = 0; k < listView2.Items[0].SubItems.Count; k++)
                {
                    if (listView2.Items[i].SubItems[k].Text == "")
                    {
                        MessageBox.Show("Заполните все поля!");
                        return;
                    }
                }
            }
            conn.Open();
            for (int i = 0; i < listView1.Items.Count-2; i++)
            {
                //conn.Open();
                command.CommandText = "update sobitie set price=" + listView1.Items[i].SubItems[5].Text.Replace(',','.') + " where id=" + listView1.Items[i].SubItems[4].Text;
                command.ExecuteNonQuery();
                int sm = 0;
                if (listView2.Items[0].SubItems[1].Text == "День") sm = 1;
                if (listView2.Items[0].SubItems[1].Text == "Ночь") sm = 0;
                command.CommandText = "update sobitie set smena=" + sm.ToString() + " where id=" + listView1.Items[i].SubItems[4].Text;
                command.ExecuteNonQuery();
                //conn.Close();
            }

            command.CommandText = "select id from rabotnik where surname='" + listView2.Items[0].SubItems[0].Text + "'";
            int idrab = (int)command.ExecuteScalar();
            command.CommandText = "insert into zarplata values(" + idrab.ToString() + ",'" + comboBox4.Text + "','" + listView2.Items[0].SubItems[1].Text.Replace(',', '.') + "'," + listView2.Items[0].SubItems[2].Text + ")";
            command.ExecuteNonQuery();

            for (int i = 1; i < listView2.Items.Count; i++)
            {
                //conn.Open();
                command.CommandText = "select id from rabotnik where surname='"+listView2.Items[i].SubItems[0].Text+"'";
                int idrab1 = (int)command.ExecuteScalar();
                command.CommandText = "insert into zarplata values(" + idrab1.ToString() + ",'" + comboBox4.Text + "','" + listView2.Items[i].SubItems[1].Text.Replace(',', '.') + "'," + listView2.Items[i].SubItems[2].Text + ")";
                command.ExecuteNonQuery();
                
                for(int k=0;k<listView1.Items.Count-2;k++)
                {
                    command.CommandText = "insert into podsobniki values("+listView1.Items[k].SubItems[4].Text+","+idrab1.ToString()+")";
                    command.ExecuteNonQuery();
                }
            }
            conn.Close();
            listView2.Items.Clear();
            listView1.Items.Clear();
            comboBox1.Items.Clear();
            comboBox4.Items.Clear();
            comboBox2.ResetText();
            comboBox3.ResetText();
            textBox1.Text = "";


            //conn.Close();
            //ищем фамилии рабочих без з/п
            conn.Open();
            command.CommandText = "select rabotnik.surname from rabotnik,sobitie,proizv,vessklad where vessklad.idrabotnik=rabotnik.id and proizv.idprodykt1=vessklad.id and sobitie.idproizv=proizv.id and sobitie.price is null and iddvig=2 and idbalans=2";
            r = command.ExecuteReader();
            int x = 0;
            if (r.HasRows)
            {
                while (r.Read())
                {
                    x = 0;
                    for (int i = 0; i < comboBox1.Items.Count; i++)
                    {
                        if ((string)r[0] == (string)comboBox1.Items[i]) x = 1;
                    }
                    if (x != 1) comboBox1.Items.Add(r[0]);
                }
            }
            conn.Close();
        }
    }
}