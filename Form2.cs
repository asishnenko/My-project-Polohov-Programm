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
    public partial class Form2 : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        string fam, stanok, partiya, prod;
        //int ves = 0;
        //int idproizv = 0;
        //int idsklad = 0;
        ArrayList arr;
        ArrayList temp;
        public Form2(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;
            InitializeComponent();
            listView1.Columns.Add("Фамилия",80);
            listView1.Columns.Add("Станок",80);
            listView1.Columns.Add("Партия");
            listView1.Columns.Add("Наименование",120);
            listView1.Columns.Add("Взято,кг");
            listView1.Columns.Add("Дата",100);
            listView1.Columns.Add("З/п", 50);
            listView1.Columns.Add("", 0);
            listView1.Columns.Add("Смена", 50);
            comboBox2.Items.Add("Ночь");
            comboBox2.Items.Add("День");            
            FillListView();
        }
        void FillListView()
        {
            //находим дату и вес из события каждой переработки у которой цена =0
            arr=new ArrayList();
            conn.Open();
            command.CommandText = "select idsklad,ves,data,idproizv,id from sobitie where iddvig=2 and idbalans=2 and price is null";
            r = command.ExecuteReader();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    temp = new ArrayList();
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        temp.Add(r[i]);
                    }
                    arr.Add(temp);
                }
            }
            conn.Close();
            //находим название сырья и партию из вессклад по айди переработанного
            for (int i = 0; i < arr.Count; i++)
            {
                temp=(ArrayList)arr[i];  
                conn.Open();
                command.CommandText = "select partiya.name,prodykt.name from vessklad,partiya,prodykt where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id="+(int)temp[0];
                r = command.ExecuteReader();
                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        partiya = (string)r[0];
                        prod = (string)r[1];
                    }
                }
                conn.Close();
                //находим айди любого полученного из айди произв
                conn.Open();
                command.CommandText = "select idprodykt1 from proizv where id="+(int)temp[3];
                int idprod = (int)command.ExecuteScalar();
                //находим фамилию и станок
                command.CommandText = "select rabotnik.surname, stanok.name from rabotnik,stanok,vessklad where vessklad.idrabotnik=rabotnik.id and vessklad.idstanok=stanok.id and vessklad.id="+idprod.ToString();
                r = command.ExecuteReader();
                r.Read();
                fam = (string)r[0];
                stanok = (string)r[1];
                conn.Close();
                //далее заполнять список
                string[] s = new string[9];
                s[0] = fam; s[1] = stanok; s[2] = partiya; s[3] = prod; s[4] = (string)temp[1].ToString(); s[5] = (string)temp[2]; s[6] = ""; s[7] = temp[4].ToString(); s[8] = "";
                ListViewItem li = new ListViewItem(s);
                listView1.Items.Add(li);                
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count != 0) //MessageBox.Show("Выдайте всем з/п!");
            //else
            {
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    if(listView1.Items[i].SubItems[6].Text!="")
                    {
                        conn.Open();
                        command.CommandText = "update sobitie set price=" + listView1.Items[i].SubItems[6].Text + " where id=" + listView1.Items[i].SubItems[7].Text;
                        command.ExecuteNonQuery();
                        int sm = 0;
                        if (listView1.Items[i].SubItems[8].Text == "День") sm = 1;
                        if (listView1.Items[i].SubItems[8].Text == "Ночь") sm = 2;
                        command.CommandText = "update sobitie set smena=" + sm.ToString() +" where id=" + listView1.Items[i].SubItems[7].Text;
                        command.ExecuteNonQuery();
                        conn.Close();
                        listView1.Items[i].Remove();
                        listView1.Refresh();
                    }                    
                }                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count != 0) MessageBox.Show("Не забудьте выдать з/п!");
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && textBox2.Text != "" && textBox1.Text != "")
            {
                try
                {
                    if (Convert.ToInt32(textBox1.Text) > 0 && Convert.ToInt32(textBox1.Text) < 300)
                    {
                        ArrayList same = new ArrayList();
                        for (int i = 0; i < listView1.Items.Count; i++)
                        {
                            if (textBox2.Text == listView1.Items[i].SubItems[0].Text && textBox3.Text == listView1.Items[i].SubItems[5].Text)
                            {
                                same.Add(i);
                            }
                        }
                        if (same.Count == 1)
                        {
                            //conn.Open();
                            //command.CommandText = "update sobitie set price="+textBox1.Text+" where id="+listView1.Items[(int)same[0]].SubItems[7];
                            //command.ExecuteNonQuery();
                            //command.CommandText = "update sobitie set smena=" + comboBox2.SelectedIndex + " where id=" + listView1.Items[(int)same[0]].SubItems[7];
                            //command.ExecuteNonQuery();
                            //conn.Close();
                            listView1.Items[(int)same[0]].SubItems[6].Text = textBox1.Text;
                            listView1.Items[(int)same[0]].SubItems[8].Text = comboBox2.Text;
                            textBox1.SelectAll();
                        }
                        if (same.Count > 1)
                        {
                            int summVes = 0;
                            for (int i = 0; i < same.Count; i++)
                            {
                                summVes += Convert.ToInt32(listView1.Items[(int)same[i]].SubItems[4].Text);
                            }
                            //conn.Open();
                            for (int i = 0; i < same.Count; i++)
                            {
                                int ctena = Convert.ToInt32(textBox1.Text) * Convert.ToInt32(listView1.Items[(int)same[i]].SubItems[4].Text) / summVes;
                                
                                //command.CommandText = "update sobitie set price=" + ctena + " where id=" + listView1.Items[(int)same[i]].SubItems[7];
                                //command.ExecuteNonQuery();
                                //command.CommandText = "update sobitie set smena=" + comboBox2.SelectedIndex + " where id=" + listView1.Items[(int)same[i]].SubItems[7];
                                //command.ExecuteNonQuery();
                                listView1.Items[(int)same[i]].SubItems[6].Text = ctena.ToString();
                                listView1.Items[(int)same[i]].SubItems[8].Text = comboBox2.Text;
                            }
                            //conn.Close();
                            textBox1.SelectAll();
                        }

                    }
                }
                catch (System.Exception)
                {
                    MessageBox.Show("Неверно заполнены поля!");
                    return;
                }
            }
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            try
            {
                textBox2.Text = listView1.SelectedItems[0].SubItems[0].Text;
                textBox3.Text = listView1.SelectedItems[0].SubItems[5].Text;
            }
            catch (System.Exception)
            {
                return;
            }
        }
    }
}