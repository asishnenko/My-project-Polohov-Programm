using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Polohov
{
    public partial class AddProd : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        AddProd1 addka;
        AddProd1 adpr1;
        public AddProd(SqlConnection conn)
        {
            InitializeComponent();
            if (Form1.price == false)
            {
                label6.Visible = false;
                textBox3.Visible = false;
            }
            if (Form1.price == true)
            {
                label6.Visible = true;
                textBox3.Visible = true;
            }
            this.conn = conn;
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select name from prodykt";
            r = command.ExecuteReader();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    comboBox1.Items.Add((string)r[0]);
                }
            }
            conn.Close();

            conn.Open();
            command.CommandText = "select name from kAgent";
            r = command.ExecuteReader();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    comboBox3.Items.Add((string)r[0]);
                }
            }
            conn.Close();

            //comboBox2.Items.Add("Весь склад");
            comboBox2.Items.Add("Готовая продукция");
            comboBox2.Items.Add("Склад производства");
            //conn.Open();
            //command.CommandText = "select name from state";
            //r = command.ExecuteReader();
            //if (r.HasRows == true)
            //{
            //    while (r.Read() == true)
            //    {
            //        if ((string)r[0] == "Продано" || (string)r[0] == "В работе") continue;
            //        comboBox2.Items.Add((string)r[0]);
            //    }
            //}
            //conn.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "" && comboBox1.Text != "" && dateTimePicker1.Text != "" && comboBox2.Text != "")
            {
                try
                {
                    Convert.ToDecimal(textBox3.Text);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("В поле цена введите цифры!");
                    return;
                }
                if (dateTimePicker1.Value > DateTime.Now)
                {
                    MessageBox.Show("Неверно указана дата!");
                    return;
                }
                conn.Open();
                command.CommandText = "select id from partiya where name='" + textBox1.Text + "'";
                r = command.ExecuteReader();
                if (r.HasRows == false)
                {
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from kAgent where name ='" + comboBox3.Text + "'";
                    int kagent = (int)command.ExecuteScalar();
                    conn.Close();
                    conn.Open();
                    command.CommandText = "insert into partiya(name) values('" + textBox1.Text.Replace(',','.') + "')";
                    command.ExecuteNonQuery();
                    
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from partiya where name ='" + textBox1.Text.Replace(',', '.') + "'";
                    int partiya = (int)command.ExecuteScalar();
                    conn.Close();
                    //вставляем в событие партии о ее начале
                    conn.Open();
                    command.CommandText = "insert into sobitiepartii values("+partiya+",1,'"+dateTimePicker1.Text+"')";
                    command.ExecuteNonQuery();
                    conn.Close();
                    
                    conn.Open();
                    command.CommandText = "select id from prodykt where name ='"+comboBox1.Text+"'";
                    int prodykt = (int)command.ExecuteScalar();
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from state where name ='" + comboBox2.Text + "'";
                    int state = (int)command.ExecuteScalar();
                    conn.Close();
                    conn.Open();
                    string ostatok= textBox2.Text.Replace(',', '.');
                    string price = textBox3.Text.Replace(',', '.');
                    string s1 = "";
                    string s2 = "";
                    if (textBox3.Text != "")
                    {
                        s1 = ",tcena";
                        s2 = "," + price;
                    }
                    //MessageBox.Show(textBox2.Text);
                    command.CommandText = "insert into vessklad (idpartiya,  idprodykt, nachves, ostatok, data, idstate, idsost, recordtime"+s1+") values (" + partiya + "," + prodykt + ",'" + ostatok + "','" + ostatok + "','" + dateTimePicker1.Text + "',"+state+",1,'" + DateTime.Now.ToString() + "'"+s2+")";
                    if (command.ExecuteNonQuery() != 0)
                    {
                        conn.Close();
                        conn.Open();
                        command.CommandText = "select max(id) from vessklad";
                        int id = (int)command.ExecuteScalar();
                        conn.Close();



                        conn.Open();
                        s1 = "";
                        s2 = "";
                        if (textBox3.Text != "")
                        {
                            s1 = ",price";
                            s2 = "," + price;
                        }
                        command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime" + s1 + ",idkagent) values(" + id + "," + ostatok + ",1,1,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "'" + s2 + ","+kagent+")";
                        command.ExecuteNonQuery();
                        conn.Close();
                        if (MessageBox.Show("Добавлено!\nДобавить еще?", "Ok", MessageBoxButtons.YesNo) == DialogResult.No)
                        {                           
                            this.Close();
                        }
                        else
                        {                           
                            textBox1.Text = "";
                            textBox2.Text = "";
                        }
                    }
                    else
                    {
                        conn.Close();
                        MessageBox.Show("Ошибка! Добавление нового продукта не произошло!");
                        return;
                    }
                    
                    //else
                    //{
                    //    MessageBox.Show("Запись наименования партии не произошло!");
                    //    conn.Close();
                    //    return;
                    //}
                }
                else
                {
                    MessageBox.Show("Указанная партия уже существует.\nНеобходимо ввести наименование новой партии.");
                    conn.Close();
                    return;
                }
            }
            else
            {
                MessageBox.Show("Заполните правильно все поля!");
                conn.Close();
                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            addka = new AddProd1(conn, "Контрагент");
            addka.Show();
            addka.FormClosed += new FormClosedEventHandler(addka_FormClosed);
        }
        void addka_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox3.Items.Clear();
            conn.Open();
            command.CommandText = "select name from kAgent";
            r = command.ExecuteReader();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    comboBox3.Items.Add((string)r[0]);
                }
            }
            conn.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            adpr1 = new AddProd1(conn, "Новый продукт");
            adpr1.Show();
            adpr1.FormClosed += new FormClosedEventHandler(adpr1_FormClosed);
        }
        void adpr1_FormClosed(object sender, FormClosedEventArgs e)
        {
            conn.Open();
            comboBox1.Items.Clear();
            command.CommandText = "select name from prodykt";
            r = command.ExecuteReader();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    comboBox1.Items.Add((string)r[0]);
                }
            }
            conn.Close();
        }
    }
}