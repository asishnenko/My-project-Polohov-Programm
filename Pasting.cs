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
    public partial class Pasting : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        int id;
        //int ostatok;
        //ArrayList ar;
        public Pasting(SqlConnection conn, int id)
        {
            InitializeComponent();
            this.conn = conn;
            this.id = id;
            if (Form1.price == true) { label4.Visible = true; textBox2.Visible = true; }
            if (Form1.price == false) { label4.Visible = false; textBox2.Visible = false; }
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select partiya.name,prodykt.name,vessklad.ostatok,state.name from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and state.id=vessklad.idstate and vessklad.id="+id;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    textBox3.Text = r[0].ToString();
                    textBox4.Text = r[1].ToString();
                    textBox1.Text = r[2].ToString();
                    textBox2.Text = r[3].ToString();
                    //ostatok = Convert.ToInt32(r[2].ToString());

                }
            }
            conn.Close();
            if (textBox2.Text == "Готовая продукция") comboBox1.Items.Add("Склад производства");
            else comboBox1.Items.Add("Готовая продукция");
            //comboBox1.Items.Add("Готовая продукция");
            //comboBox1.Items.Add("Склад производства");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > DateTime.Now)
            {
                MessageBox.Show("Неверно указана дата!");
                return;
            }
            conn.Close();
            conn.Open();
            if (comboBox1.Text != textBox2.Text)// проверить работу перемещения
            {
                command.CommandText = "select id from state where name='"+comboBox1.Text+"'";
                int idstate=(int)command.ExecuteScalar();
                command.CommandText = "update vessklad set idstate="+idstate+" where id="+id;
                command.ExecuteNonQuery();
                int s=0;
                int s1=0;
                if(textBox2.Text=="Готовая продукция")s=7;
                if(textBox2.Text=="Склад производства")s=2;
                if(textBox2.Text=="Исходное сырье")s=1;
                if (comboBox1.Text == "Готовая продукция") s1 = 7;
                if (comboBox1.Text == "Склад производства") s1 = 2;
                if (comboBox1.Text == "Исходное сырье") s1 = 1;
                command.CommandText = "insert into sobitie (idsklad,ves,iddvigfrom,iddvig,data,recordtime) values (" + id + "," + textBox1.Text.Replace(',', '.') + "," + s.ToString() + "," + s1.ToString() + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";// дописать добавление в событие!!!!
                command.ExecuteNonQuery();
                conn.Close();
                if (comboBox1.Text == "Склад производства")
                {
                    conn.Open();
                    command.CommandText = "update partiya set konetc=1 where name=(select partiya.name from partiya,vessklad where vessklad.idpartiya=partiya.id and vessklad.id="+id+")";
                    command.ExecuteNonQuery();
                    conn.Close();
                }
                MessageBox.Show("Перемещено успешно!");
            }
            else
            {
                MessageBox.Show("Продукт уже находится на выбранном складе!");
                return;
            }
            conn.Close();
            //if (textBox1.Text != "")
            //{
            //    if (Convert.ToInt32(textBox1.Text) <= ostatok)
            //    {
            //        if (Convert.ToInt32(textBox1.Text) < ostatok)
            //        {
            //            ar = new ArrayList();
            //            conn.Open();
            //            command.CommandText = "update vessklad set ostatok="+(ostatok-Convert.ToInt32(textBox1.Text))+" where id="+id;
            //            command.ExecuteNonQuery();
            //            conn.Close();
            //            conn.Open();
            //            command.CommandText = "select idpartiya, idprodykt, nachves, ostatok, idrabotnik, idstanok, zatracheno, data,idstate,idsost from vessklad where id="+id;
            //            r = command.ExecuteReader();
            //            if (r.HasRows)
            //            {                            
            //                while (r.Read())
            //                {
            //                    for (int i = 0; i < 10; i++)
            //                    {                                    
            //                        ar.Add(r[i]);
            //                    }                                    
            //                }
            //            }
            //            conn.Close();
            //            conn.Open();
            //            command.CommandText = "insert into vessklad values(" + ar[0].ToString() + "," + ar[1].ToString() + "," + ar[2].ToString() + "," + textBox1.Text + "," + ar[4].ToString() + "," + ar[5].ToString() + "," + ar[6].ToString() + ",'" + dateTimePicker1.Text + "',3," + ar[9].ToString() + ",'"+DateTime.Now.ToString()+"')";
            //            command.ExecuteNonQuery();
            //            conn.Close();
            //            string s="";
            //            string s1="";
            //            if(Form1.price==true&&textBox2.Text!=""){s=" ,price";s1=" ,"+textBox2.Text;}
            //            conn.Open();
            //            command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime"+s+") values("+id+","+textBox1.Text+",3,2,'"+dateTimePicker1.Text+"','"+DateTime.Now.ToString()+"'"+s1+")";
            //            command.ExecuteNonQuery();
            //            conn.Close();

            //        }
            //        if (Convert.ToInt32(textBox1.Text) == ostatok)
            //        {
            //            conn.Open();
            //            command.CommandText = "update vessklad set ostatok=" + 0 + ",idstate=3 where id=" + id;
            //            command.ExecuteNonQuery();
            //            conn.Close();

            //            string s = "";
            //            string s1 = "";
            //            if (Form1.price == true) { s = " ,price"; s1 = " ," + textBox2.Text; }
            //            conn.Open();
            //            command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime" + s + ") values(" + id + "," + textBox1.Text + ",3,2,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "'" + s1 + ")";
            //            command.ExecuteNonQuery();
            //            conn.Close();
            //        }
            //    }
            //    else { MessageBox.Show("Указанный вес превышает возможный"); return; }
            //    MessageBox.Show("Продано успешно");
            //    this.Close();
            //}
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}