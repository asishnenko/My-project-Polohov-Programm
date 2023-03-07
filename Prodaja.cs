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
    public partial class Prodaja : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        int id;
        decimal ostatok;
        ArrayList ar;
        public Prodaja(SqlConnection conn, int id)
        {
            InitializeComponent();
            this.conn = conn;
            this.id = id;
            if (Form1.price == true) { label4.Visible = true; textBox2.Visible = true; }
            if (Form1.price == false) { label4.Visible = false; textBox2.Visible = false; }
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select partiya.name,prodykt.name,vessklad.ostatok from partiya,prodykt,vessklad where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id="+id;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    textBox3.Text = r[0].ToString();
                    textBox4.Text = r[1].ToString();
                    textBox1.Text = r[2].ToString();
                    ostatok = Convert.ToDecimal(r[2].ToString());
                }
            }
            conn.Close();
            conn.Open();
            command.CommandText = "select name from kAgent";
            r = command.ExecuteReader();
            if(r.HasRows)
                while (r.Read())
                {
                    comboBox1.Items.Add(r[0].ToString());
                }
            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                if (dateTimePicker1.Value > DateTime.Now)
                {
                    MessageBox.Show("ƒата указана неверно!");
                    return;
                }
                if (Convert.ToDecimal(textBox1.Text) <= ostatok)
                {
                    //if (Convert.ToInt32(textBox1.Text) < ostatok)
                    
                    //провер€ем на дату!
                    conn.Open();
                    command.CommandText = "select data from vessklad where id="+id;
                    if (Convert.ToDateTime((string)command.ExecuteScalar()) >= dateTimePicker1.Value)
                    {
                        MessageBox.Show("Ќесоответствие даты продажи и даты переработки");
                        return;
                    }
                    conn.Close();
                        //добавление только в событие. элементы склада не раздел€ютс€
                        ar = new ArrayList();
                        conn.Open();
                        decimal rez = ostatok - Convert.ToDecimal(textBox1.Text);
                        command.CommandText = "update vessklad set ostatok="+rez.ToString().Replace(',','.')+" where id="+id;
                        command.ExecuteNonQuery();
                        conn.Close();
                        //conn.Open();
                        //command.CommandText = "select idpartiya, idprodykt, nachves, ostatok, idrabotnik, idstanok, zatracheno, data,idstate,idsost from vessklad where id="+id;
                        //r = command.ExecuteReader();
                        //if (r.HasRows)
                        //{                            
                        //    while (r.Read())
                        //    {
                        //        for (int i = 0; i < 10; i++)
                        //        {                                    
                        //            ar.Add(r[i]);
                        //        }                                    
                        //    }
                        //}
                        //conn.Close();
                        //conn.Open();
                        //command.CommandText = "insert into vessklad values(" + ar[0].ToString() + "," + ar[1].ToString() + "," + ar[2].ToString() + "," + textBox1.Text + "," + ar[4].ToString() + "," + ar[5].ToString() + "," + ar[6].ToString() + ",'" + dateTimePicker1.Text + "',3," + ar[9].ToString() + ",'"+DateTime.Now.ToString()+"')";
                        //command.ExecuteNonQuery();
                        //conn.Close();
                        conn.Open();
                        command.CommandText = "select id from kAgent where name='"+comboBox1.Text+"'";
                        int idkagent = (int)command.ExecuteScalar();
                        conn.Close();
                        string s="";
                        string s1="";
                        if(Form1.price==true&&textBox2.Text!=""){s=" ,price";s1=" ,"+textBox2.Text.Replace(',','.');}
                        conn.Open();
                        command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime"+s+",idkagent) values("+id+","+textBox1.Text.Replace(',','.')+",3,2,'"+dateTimePicker1.Text+"','"+DateTime.Now.ToString()+"'"+s1+","+idkagent.ToString()+")";
                        command.ExecuteNonQuery();
                        
                        conn.Close();
                }


                if (Convert.ToDecimal(textBox1.Text) == ostatok)
                {
                    //conn.Open();
                    //command.CommandText = "update vessklad set ostatok=" + 0 + ",idstate=3 where id=" + id;
                    //command.ExecuteNonQuery();
                    //conn.Close();
                    string s = "";
                    string s1 = "";
                    if (Form1.price == true&&textBox2.Text!="") { s = " ,price"; s1 = " ," + textBox2.Text.Replace(',', '.'); }
                    conn.Open();
                    command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime" + s + ") values(" + id + ",0,4,2,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "'" + s1 + ")";
                    command.ExecuteNonQuery();
                    //command.CommandText = "update vessklad set idstate=3 where id="+id;
                    //command.ExecuteNonQuery();
                    conn.Close();
                }

                if (Convert.ToDecimal(textBox1.Text) > ostatok) { MessageBox.Show("”казанный вес превышает возможный"); return; }
                MessageBox.Show("ѕродано успешно");
                this.Close();
            }
        }
    }
}