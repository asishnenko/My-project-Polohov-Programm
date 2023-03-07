using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace Polohov
{
    public partial class delete : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        int id;
        string parol;
        decimal ostatok;
        public delete(SqlConnection conn, int id)
        {
            InitializeComponent();
            this.conn = conn;
            this.id = id;
            parol = "";
            ostatok = 0;
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select partiya.name,prodykt.name,vessklad.ostatok from partiya,prodykt,vessklad where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + id;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    textBox2.Text = r[0].ToString() + " " + r[1].ToString() + " " + r[2].ToString()+"кг";
                    textBox3.Text = r[2].ToString();
                    ostatok = Convert.ToDecimal(r[2].ToString().Replace(".",","));
                }
            }
            conn.Close();
            StreamReader sreader;
            FileInfo pwd = new FileInfo("oll");
            if (!pwd.Exists)
            {
                MessageBox.Show("Не найден файл с паролем. Удалить невозможно.");
                button2.Enabled = false;
            }
            sreader = pwd.OpenText();
            parol = sreader.ReadLine();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != parol) { MessageBox.Show("Неверный пароль!"); return; }
            else
            {
                decimal ves=0;
                try
                {
                    ves = Convert.ToDecimal(textBox3.Text.Replace(".",","));
                }
                catch (System.Exception)
                {
                    MessageBox.Show("В поле (Вес) укажите только цифры!");
                    return;
                }
                if (ves > ostatok || ves <= 0) { MessageBox.Show("Неверно указан вес!"); return; }
                if (ves <= ostatok)
                {
                    conn.Open();
                    int okon = 0;
                    okon = (int)ostatok - (int)ves;
                    command.CommandText = "update vessklad set ostatok="+okon+" where id="+id;
                    command.ExecuteNonQuery();
                    command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime) values(" + id + "," + textBox3.Text.Replace(',', '.') + ",8,2,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                    command.ExecuteNonQuery();
                    conn.Close();
                }
                if (ves == ostatok)
                {
                    conn.Open();
                    command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime) values(" + id + "," + textBox3.Text.Replace(',', '.') + ",4,2,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                    command.CommandText = "update vessklad set idstate=6 where id=" + id;
                    command.ExecuteNonQuery();
                    conn.Close();
                }
                MessageBox.Show("Удалено успешно!");
                this.Close();
            }
        }
    }
}