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
    public partial class Vhod : Form
    {
        SqlConnection conn;
        SqlCommand comand = new SqlCommand();
        SqlDataReader r;
        public Vhod(SqlConnection conn)
        {
            InitializeComponent();
            this.conn = conn;
            comand.Connection = conn;
            conn.Open();
            comand.CommandText = "select login from users";
            r = comand.ExecuteReader();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    comboBox1.Items.Add((string)r[0]);
                }
            }
            else
            {
                MessageBox.Show("В базе нет ни одного пользователя. Обратитесь к создателю программы.");
                Application.Exit();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            conn.Close();
            if (comboBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Заполните поля!");
                return;
            }
            conn.Open();
            comand.CommandText = "select name from users,prava where login='"+comboBox1.Text+"' and pass='"+textBox2.Text+"' and prava.id=users.idprava";
            r = comand.ExecuteReader();
            if (r.HasRows == true)
            {
                r.Read();
                if ((string)r[0] == "admin") { Form1.prava = true; }
                if ((string)r[0] == "rabotnik") { Form1.prava = false; }
                this.Close();
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль");
                return;
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}