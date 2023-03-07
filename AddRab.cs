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

    public partial class AddRab : Form
    {
        SqlConnection conn;
        SqlCommand comand = new SqlCommand();
        public AddRab(SqlConnection conn)
        {
            InitializeComponent();
            this.conn = conn;
            comand.Connection = conn;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Заполните имя и фамилию!");
                return;
            }
            comand.CommandText = "select id from rabotnik where name='"+textBox1.Text+"' and surname='"+textBox2.Text+"'";
            conn.Open();
            object ob=comand.ExecuteScalar();
            conn.Close();
            if (ob == null)
            {
                comand.CommandText = "insert into rabotnik values ('" + textBox1.Text + "','" + textBox2.Text + "',1,'"+System.DateTime.Now.ToString()+"')";
                conn.Open();
                int ss=comand.ExecuteNonQuery();
                conn.Close();
                if (MessageBox.Show("Добавлено записей: " + ss.ToString() + "\nДобавить еще?", "Добавлено", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    this.Close();
                }
                else
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox1.Select();
                    this.Refresh();
                }
                //textBox1.Text = "";
                //textBox2.Text = "";
                //textBox1.Select();
            }
            else
            {
                MessageBox.Show("Такая запись уже существует! "+'\n'+"Если этот работник был уволен восстановите его.");
                return;
            }
        }

        private void AddRab_Load(object sender, EventArgs e)
        {

        }
    }
}