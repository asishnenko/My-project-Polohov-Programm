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
    public partial class AddProd1 : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        public AddProd1(SqlConnection conn, string name)
        {
            InitializeComponent();
            this.conn = conn;
            command.Connection = conn;
            this.Text = name;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Contains(",")) { MessageBox.Show("Строка содержит недопустимый символ"); return; }
            if (textBox1.Text != "")
            {
                conn.Open();
                if (this.Text == "Новый продукт") command.CommandText = "select id from prodykt where name='" + textBox1.Text + "'";
                if (this.Text == "Новое оборудование") command.CommandText = "select id from stanok where name='" + textBox1.Text + "'";
                if (this.Text == "Новая партия") command.CommandText = "select id from partiya where name='" + textBox1.Text + "'";
                if (this.Text == "Вид отходов") command.CommandText = "select id from tipothoda where name='" + textBox1.Text + "'";
                if (this.Text == "Контрагент") command.CommandText = "select id from kAgent where name='" + textBox1.Text + "'";
                //if (this.Text == "Заполнение партии") command.CommandText = "select id from partiya where name='"+textBox1.Text+"'";
                r = command.ExecuteReader();
                if (r.HasRows == false)
                {
                    conn.Close();
                    conn.Open();
                    if (this.Text == "Новый продукт")command.CommandText = "insert into prodykt values('" + textBox1.Text + "')";
                    if (this.Text == "Новое оборудование") command.CommandText = "insert into stanok values('" + textBox1.Text + "')";
                    if (this.Text == "Новая партия") { command.CommandText = "insert into partiya(name) values('" + textBox1.Text + "')"; command.ExecuteNonQuery(); int idp = 0; command.CommandText = "select id from partiya where name='" + textBox1.Text + "'"; idp = (int)command.ExecuteScalar(); command.CommandText = "insert into sobitiepartii values("+idp+",1,'" + DateTime.Now.ToShortDateString() + "')"; }
                    if (this.Text == "Вид отходов") command.CommandText = "insert into tipothoda values('" + textBox1.Text + "')";
                    if (this.Text == "Контрагент") command.CommandText = "insert into kAgent values('" + textBox1.Text + "')";
                    //if (this.Text == "Заполнение партии") command.CommandText = "insert into ";
                    if (command.ExecuteNonQuery() != 0)
                    {
                        conn.Close();
                        if (MessageBox.Show("Запись добавлена.\nДобавить еще?", "Ok", MessageBoxButtons.YesNo) == DialogResult.No)
                        {
                            this.Close();
                        }
                        else
                        {
                            textBox1.Text = "";
                            this.Refresh();
                        }
                    }
                    else
                    {
                        conn.Close();
                        MessageBox.Show("Добавление в базу не произошло!");
                        return;
                    }
                }
                else
                {
                    conn.Close();
                    MessageBox.Show("Такое наименование уже существует!");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }
        }
    }
}