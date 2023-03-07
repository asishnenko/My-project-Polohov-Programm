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
    public partial class Zapolnenie : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        public int id;
        public Zapolnenie(SqlConnection conn)
        {
            InitializeComponent();
            id = 0;
            this.conn = conn;
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select name from partiya where name!='Не определено'";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                comboBox1.Items.Add((string)r[0]);
            }
            conn.Close();

            conn.Open();
            command.CommandText = "select name from prodykt";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                comboBox2.Items.Add((string)r[0]);
            }
            conn.Close();
     
            //conn.Open();
            //command.CommandText = "select name from state";
            //r = command.ExecuteReader();
            //while (r.Read() == true)
            //{
            //    if ((string)r[0] == "Продано" || (string)r[0] == "В работе") break;
            //    comboBox5.Items.Add((string)r[0]);
            //}
            //conn.Close();
            comboBox5.Items.Add("Склад производства");
            comboBox5.Items.Add("Готовая продукция");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Создавать новую партию только в случае отсутствия ее в списке!","Важно!",MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                AddProd1 adrp = new AddProd1(conn, "Новая партия");
                adrp.Show();
                adrp.FormClosed += new FormClosedEventHandler(adrp_FormClosed);
            }
        }

        void adrp_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1.Items.Clear();
            conn.Open();
            command.CommandText = "update partiya set ostatki=1 where id=(select max(id) from partiya)";
            command.ExecuteNonQuery();
            command.CommandText = "select name from partiya where name!='Не определено'";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                comboBox1.Items.Add((string)r[0]);
            }
            conn.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "" && dateTimePicker1.Text != "" && comboBox5.Text != ""&&textBox1.Text!="")
            {
                conn.Open();
                command.CommandText = "select id from partiya where name='" + comboBox1.Text + "'";
                int partiya = (int)command.ExecuteScalar();
                conn.Close();

                conn.Open();
                command.CommandText = "select id from prodykt where name='" + comboBox2.Text + "'";
                int prodykt = (int)command.ExecuteScalar();
                conn.Close();
                conn.Open();
                command.CommandText = "select id from state where name='" + comboBox5.Text + "'";
                int state = (int)command.ExecuteScalar();
                conn.Close();

                conn.Open();
                command.CommandText = "select id from vessklad where idpartiya=(select id from partiya where name='"+comboBox1.Text+"') and idprodykt=(select id from prodykt where name='"+comboBox2.Text+"') and idstate=(select id from state where name='"+comboBox5.Text+"')";
                int idd=0;
                if (command.ExecuteScalar() != null) idd = (int)command.ExecuteScalar();
                conn.Close();
                if (idd != 0)
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Продукт этой партии уже существует.\nВыберите 'Ok' чтобы добавить вес к уже существующей записи\nНажмите 'Cancel' чтобы изменить вводимые параметры", "Предупреждение", MessageBoxButtons.OKCancel);
                    if (dr == DialogResult.OK)
                    {
                        conn.Open();
                        command.CommandText = "select ostatok from vessklad where id=" + idd;
                        decimal ostatok = Convert.ToDecimal(command.ExecuteScalar().ToString().Replace('.',','));
                        conn.Close();
                        conn.Open();
                        ostatok+=Convert.ToDecimal(textBox1.Text.Replace('.',','));
                        command.CommandText = "update vessklad set ostatok=" + ostatok.ToString().Replace(',','.') + " where id=" + idd;
                        command.ExecuteNonQuery();
                        conn.Close();

                        conn.Open();
                        command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime) values(" + idd + "," + textBox1.Text + ",9,1,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString()+ "')";
                        command.ExecuteNonQuery();
                        conn.Close();
                        goto m2;
                    }
                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                
                conn.Open();
                command.CommandText = "insert into vessklad(idpartiya,idprodykt,nachves,ostatok,data,idstate,idsost,recordtime,kyski) values(" + partiya + "," + prodykt + "," + textBox1.Text.Replace(",", ".") + "," + textBox1.Text.Replace(",", ".") + ",'" + dateTimePicker1.Text + "'," + state + ",1,'" + DateTime.Today.ToString() + "',1)";
                command.ExecuteNonQuery();
                conn.Close();

                conn.Open();
                command.CommandText = "select max(id) from vessklad";
                id = (int)command.ExecuteScalar();
                conn.Close();

                conn.Open();
                command.CommandText = "insert into sobitie(idsklad,ves,iddvig,idbalans,data,recordtime) values(" + id + "," + textBox1.Text.Replace(",", ".") + ",9,1,'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                command.ExecuteNonQuery();
                conn.Close();

            m2:if (MessageBox.Show("Запись добавлена.\nДобавить еще?", "Ok", MessageBoxButtons.YesNo) == DialogResult.No)
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
                MessageBox.Show("Заполните обязательные поля!!");
            }
        }
    }
}