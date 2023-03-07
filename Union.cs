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
    public partial class Union : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        int id;
        public Union(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;
            InitializeComponent();
            
            listView1.Columns.Add("Партия",50);
            listView1.Columns.Add("Продукт", 150);
            listView1.Columns.Add("Вес", 50);
            listView1.Columns.Add("Склад", 150);
            listView1.Columns.Add("id", 0);

            listView2.Columns.Add("Партия", 50);
            listView2.Columns.Add("Продукт", 150);
            listView2.Columns.Add("Вес", 50);
            listView2.Columns.Add("Склад", 150);

            conn.Open();
            command.CommandText = "select name from prodykt";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox1.Items.Add(r[0].ToString());
                }
            }
            conn.Close();
            comboBox2.Items.Add("Готовая продукция");
            comboBox2.Items.Add("Склад производства");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "") { MessageBox.Show("Введите название новой партии!"); return; }
            if (comboBox1.Text == "") { MessageBox.Show("Выберите наименование нового продукта!"); return; }
            if (comboBox2.Text == "") { MessageBox.Show("Выберите склад!"); return; }
            if (MessageBox.Show("Выберите продукт в главном окне") == DialogResult.OK)
            {
                this.Hide();
            }
        }
        public DialogResult ShowDialog(int id)
        {
            this.id = id;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (Convert.ToInt32(listView1.Items[i].SubItems[4].Text) == id)
                {
                    MessageBox.Show("Такой продукт уже добавлен. \nВыберите другой или добавьте на склад новый.");
                    return this.ShowDialog();
                }
            }
            button1.Text = "Еще один";
            conn.Open();
            command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok, state.name from vessklad, partiya, prodykt, state where state.id=vessklad.idstate and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + id.ToString();
            r = command.ExecuteReader();
            ListViewItem lvi;
            string[] s = new string[r.FieldCount + 1];
            if (r.HasRows)
            {
                while (r.Read())
                {
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        s[i] = r[i].ToString();
                    }
                    s[r.FieldCount] = id.ToString();
                    //s[r.FieldCount + 1] = "";
                    lvi = new ListViewItem(s);
                    listView1.Items.Add(lvi);
                }
            }
            listView2.Items.Clear();
            decimal sum = 0;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                sum += Convert.ToDecimal(listView1.Items[i].SubItems[2].Text);
            }
            s = new string[4];
            s[0] = textBox1.Text;
            s[1] = comboBox1.Text;
            s[2] = sum.ToString();
            s[3] = comboBox2.Text;
            lvi = new ListViewItem(s);
            listView2.Items.Add(lvi);

            conn.Close();
            return this.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) { MessageBox.Show("Выберите строку для удаления!"); return; }
            int selind = listView1.SelectedIndices[0];
            listView1.Items.RemoveAt(selind);
            listView1.Refresh();
            
            listView2.Items.Clear();
            decimal sum = 0;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                sum += Convert.ToDecimal(listView1.Items[i].SubItems[2].Text);
            }
            string[] s = new string[4];
            s[0] = textBox1.Text;
            s[1] = comboBox1.Text;
            s[2] = sum.ToString();
            s[3] = comboBox2.Text;
            ListViewItem lvi = new ListViewItem(s);
            listView2.Items.Add(lvi);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //dateTimePicker1.Value = DateTime.Now;
            //MessageBox.Show(dateTimePicker1.Value.ToString());
            if (dateTimePicker1.Value > DateTime.Now)
            {
                MessageBox.Show("Неверно указана дата!");
                return;
            }
            if (listView1.Items.Count == 0 || listView2.Items.Count == 0) return;
            if (textBox1.Text.Contains(",")) { MessageBox.Show("Недопустимый символ в наименовании партии!"); return; }
            conn.Open();
            command.CommandText = "select id from partiya where name='" + textBox1.Text + "'";
            r = command.ExecuteReader();
            if (r.HasRows) { MessageBox.Show("Такая партия уже существует. Дайте другое название."); conn.Close(); return; }
            conn.Close();
            
            decimal sum = 0;
            conn.Open();
            int permax = 0;
            command.CommandText = "select max(id) from proizv";
            permax = (int)command.ExecuteScalar();
            permax++;
            command.CommandText = "insert into proizv(id) values ("+permax.ToString()+")";
            command.ExecuteNonQuery();
            
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                int idprodMax = i + 1;
                command.CommandText = "update vessklad set ostatok=0 where id="+listView1.Items[i].SubItems[4].Text;
                command.ExecuteNonQuery();

                try
                {

                    command.CommandText = "update proizv set idprodykt" + idprodMax + "=" + listView1.Items[i].SubItems[4].Text + " where id=" + permax.ToString();
                    command.ExecuteNonQuery();
                }
                catch (System.Exception)
                {
                    command.CommandText = "ALTER TABLE proizv ADD idprodykt" + idprodMax + " int CONSTRAINT fk_prodykt" + idprodMax + "_element FOREIGN KEY  REFERENCES vessklad(id)on delete no action on update no action";
                    command.ExecuteNonQuery();
                    command.CommandText = "update proizv set idprodykt" + idprodMax + "=" + listView1.Items[i].SubItems[4].Text + " where id=" + permax.ToString();
                    command.ExecuteNonQuery();
                }

                //command.CommandText = "update proizv set idprodykt"+(i+1)+"="+listView1.Items[i].SubItems[4].Text+" where id="+permax.ToString();
                //command.ExecuteNonQuery();
                command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values(" + listView1.Items[i].SubItems[4].Text + ","+listView1.Items[i].SubItems[2].Text.Replace(',','.')+",5,5,2,"+permax+",'"+dateTimePicker1.Text+"','"+DateTime.Now.ToString()+"')";
                command.ExecuteNonQuery();
                sum += Convert.ToDecimal(listView1.Items[i].SubItems[2].Text);
            }
            conn.Close();
            conn.Open();
            command.CommandText = "select id from partiya where name='"+textBox1.Text+"'";
            r = command.ExecuteReader();
            if (r.HasRows) { MessageBox.Show("Такая партия уже существует. Дайте другое название."); conn.Close(); return; }
            conn.Close();
            conn.Open();
            command.CommandText = "insert into partiya(name,sbor) values ('"+textBox1.Text+"',1)";
            command.ExecuteNonQuery();
            command.CommandText = "select max(id) from partiya";
            int idpartiya = (int)command.ExecuteScalar();
            //вставляем в событие партий
            //conn.Open();
            command.CommandText = "insert into sobitiepartii values(" + idpartiya + ",1,'" + dateTimePicker1.Value.ToShortDateString() + "')";
            command.ExecuteNonQuery();
            //conn.Close();

            command.CommandText = "select id from prodykt where name='"+comboBox1.Text+"'";
            int idprodykt = (int)command.ExecuteScalar();
            command.CommandText = "select id from state where name='"+comboBox2.Text+"'";
            int idstate = (int)command.ExecuteScalar();
            command.CommandText = "insert into vessklad(idpartiya,idprodykt,nachves,ostatok,data,idstate,idsost,recordtime) values(" + idpartiya.ToString() + "," + idprodykt.ToString() + "," + sum.ToString().Replace(',', '.') + "," + sum.ToString().Replace(',', '.') + ",'" + dateTimePicker1.Text + "'," + idstate.ToString() + ",1,'" + DateTime.Now.ToString() + "')";
            command.ExecuteNonQuery();
            command.CommandText = "select max(id) from vessklad";
            int idnew = (int)command.ExecuteScalar();
            command.CommandText = "update proizv set idsirie=" + idnew.ToString() + " where id=" + permax.ToString();
            command.ExecuteNonQuery();
            //MessageBox.Show(dateTimePicker1.Text);
            command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values(" + idnew.ToString() + "," + sum.ToString().Replace(',', '.') + ",5,1,1," + permax + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
            int rez = 0;
            rez=command.ExecuteNonQuery();
            conn.Close();
            if ( rez!= 0)
            {
                if (MessageBox.Show("Добавлено успешно!") == DialogResult.OK)
                {
                    this.Close();
                }
            }
        }
    }
}