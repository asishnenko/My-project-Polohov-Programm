using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;

namespace Polohov
{
    public partial class deletepartiya : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        string parol;
        public deletepartiya(SqlConnection conn)
        {
            InitializeComponent();
            this.conn = conn;
            parol = "";
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select name from partiya where name!='Не определено'";
            r = command.ExecuteReader();
            if(r.HasRows)
                while (r.Read())
                {
                    listBox1.Items.Add(r[0].ToString());
                }
            conn.Close();

            StreamReader sreader;
            FileInfo pwd = new FileInfo("oll");
            if (!pwd.Exists)
            {
                MessageBox.Show("Не найден файл с паролем. Удалить невозможно.");
                button1.Enabled = false;
            }
            sreader = pwd.OpenText();
            parol = sreader.ReadLine();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != parol) { MessageBox.Show("Неверный пароль!"); return; }
            if (MessageBox.Show("Будет удалена вся информация о партии безвозвратно!", "Внимание!", MessageBoxButtons.OKCancel) == DialogResult.Cancel) return;
            for (int i = 0; i < listBox1.SelectedItems.Count; i++)
            {
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,partiya where partiya.name='"+listBox1.SelectedItems[i].ToString()+"' and partiya.id=vessklad.idpartiya";
                ArrayList idsklad = new ArrayList();
                r = command.ExecuteReader();
                if (r.HasRows)
                    while (r.Read())
                        idsklad.Add(r[0].ToString());
                conn.Close();

              
                ArrayList idproizv = new ArrayList();
                ArrayList idsobitie = new ArrayList();
                for (int n = 0; n < idsklad.Count; n++)
                {
                    conn.Open();
                    command.CommandText = "select id from proizv where idsirie="+idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if(r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt1=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt2=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt3=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt4=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt5=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt6=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt7=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt8=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();
                    conn.Open();
                    command.CommandText = "select id from proizv where idprodykt9=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idproizv.Contains(r[0].ToString())) idproizv.Add(r[0].ToString());
                        }
                    conn.Close();

                    

                    conn.Open();
                    command.CommandText = "select id from sobitie where idsklad=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            if (!idsobitie.Contains(r[0].ToString())) idsobitie.Add(r[0].ToString());
                        }
                    conn.Close();
                    //удалние из события
                    conn.Open();                  
                    command.CommandText = "delete from sobitie where idsklad=" + idsklad[n].ToString();
                    command.ExecuteNonQuery();                  
                    conn.Close();  
                }

                conn.Open();
                for (int k = 0; k < idproizv.Count; k++)
                {
                    command.CommandText = "delete from othodi where idproizv=" + idproizv[k].ToString();
                    command.ExecuteNonQuery();
                    command.CommandText = "delete from sirie where idproizv=" + idproizv[k].ToString();
                    command.ExecuteNonQuery();
                }
                conn.Close();

                conn.Open();
                for (int k = 0; k < idsobitie.Count; k++)
                {
                    command.CommandText = "delete from podsobniki where idsobitie=" + idsobitie[k].ToString();
                    command.ExecuteNonQuery();
                    //command.CommandText = "delete from sirie where idproizv=" + idsobitie[k].ToString();
                    //command.ExecuteNonQuery();
                }
                conn.Close();


                conn.Open();
                for (int k = 0; k < idproizv.Count; k++)
                {
                    command.CommandText = "delete from proizv where id=" + idproizv[k].ToString();
                    try
                    {
                        command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        command.CommandText = "delete from sobitie where idproizv="+idproizv[k].ToString();
                        command.ExecuteNonQuery();
                        k--;
                    }
                }
                conn.Close();

                conn.Open();
                for (int k = 0; k < idsklad.Count; k++)
                {
                    command.CommandText = "delete from vessklad where id=" + idsklad[k].ToString();
                    command.ExecuteNonQuery();
                }
                conn.Close();

                conn.Open();
                command.CommandText = "delete from partiya where name='" + listBox1.SelectedItems[i].ToString() + "'";
                command.ExecuteNonQuery();
                conn.Close();
            }
            listBox1.Items.Clear();
            conn.Open();
            command.CommandText = "select name from partiya where name!='Не определено'";
            r = command.ExecuteReader();
            if (r.HasRows)
                while (r.Read())
                {
                    listBox1.Items.Add(r[0].ToString());
                }
            conn.Close();
        }
    }
}
