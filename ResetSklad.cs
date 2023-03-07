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
    public partial class ResetSklad : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        public ResetSklad(SqlConnection conn)
        {
            this.conn = conn;            
            command.Connection = conn;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "19791021") { MessageBox.Show("Введите правильный пароль!"); return; }
            ArrayList prt = new ArrayList();
            conn.Open();
            command.CommandText = "select name from partiya where name!='Не определено'";
            r = command.ExecuteReader();
            if (r.HasRows)
                while (r.Read())
                {
                    prt.Add(r[0].ToString());
                }
            conn.Close();

            for (int i = 0; i < prt.Count; i++)
            {
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,partiya where partiya.name='" + prt[i].ToString() + "' and partiya.id=vessklad.idpartiya";
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
                    command.CommandText = "select id from proizv where idsirie=" + idsklad[n].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
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
                        command.CommandText = "delete from sobitie where idproizv=" + idproizv[k].ToString();
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
                command.CommandText = "delete from partiya where name='" + prt[i].ToString() + "'";
                command.ExecuteNonQuery();
                conn.Close();

            }
            MessageBox.Show("Успешно удалено!");
            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "") { MessageBox.Show("Введите наименование новой базы"); return; }
            if (textBox1.Text != "19791021") { MessageBox.Show("Введите правильный пароль!"); return; }
            conn.Open();
            string spath="";
            StreamReader sreader;
            FileInfo path = new FileInfo("path.txt");
            if (!path.Exists)
            {
                MessageBox.Show("Не найден файл с путем к базе!");
                return;
            }
            sreader = path.OpenText();
            spath = sreader.ReadLine();
            sreader.Close();
            //string tyu = conn.Database;
            //int start = tyu.IndexOf(":\\");
            //int end = tyu.IndexOf(".mdf") - tyu.IndexOf(":\\");
            //string basename = tyu.Substring(start, end);
            string basename = conn.Database ;
            command.CommandText = "CREATE DATABASE " + textBox2.Text + " ON ( NAME = " + textBox2.Text + "_dat,FILENAME = '" + spath + textBox2.Text + "dat.mdf',SIZE = 10,MAXSIZE = 50,FILEGROWTH = 5 ) LOG ON ( NAME = "+textBox2.Text+"_log, FILENAME = '" + spath + textBox2.Text + "log.ldf', SIZE = 5MB, MAXSIZE = 25MB, FILEGROWTH = 5MB )";
            try
            {
                command.ExecuteNonQuery();
            }
            catch (SqlException err)
            {
                MessageBox.Show(err.Message);
                conn.Close();
                return;
            }
            conn.Close();
            conn.ChangeDatabase(textBox2.Text);
            conn.Open();
            command.CommandText = "use " + textBox2.Text;
            command.ExecuteNonQuery();
            StreamWriter swriter;
            FileInfo con = new FileInfo("conn.txt");
            sreader = con.OpenText();
            
            string ttt = sreader.ReadToEnd().Replace(basename, conn.Database);
            sreader.Close();
            swriter = con.CreateText();
            swriter.Write(ttt);
            swriter.Close();                                  
            MessageBox.Show("Успешно создана новая база "+conn.Database);
            conn.Close();
            Close();
        }
    }
}
