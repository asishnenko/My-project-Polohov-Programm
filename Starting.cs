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
    public partial class Starting : Form
    {
        //SqlConnection conn;
        //SqlCommand command = new SqlCommand();
        SqlDataReader r;
        //bool fwd;
        public Starting()
        {
            //this.conn = conn;
            //this.fwd=fwd;
            InitializeComponent();

            StreamReader sreader;
            StreamWriter swriter;
            //parol = "1111";
            FileInfo usr = new FileInfo("users");
            if (!usr.Exists)
            {
                swriter = usr.CreateText();
                swriter.WriteLine("user1");
                swriter.WriteLine("1");
                swriter.WriteLine("user2");
                swriter.WriteLine("2");
                swriter.Close();
            }
            else
            {
                sreader = usr.OpenText();
                for(int i=0;i<2;i++)
                {
                    comboBox1.Items.Add(sreader.ReadLine());
                    sreader.ReadLine();
                }
            }

            //FileInfo sconn = new FileInfo("conn.txt");
            //if (!sconn.Exists)
            //{
            //    MessageBox.Show("Не найден файл со строкой подключения. Дальнейшая работа невозможна");
            //    return;
            //}
            //sreader = sconn.OpenText();
            //conn = new SqlConnection(sreader.ReadLine());
            //try
            //{
            //    conn.Open();
            //}
            //catch (System.Exception)
            //{
            //    MessageBox.Show("Неверные параметры в строке подключения. Программа не сможет работать.");
            //    return;
            //}
            //conn.Close();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            StreamReader sreader;
            StreamReader sreaderconn;
            if (textBox1.Text == "")
            {
                MessageBox.Show("Введите пароль");
                return;
            }
            FileInfo usr = new FileInfo("users");
            sreader = usr.OpenText();
            while (!sreader.EndOfStream)
            {
                if (comboBox1.Text == sreader.ReadLine())
                {
                    if (textBox1.Text == sreader.ReadLine())
                    {
                        FileInfo sconn = new FileInfo("conn.txt");
                        if (!sconn.Exists)
                        {
                            MessageBox.Show("Не найден файл со строкой подключения. Дальнейшая работа невозможна");
                            return;
                        }
                        sreaderconn = sconn.OpenText();
                        while (!sreaderconn.EndOfStream)
                        {
                            string cn=sreaderconn.ReadLine();
                            if (cn.Contains(comboBox1.Text))
                            {
                                cn=cn.Remove(0, comboBox1.Text.Length);
                                Polohov.Form1.conn = new SqlConnection(cn);
                                //MessageBox.Show(cn);
                            }
                        }
                        
                        try
                        {
                            Polohov.Form1.conn.Open();
                            Polohov.Form1.fwd = true;
                            Polohov.Form1.conn.Close();
                            Polohov.Form1.user = comboBox1.Text;
                            this.Close();
                            return;
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show("Неверные параметры в строке подключения. Программа не сможет работать.");
                            return;
                        }
                        
                    }
                }
                
            }
            MessageBox.Show("Неправильный пароль!");
        }
    }
}
