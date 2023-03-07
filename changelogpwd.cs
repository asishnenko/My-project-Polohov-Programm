using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Polohov
{
    public partial class changelogpwd : Form
    {
        public changelogpwd()
        {
            InitializeComponent();
            textBox1.Text = Polohov.Form1.user;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == ""||textBox2.Text==textBox1.Text)
            {
                MessageBox.Show("Заполните правильно все поля!");
                return;
            }
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
                string ttt = sreader.ReadToEnd().Replace(textBox1.Text, textBox2.Text);
                sreader.Close();
                swriter = usr.CreateText();
                swriter.Write(ttt);
                swriter.Close();
                Polohov.Form1.user = textBox2.Text;
                FileInfo con = new FileInfo("conn.txt");
                sreader = con.OpenText();
                ttt = sreader.ReadToEnd().Replace(textBox1.Text, textBox2.Text);
                sreader.Close();
                swriter = con.CreateText();
                swriter.Write(ttt);
                swriter.Close();
                MessageBox.Show("Успешно изменено!");
                this.Close();
            }
        }
    }
}
