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
    public partial class changepwd : Form
    {
        public changepwd()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }
            if (textBox2.Text != textBox3.Text)
            {
                MessageBox.Show("Не совпадают пароли!");
                return;
            }
            StreamReader sreader;
            StreamWriter swriter;
            FileInfo usr = new FileInfo("users");
            if (usr.Exists)
            {
                sreader = usr.OpenText();
                string ttt = sreader.ReadToEnd().Replace(textBox1.Text, textBox2.Text);
                sreader.Close();
                swriter = usr.CreateText();
                swriter.Write(ttt);
                swriter.Close();
                MessageBox.Show("Успешно изменено!");
                this.Close();
            }

        }
    }
}
