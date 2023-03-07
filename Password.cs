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
    public partial class Password : Form
    {
        public Password()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
            {
                StreamReader sreader;
                StreamWriter swriter;
                //parol = "1111";
                string pass = "1111";
                FileInfo pwdd = new FileInfo("oll");
                if (!pwdd.Exists)
                {
                    swriter = pwdd.CreateText();
                    swriter.WriteLine("1111");
                    swriter.Close();
                    //pass = "1111";
                }
                else
                {
                    sreader = pwdd.OpenText();
                    pass = sreader.ReadLine();
                    sreader.Close();
                }
                if (textBox3.Text != pass) { MessageBox.Show("Неверно указан старый пароль!"); return; }
                if (textBox1.Text != textBox2.Text) { MessageBox.Show("Неверно указан повтор нового пароля!"); return; }
                swriter = pwdd.CreateText();
                swriter.WriteLine(textBox2.Text);
                swriter.Close();
                sreader = pwdd.OpenText();
                if (sreader.ReadLine() == textBox2.Text) MessageBox.Show("Пароль успешно изменен!");
                else MessageBox.Show("Ошибка изменения!");
            }
            else
            {
                MessageBox.Show("Заполните все поля!");
            }
        }
  
    }
}