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
    public partial class Othodi : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        public decimal vesmusor;
        public string musor;
        decimal ves;
        public Othodi(SqlConnection conn,decimal ves)
        {
            InitializeComponent();
            vesmusor = 0;
            musor = "";
            this.conn = conn;
            this.ves = ves;            
            command.Connection = conn;
            conn.Open();
            command.CommandText = "select name from tipothoda";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                if ((string)r[0] != "")
                    comboBox1.Items.Add((string)r[0]);
            }
            conn.Close();
            textBox1.Text = ves.ToString();
            label1.Text = "Вес отходов составляет "+ves+"кг\nВыберите в списке тип отходов и укажите их вес";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && textBox1.Text != "")
            {
                musor = comboBox1.Text;
                vesmusor = Convert.ToDecimal(textBox1.Text.Replace('.',','));
                this.Close();
            }
            else
            {
                MessageBox.Show("Заполните все поля");
                return;
            }
        }
    }
}