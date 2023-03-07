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
    public partial class SetPartiya : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;        
        public SetPartiya(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;
            InitializeComponent();

            conn.Open();
            command.CommandText = "select name, show from partiya where name!='Не определено'";
            r = command.ExecuteReader();
            if(r.HasRows)
                while (r.Read())
                {
                    checkedListBox1.Items.Add((string)r[0], (bool)r[1]);
                }
            conn.Close();
            label1.Text = "Уберите галочки, если не хотите \r\nучитывать партии при формировании\r\nотчетов";
            checkedListBox1.CheckOnClick = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ArrayList arl = new ArrayList();
            conn.Open();

            for (int i=0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                arl.Add(checkedListBox1.CheckedItems[i].ToString());
            }
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (!arl.Contains(checkedListBox1.Items[i].ToString()))
                {
                    command.CommandText = "update partiya set show=0 where name='" + checkedListBox1.Items[i].ToString()+"'";
                    //MessageBox.Show("0");
                }
                else command.CommandText = "update partiya set show=1 where name='" + checkedListBox1.Items[i].ToString()+"'";
                command.ExecuteNonQuery();
            }
            conn.Close();
            MessageBox.Show("Изменения сохранены!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemCheckState(i) == CheckState.Checked)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemCheckState(i) == CheckState.Unchecked)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }
        }

    }
}