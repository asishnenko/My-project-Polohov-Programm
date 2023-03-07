using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Polohov
{
    public partial class Zarplata : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        ListViewItem lvi;
        public Zarplata(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;

            InitializeComponent();
            listView1.Columns.Add("Дата", 120);
            listView1.Columns.Add("Смена", 60);
            listView1.Columns.Add("Рабочий", 100);
            listView1.Columns.Add("Станок", 100);
            listView1.Columns.Add("Партия", 100);
            listView1.Columns.Add("Продукт", 120);
            listView1.Columns.Add("Вес(кг)", 50);
            listView1.Columns.Add("Подсобники", 100);
            listView1.Columns.Add("Зарплата(грн)", 70);
            listView1.Columns.Add("idsob", 0);           
            

            conn.Open();
            command.CommandText = "select surname from rabotnik where working=1 and name!=''";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox1.Items.Add(r[0].ToString());
                }
            }
            conn.Close();
        }
        int startind = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            conn.Close();
            conn.Open();
            command.CommandText = "select idsklad,ves,data,idproizv,id from sobitie where iddvig=2 and idbalans=2 and price is null";
            r = command.ExecuteReader();            
            if (r.HasRows)
            {
                MessageBox.Show("Не у всех рабочих указана з/п! \r\nВыводимая информация будет не точной!");
                //Form1.label4.Text = "Не забудьте выдать зарплату!";
            }
            conn.Close();

            startind=listView1.Items.Count;

            //listView1.Items.Clear();
            if (comboBox1.Text != "" || dateTimePicker1.Text != "")
            {
                conn.Open();
                if (!checkBox1.Checked) command.CommandText = "select sobitie.data,sobitie.smena,rabotnik.surname,stanok.name,sobitie.id,sobitie.price from sobitie, rabotnik,proizv,vessklad,stanok where stanok.id=vessklad.idstanok and sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id and vessklad.idrabotnik=rabotnik.id  and sobitie.iddvigfrom=2 and sobitie.iddvig=2 and sobitie.idbalans=2 and rabotnik.surname='" + comboBox1.Text + "' and sobitie.data='" + dateTimePicker1.Text + "' and sobitie.price is not null";
                else
                {
                    string data4 = "";
                    ArrayList data22 = new ArrayList();
                    DateTime dt1 = Convert.ToDateTime(dateTimePicker1.Text);
                    DateTime dt2 = Convert.ToDateTime(dateTimePicker2.Text);
                    TimeSpan dt4 = new TimeSpan(1, 0, 0, 0);
                    while (dt1 <= dt2)
                    {
                        data22.Add(dt1.ToLongDateString());
                        dt1 = dt1 + dt4;
                    }

                    //цикл для периода       
                    if (data22.Count == 1) data4 = " and sobitie.data='" + (string)data22[0] + "'";
                    else
                    {
                        data4 = " and sobitie.data in (";
                        for (int i = 0; i < data22.Count - 1; i++)
                        {
                            data4 += "'" + (string)data22[i] + "',";
                        }
                        data4 += "'" + (string)data22[data22.Count - 1] + "') ";
                    }
                    command.CommandText = "select sobitie.data,sobitie.smena,rabotnik.surname,stanok.name,sobitie.id, sobitie.price from sobitie, rabotnik,proizv,vessklad,stanok where stanok.id=vessklad.idstanok and sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id and vessklad.idrabotnik=rabotnik.id  and sobitie.iddvigfrom=2 and sobitie.iddvig=2 and sobitie.idbalans=2 and rabotnik.surname='" + comboBox1.Text + "'" + data4 + "and sobitie.price is not null order by sobitie.data";
                }
                ArrayList idlist = new ArrayList();
                string data = "";
                string stanok = "";
                //string data = "";
                string smena = "";                
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        if (data == r[0].ToString() && stanok == r[3].ToString())
                        {
                            lvi = new ListViewItem(new string[] { "", "", "", "", "", "", "", "", "", r[4].ToString() });
                            listView1.Items.Add(lvi);
                        
                            continue;
                        }
                        smena = "Ночь";
                        if ((bool)r[1]) smena = "День";
                        data = r[0].ToString();
                        lvi = new ListViewItem(new string[] { data, smena, r[2].ToString(), r[3].ToString(), "", "", "", "", "", r[4].ToString()});
                        listView1.Items.Add(lvi);
                        
                        stanok = r[3].ToString();
                    }
                    lvi = new ListViewItem(new string[] { "Итого:", "", "", "", "", "", "", "", "","" });
                    lvi.BackColor = Color.Honeydew;
                    listView1.Items.Add(lvi);
                    
                    lvi = new ListViewItem(new string[] { "", "", "", "", "", "", "", "", "","" });
                    lvi.BackColor = Color.FromArgb(240,240,240);
                    listView1.Items.Add(lvi);
                    
                }
                conn.Close();
                

                //цикл взято
                //if (listView1.Items.Count == 0) return;
                decimal sum = 0;
                decimal sumzp = 0;
               
                for (int i = startind; i < listView1.Items.Count; i++)
                {
                    if (listView1.Items[i].SubItems[9].Text != "" && listView1.Items[i].SubItems[0].Text != "")
                    {
                        conn.Close();
                        conn.Open();
                        command.CommandText = "select zarplata.zp from zarplata where zarplata.data='" + listView1.Items[i].SubItems[0].Text + "' and zarplata.idrabotnik=(select id from rabotnik where surname='" + listView1.Items[i].SubItems[2].Text + "') and zarplata.smena='" + listView1.Items[i].SubItems[1].Text + "'";
                        decimal zzp = 0;
                        zzp = Convert.ToDecimal(command.ExecuteScalar().ToString());
                        listView1.Items[i].SubItems[8].Text = zzp.ToString();
                        sumzp += zzp;
                        conn.Close();

                        conn.Open();
                        command.CommandText = "select rabotnik.surname from rabotnik,podsobniki where podsobniki.idpodsobnik=rabotnik.id and podsobniki.idsobitie=" + listView1.Items[i].SubItems[9].Text;
                        r = command.ExecuteReader();
                        if (r.HasRows)
                        {
                            while (r.Read())
                            {
                                listView1.Items[i].SubItems[7].Text += r[0] + ",";
                            }
                        }
                        conn.Close();
                    }

                    if (listView1.Items[i].SubItems[9].Text != "")
                    {
                        conn.Open();
                        command.CommandText = "select partiya.name,prodykt.name,sobitie.ves from partiya,prodykt,sobitie,vessklad where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.id=" + listView1.Items[i].SubItems[9].Text;
                        r = command.ExecuteReader();
                        if (r.HasRows)
                        {
                            r.Read();
                            listView1.Items[i].SubItems[4].Text = r[0].ToString();
                            listView1.Items[i].SubItems[5].Text = r[1].ToString();
                            listView1.Items[i].SubItems[6].Text = r[2].ToString();
                            sum += Convert.ToDecimal(r[2]);
                        }
                        conn.Close();
                    }                    
                    if (listView1.Items[i].SubItems[0].Text == "Итого:")
                    {
                        listView1.Items[i].SubItems[8].Text = sumzp.ToString();
                        sumzp = 0;
                        listView1.Items[i].SubItems[6].Text = sum.ToString();
                        sum = 0;                                
                    }
                }
                //ищем как подсобников
                startind = listView1.Items.Count;
                conn.Close();
                conn.Open();
                if (!checkBox1.Checked) command.CommandText = "select sobitie.data,sobitie.smena,stanok.name,sobitie.id from sobitie,proizv,vessklad,stanok where stanok.id=vessklad.idstanok and sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id  and sobitie.id in (select podsobniki.idsobitie from podsobniki,rabotnik where podsobniki.idpodsobnik=rabotnik.id and rabotnik.surname='"+comboBox1.Text+"') and sobitie.data='"+dateTimePicker1.Text+"'  and sobitie.price is not null";
                else
                {
                    string data4 = "";
                    ArrayList data22 = new ArrayList();
                    DateTime dt1 = Convert.ToDateTime(dateTimePicker1.Text);
                    DateTime dt2 = Convert.ToDateTime(dateTimePicker2.Text);
                    TimeSpan dt4 = new TimeSpan(1, 0, 0, 0);
                    while (dt1 <= dt2)
                    {
                        data22.Add(dt1.ToLongDateString());
                        dt1 = dt1 + dt4;
                    }

                    //цикл для периода       
                    if (data22.Count == 1) data4 = " and sobitie.data='" + (string)data22[0] + "'";
                    else
                    {
                        data4 = " and sobitie.data in (";
                        for (int i = 0; i < data22.Count - 1; i++)
                        {
                            data4 += "'" + (string)data22[i] + "',";
                        }
                        data4 += "'" + (string)data22[data22.Count - 1] + "') ";
                    }
                
                    command.CommandText = "select sobitie.data,sobitie.smena,stanok.name,sobitie.id from sobitie,proizv,vessklad,stanok where stanok.id=vessklad.idstanok and sobitie.idproizv=proizv.id and proizv.idprodykt1=vessklad.id  and sobitie.id in (select podsobniki.idsobitie from podsobniki,rabotnik where podsobniki.idpodsobnik=rabotnik.id and rabotnik.surname='" + comboBox1.Text + "') "+data4+" and sobitie.price is not null order by sobitie.data";
                }
                    data = "";
                    stanok = "";
                    string rabotnik = "";
                    //string data = "";
                    smena = "";
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        rabotnik = comboBox1.Text;
                        lvi = new ListViewItem(new string[] { "Как подсобник:", "", "", "", "", "", "", "", "", "" });
                        listView1.Items.Add(lvi);
                        while (r.Read())
                        {
                            if (data == r[0].ToString() && stanok == r[2].ToString())
                            {
                                lvi = new ListViewItem(new string[] { "", "", "", "", "", "", "", "", "", r[3].ToString() });
                                listView1.Items.Add(lvi);

                                continue;
                            }
                            smena = "Ночь";
                            if ((bool)r[1]) smena = "День";
                            data = r[0].ToString();
                            lvi = new ListViewItem(new string[] { data, smena, rabotnik, r[2].ToString(), "", "", "", "", "", r[3].ToString() });
                            listView1.Items.Add(lvi);

                            stanok = r[2].ToString();
                        }
                        lvi = new ListViewItem(new string[] { "Итого:", "", "", "", "", "", "", "", "", "" });
                        lvi.BackColor = Color.Honeydew;
                        listView1.Items.Add(lvi);

                        lvi = new ListViewItem(new string[] { "", "", "", "", "", "", "", "", "", "" });
                        lvi.BackColor = Color.Gray;
                        listView1.Items.Add(lvi);

                    }
                    conn.Close();

                    //цикл взято
                    //if (listView1.Items.Count == 0) return;
                    sum = 0;
                    sumzp = 0;

                    for (int i = startind; i < listView1.Items.Count; i++)
                    {
                        if (listView1.Items[i].SubItems[9].Text != "" && listView1.Items[i].SubItems[0].Text != "")
                        {
                            conn.Close();
                            conn.Open();
                            command.CommandText = "select zarplata.zp from zarplata where zarplata.data='" + listView1.Items[i].SubItems[0].Text + "' and zarplata.idrabotnik=(select id from rabotnik where surname='" + listView1.Items[i].SubItems[2].Text + "') and zarplata.smena='" + listView1.Items[i].SubItems[1].Text + "'";
                            decimal zzp = 0;
                            zzp = Convert.ToDecimal(command.ExecuteScalar().ToString());
                            listView1.Items[i].SubItems[8].Text = zzp.ToString();
                            sumzp += zzp;
                            conn.Close();

                            //conn.Open();
                            //command.CommandText = "select rabotnik.surname from rabotnik,podsobniki where podsobniki.idpodsobnik=rabotnik.id and podsobniki.idsobitie=" + listView1.Items[i].SubItems[9].Text;
                            //r = command.ExecuteReader();
                            //if (r.HasRows)
                            //{
                            //    while (r.Read())
                            //    {
                            //        listView1.Items[i].SubItems[7].Text += r[0] + ",";
                            //    }
                            //}
                            //conn.Close();
                        }

                        if (listView1.Items[i].SubItems[9].Text != "")
                        {
                            conn.Open();
                            command.CommandText = "select partiya.name,prodykt.name,sobitie.ves from partiya,prodykt,sobitie,vessklad where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.id=" + listView1.Items[i].SubItems[9].Text;
                            r = command.ExecuteReader();
                            if (r.HasRows)
                            {
                                r.Read();
                                listView1.Items[i].SubItems[4].Text = r[0].ToString();
                                listView1.Items[i].SubItems[5].Text = r[1].ToString();
                                listView1.Items[i].SubItems[6].Text = r[2].ToString();
                                sum += Convert.ToDecimal(r[2]);
                            }
                            conn.Close();
                        }
                        if (listView1.Items[i].SubItems[0].Text == "Итого:")
                        {
                            listView1.Items[i].SubItems[8].Text = sumzp.ToString();
                            sumzp = 0;
                            listView1.Items[i].SubItems[6].Text = sum.ToString();
                            sum = 0;
                        }
                    }                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Зарплата";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
            oWord.Visible = true;

            int r = 0;
            int c = 0;
            c = listView1.Columns.Count-1;
            r = listView1.Items.Count+1;
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, r, c, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;

            for (int i = 1; i <= c; i++)
            {
                oTable.Cell(1, i).Range.Text = listView1.Columns[i - 1].Text;
            }

            for (int i = 2; i <= r; i++)
            {
                for (int j = 1; j <= c; j++)
                {
                    if (listView1.Items[i - 2].SubItems[j - 1].Text == "Итого:") oTable.Rows[i].Range.Font.Shadow = 5;
                    oTable.Cell(i, j).Range.Text = listView1.Items[i - 2].SubItems[j - 1].Text;
                }
            }
            oTable.Borders.Enable = 1;
            oTable.Borders[WdBorderType.wdBorderTop].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Borders[WdBorderType.wdBorderLeft].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Borders[WdBorderType.wdBorderRight].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) dateTimePicker2.Enabled = true;
            else dateTimePicker2.Enabled = false;
        }

    }
}