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
using Word=Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Polohov
{
    public partial class checkProdaja : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        ListViewItem lvi;
        string wtext;
        string wordtext;
        public checkProdaja(SqlConnection conn,string label)
        {
            this.conn = conn;
            command.Connection = conn;
            wtext = label;
            InitializeComponent();
            Text = label;
            if (Text == "Показать проданное")
            {
                wordtext = "Продано";
                listView1.Columns.Add("Дата", 120);
                listView1.Columns.Add("Клиент", 100);
                listView1.Columns.Add("Партия", 100);
                listView1.Columns.Add("Продукт", 120);
                listView1.Columns.Add("Вес(кг)", 50);

                comboBox1.Items.Add("Все");
                comboBox2.Items.Add("Все");
                comboBox3.Items.Add("Все");

                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;

                conn.Open();
                command.CommandText = "select name from partiya where show=1 and name!='Не определено'";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox1.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();

                conn.Open();
                command.CommandText = "select distinct prodykt.name from prodykt,vessklad,sobitie where sobitie.iddvig=3 and vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox2.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();


                conn.Open();
                command.CommandText = "select distinct kAgent.name from kAgent,sobitie where sobitie.iddvig=3 and sobitie.idkAgent=kAgent.id";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox3.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();
            }
            if (Text == "Показать объединенное")
            {
                wordtext = "Объединено";
                comboBox3.Visible = false;
                label4.Visible = false;
                listView1.Columns.Add("Дата", 120);
                //listView1.Columns.Add("Клиент", 100);
                listView1.Columns.Add("Партия", 100);
                listView1.Columns.Add("Продукт", 120);
                listView1.Columns.Add("Вес(кг)", 50);
                listView1.Columns.Add("Склад", 100);

                comboBox1.Items.Add("Все");
                comboBox2.Items.Add("Все");
                comboBox3.Items.Add("Все");

                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;

                conn.Open();
                command.CommandText = "select name from partiya where show=1 and name!='Не определено'";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox1.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();

                conn.Open();
                command.CommandText = "select prodykt.name from prodykt";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox2.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();
            }

            if (Text == "История")
            {
                wordtext = "История";
                comboBox1.Visible = false;
                label1.Visible = false;
                comboBox2.Visible = false;
                label3.Visible = false;
                comboBox3.Visible = false;
                label4.Visible = false;
                checkBox1.Visible = false;
                dateTimePicker2.Visible = false;

                listView1.Columns.Add("Время записи", 160);
                listView1.Columns.Add("Дата", 160);
                //listView1.Columns.Add("Клиент", 100);
                listView1.Columns.Add("Партия", 100);
                listView1.Columns.Add("Продукт", 120);
                listView1.Columns.Add("Вес(кг)", 50);
                listView1.Columns.Add("Склад", 100);
                listView1.Columns.Add("Тип события", 120);
            }
            if (Text == "Показать приход товара")
            {
                wordtext = "Приход";
                listView1.Columns.Add("Дата", 120);
                listView1.Columns.Add("Клиент", 100);
                listView1.Columns.Add("Партия", 100);
                listView1.Columns.Add("Продукт", 120);
                listView1.Columns.Add("Вес(кг)", 50);
                listView1.Columns.Add("Цена(грн/кг)", 50);
                listView1.Columns.Add("Сумма(грн)", 50);

                comboBox1.Items.Add("Все");
                comboBox2.Items.Add("Все");
                comboBox3.Items.Add("Все");

                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;

                conn.Open();
                command.CommandText = "select name from partiya where show=1 and name!='Не определено'";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox1.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();

                conn.Open();
                command.CommandText = "select distinct prodykt.name from prodykt,vessklad,sobitie where sobitie.iddvig=1 and vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox2.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();


                conn.Open();
                command.CommandText = "select distinct kAgent.name from kAgent,sobitie where sobitie.iddvig=1 and sobitie.idkAgent=kAgent.id";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox3.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();
            }

        }
        int startind = 0;
        private void button1_Click(object sender, EventArgs e)
        {

            if (Text == "Показать приход товара")
            {
                if (dateTimePicker1.Value >= dateTimePicker2.Value) { MessageBox.Show("Несоответствие даты!"); return; }
                string partiya = "";
                if (comboBox1.Text != "Все") { partiya = " and partiya.name='" + comboBox1.Text + "'"; }
                string prodykt = "";
                if (comboBox2.Text != "Все") { prodykt = " and prodykt.name='" + comboBox2.Text + "'"; }
                string kAgetnt = "";
                if (comboBox3.Text != "Все") { kAgetnt = " and kAgent.name='" + comboBox3.Text + "'"; }

                startind = listView1.Items.Count;

                if (!checkBox1.Checked) command.CommandText = "select sobitie.data,kAgent.name,partiya.name,prodykt.name,sobitie.ves, vessklad.tcena from sobitie,kAgent,partiya,prodykt,vessklad where kAgent.id=sobitie.idkAgent and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id " + partiya + prodykt + kAgetnt + " and sobitie.data='" + dateTimePicker1.Text + "' and sobitie.idbalans=1 and sobitie.iddvig=1";
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
                    command.CommandText = "select sobitie.data,kAgent.name,partiya.name,prodykt.name,sobitie.ves, vessklad.tcena from sobitie,kAgent,partiya,prodykt,vessklad where kAgent.id=sobitie.idkAgent and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.idbalans=1 " + partiya + prodykt + kAgetnt + data4;
                }
                decimal sumves = 0;
                decimal sumsum = 0;
                conn.Close();
                conn.Open();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        decimal ddd = 0;
                        if (r[5].ToString() != "")
                        {
                            decimal d1 = Convert.ToDecimal(r[4].ToString());
                            decimal d2 = Convert.ToDecimal(r[5].ToString());
                            ddd = d1 * d2;
                        }
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString(), r[5].ToString(),ddd.ToString() });
                        listView1.Items.Add(lvi);
                        sumves += Convert.ToDecimal(r[4].ToString());
                        sumsum += Convert.ToDecimal(lvi.SubItems[6].Text);
                    }
                    lvi = new ListViewItem(new string[] { "Итого:", "", "", "", sumves.ToString(),"",sumsum.ToString() });
                    lvi.BackColor = Color.Honeydew;
                    listView1.Items.Add(lvi);

                    lvi = new ListViewItem(new string[] { "", "", "", "", "" });
                    lvi.BackColor = Color.FromArgb(240, 240, 240);
                    listView1.Items.Add(lvi);
                }
                conn.Close();
            }

            if (Text == "Показать проданное")
            {
                if (dateTimePicker1.Value >= dateTimePicker2.Value) { MessageBox.Show("Несоответствие даты!"); return; }
                string partiya = "";
                if (comboBox1.Text != "Все") { partiya = " and partiya.name='" + comboBox1.Text + "'"; }
                string prodykt = "";
                if (comboBox2.Text != "Все") { prodykt = " and prodykt.name='" + comboBox2.Text + "'"; }
                string kAgetnt = "";
                if (comboBox3.Text != "Все") { kAgetnt = " and kAgent.name='" + comboBox3.Text + "'"; }

                startind = listView1.Items.Count;

                if (!checkBox1.Checked) command.CommandText = "select sobitie.data,kAgent.name,partiya.name,prodykt.name,sobitie.ves from sobitie,kAgent,partiya,prodykt,vessklad where kAgent.id=sobitie.idkAgent and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id " + partiya + prodykt + kAgetnt + " and sobitie.data='" + dateTimePicker1.Text + "' and sobitie.idbalans=1 and sobitie.iddvig=1";
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
                    command.CommandText = "select sobitie.data,kAgent.name,partiya.name,prodykt.name,sobitie.ves from sobitie,kAgent,partiya,prodykt,vessklad where kAgent.id=sobitie.idkAgent and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.idbalans=2 " + partiya + prodykt + kAgetnt + data4;
                }
                decimal sumves = 0;
                conn.Close();
                conn.Open();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString() });
                        listView1.Items.Add(lvi);
                        sumves += Convert.ToDecimal(r[4].ToString());
                    }
                    lvi = new ListViewItem(new string[] { "Итого:", "", "", "", sumves.ToString() });
                    lvi.BackColor = Color.Honeydew;
                    listView1.Items.Add(lvi);

                    lvi = new ListViewItem(new string[] { "", "", "", "", "" });
                    lvi.BackColor = Color.FromArgb(240, 240, 240);
                    listView1.Items.Add(lvi);
                }
                conn.Close();
            }

            if (Text == "Показать объединенное")
            {
                if (dateTimePicker1.Value >= dateTimePicker2.Value) { MessageBox.Show("Несоответствие даты!"); return; }
                string partiya = "";
                if (comboBox1.Text != "Все") { partiya = " and partiya.name='" + comboBox1.Text + "'"; }
                string prodykt = "";
                if (comboBox2.Text != "Все") { prodykt = " and prodykt.name='" + comboBox2.Text + "'"; }
                //string kAgetnt = "";
                //if (comboBox3.Text != "Все") { kAgetnt = " and kAgent.name='" + comboBox3.Text + "'"; }

                startind = listView1.Items.Count;

                //if (!checkBox1.Checked) command.CommandText = "select sobitie.data,kAgent.name,partiya.name,prodykt.name,sobitie.ves from sobitie,kAgent,partiya,prodykt,vessklad where kAgent.id=sobitie.idkAgent and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id " + partiya + prodykt + " and sobitie.data='" + dateTimePicker1.Text + "' and sobitie.idbalans=2";
                if (!checkBox1.Checked) command.CommandText = "select sobitie.idproizv from sobitie,partiya,prodykt,vessklad where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id " + partiya + prodykt + " and sobitie.data='" + dateTimePicker1.Text + "' and sobitie.idbalans=1 and sobitie.iddvigfrom=5 and sobitie.iddvig=1";
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
                    command.CommandText = "select sobitie.idproizv from sobitie,partiya,prodykt,vessklad where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idsklad=vessklad.id and sobitie.idbalans=1 and sobitie.iddvig=1 and sobitie.iddvigfrom=5 " + partiya + prodykt + data4;
                }

                ArrayList sobAr = new ArrayList();
                conn.Close();
                conn.Open();
                r = command.ExecuteReader();
                if(r.HasRows)
                    while (r.Read())
                    {
                        sobAr.Add(r[0]);
                    }
                conn.Close();
                for (int i = 0; i < sobAr.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "select sobitie.data,partiya.name,prodykt.name,sobitie.ves,state.name from sobitie,partiya,prodykt,vessklad,state where state.id=vessklad.idstate and sobitie.idsklad=vessklad.id and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idbalans=1 and sobitie.idproizv="+sobAr[i].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        r.Read();
                    lvi = new ListViewItem(new string[]{r[0].ToString(),r[1].ToString(),r[2].ToString(),r[3].ToString(),r[4].ToString()});
                    listView1.Items.Add(lvi);
                    lvi.BackColor = Color.Honeydew;
                    conn.Close();

                    conn.Open();
                    command.CommandText = "select sobitie.data,partiya.name,prodykt.name,sobitie.ves,state.name from sobitie,partiya,prodykt,vessklad,state where state.id=vessklad.idstate and sobitie.idsklad=vessklad.id and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sobitie.idbalans=2 and sobitie.idproizv=" + sobAr[i].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            lvi = new ListViewItem(new string[] { "", r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString() });
                            listView1.Items.Add(lvi);
                            lvi.BackColor = Color.FromArgb(220, 220, 220);
                        }
                    conn.Close();

                    lvi = new ListViewItem(new string[] { "", "", "", "", "" });
                    lvi.BackColor = Color.FromArgb(240, 240, 240);
                    listView1.Items.Add(lvi);
                }
            }


            if (Text == "История")
            {
                //if (dateTimePicker1.Value >= dateTimePicker2.Value) { MessageBox.Show("Несоответствие даты!"); return; }
                startind = listView1.Items.Count;
                //MessageBox.Show(dateTimePicker1.Value.ToShortDateString());
                //return;
                //if (!checkBox1.Checked) 
                command.CommandText = "select sobitie.recordtime,sobitie.data,partiya.name,prodykt.name,sobitie.ves,state.name,sobitie.iddvigfrom,sobitie.iddvig, sobitie.idbalans, sobitie.idproizv from sobitie,vessklad,partiya,prodykt,state where sobitie.idsklad=vessklad.id and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstate=state.id  and sobitie.recordtime like '" + dateTimePicker1.Value.ToShortDateString() + "%' order by sobitie.recordtime asc";               
                conn.Close();
                conn.Open();
                string txt = "";
                string dataold = "";
                bool col = false;
                bool once = true;
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        txt = "";                   
                        string r1 = r[6].ToString();
                        string r2 = r[7].ToString();
                        string r3 = r[8].ToString();
                        if (r1 == "" && r2 == "1" && r3 == "1") txt = "Поступление товара";
                        if (r1 == "" && r2 == "9" && r3 == "1") txt = "Заполнение остатков";
                        if (r1 == "" && r2 == "3" && r3 == "2") txt = "Продажа";
                        if (r1 == "7" && r2 == "2" && r3 == "") txt = "Перемещение на склад ГП";
                        if (r1 == "2" && r2 == "7" && r3 == "") txt = "Перемещение на склад производства";
                        if (r1 == "5" && r2 == "5" && r3 == "2") txt = "Добавлен к сборному";
                        if (r1 == "5" && r2 == "1" && r3 == "1") txt = "Создание сборного";
                        if (r1 == "" && r2 == "8" && r3 == "2") txt = "Удаление";
                        if (r1 == "2" && r2 == "2" && r3 == "2") txt = "Взято в переработку";
                        if (r1 == "2" && r2 == "2" && r3 == "1") txt = "Получено при переработке";
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString(), r[5].ToString(), txt });
                        if (txt != "")
                        {
                            //col = false;
                            if (once) { dataold = r[0].ToString(); once = false; }
                            if (dataold != r[0].ToString())
                            {
                                dataold = r[0].ToString();
                                if (col)
                                {
                                    col = false;
                                    //continue;
                                    goto m1;
                                }
                                if (!col) col = true;
                            }
                        m1:   listView1.Items.Add(lvi);
                            if (col)
                            {
                                lvi.BackColor = Color.Honeydew;
                            }
                            else lvi.BackColor = Color.Gray;
                        }
                        
                        //sumves += Convert.ToDecimal(r[4].ToString());
                    }
                    //lvi = new ListViewItem(new string[] { "Итого:", "", "", "", sumves.ToString() });
                    //lvi.BackColor = Color.Honeydew;
                    //listView1.Items.Add(lvi);

                    //lvi = new ListViewItem(new string[] { "", "", "", "", "" });
                    //lvi.BackColor = Color.FromArgb(240, 240, 240);
                    //listView1.Items.Add(lvi);
                }
                conn.Close();
            }

        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) dateTimePicker2.Enabled = true;
            else dateTimePicker2.Enabled = false;
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
            
            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = wordtext;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

   
            int r = 0;
            int c = 0;
            c = listView1.Columns.Count;
            r = listView1.Items.Count;
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, r, c, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
  
            for (int i = 1; i <= c; i++)
            {
                oTable.Cell(1, i).Range.Text = listView1.Columns[i-1].Text;
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
            oWord.Visible = true;
                        
        }

        private void button5_Click(object sender, EventArgs e)
        {
            object m_objOpt = System.Reflection.Missing.Value;
            Excel.Application m_objExcel = new Excel.Application();
            Excel.Workbooks m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            Excel._Workbook m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            ///////
            int r = 0;
            int c = 0;
            c = listView1.Columns.Count;
            r = listView1.Items.Count;
            //if (label2.Text == "Весь склад" || label2.Text == "Склад производства" || label2.Text == "Готовая продукция") c = dataGridView1.ColumnCount - 2;

            // Add data to cells in the first worksheet in the new workbook.
            Excel.Sheets m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            Excel._Worksheet m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));

            object[] objHeaders = new object[c];
            for (int i = 1; i <= c; i++)
            {
                objHeaders[i - 1] = listView1.Columns[i - 1].Text;
            }

            Excel.Range m_objRange = m_objSheet.get_Range("A1");
            m_objRange = m_objRange.get_Resize(1, c);
            m_objRange.Value = objHeaders;
            Excel.Font m_objFont = m_objRange.Font;
            m_objFont.Bold = true;
            m_objRange.ColumnWidth = 20;

            // Create an array with 3 columns and 100 rows and add it to
            // the worksheet starting at cell A2.
            object[,] objData = new Object[r, c];
            string rrt = "";
            for (int r2 = 1; r2 < r; r2++)
            {
                for (int h = 0; h < c; h++)
                {
                    //objData[r2, h] = ;      
                    //if (listView1.Items[i - 2].SubItems[j - 1].Text == "Итого:") oTable.Rows[i].Range.Font.Shadow = 5;
                    //if (listView1.Items[r2 - 1].SubItems[0].Text != "Итого:")
                    {
                        rrt = listView1.Items[r2 - 1].SubItems[h].Text;
                        objData[r2 - 1, h] = rrt;
                    }
                }
            }
            m_objRange = m_objSheet.get_Range("A2", m_objOpt);
            m_objRange = m_objRange.get_Resize(r, c);
            m_objRange.Value = objData;
            m_objRange.NumberFormat = "@";
            m_objExcel.Visible = true;
        }

    }
}