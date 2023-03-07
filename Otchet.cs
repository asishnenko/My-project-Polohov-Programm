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
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Polohov
{
    public partial class Otchet : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        ListViewItem lvi;
        public Otchet(SqlConnection conn, string label)
        {
            this.conn = conn;
            command.Connection = conn;
            InitializeComponent();

            label3.Text = label;

            
            listView1.Columns.Add("Партия", 100);
            listView1.Columns.Add("Продукт", 100);
            listView1.Columns.Add("Вес по программе(кг)", 80);
            listView1.Columns.Add("Факт. вес(кг)", 80);
            listView1.Columns.Add("Расхождение(кг)", 80);
            listView1.Columns.Add("Цена(грн/кг)", 60);
            listView1.Columns.Add("Сумма(грн)", 60);
            listView1.Columns.Add("Примечание", 120);

            if (label3.Text == "Весь склад")
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,vessklad.tcena from partiya,prodykt,vessklad where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows)
                    while (r.Read())
                    {
                        //if (r[3].ToString() != "")
                        {
 
                        }
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[2].ToString(), "", r[3].ToString(), "", "" });
                        listView1.Items.Add(lvi);
                    }
                //lvi.BackColor = Color.Honeydew;
                conn.Close();
            }
            if (label3.Text == "Склад производства")
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,vessklad.tcena from partiya,prodykt,vessklad,state where state.name ='" + label3.Text + "' and state.id=vessklad.idstate and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows)
                    while (r.Read())
                    {
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[2].ToString(), "", r[3].ToString(), "", "" });
                        listView1.Items.Add(lvi);
                    }
                //lvi.BackColor = Color.Honeydew;
                conn.Close();
            }
            if (label3.Text == "Готовая продукция")
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,vessklad.tcena from partiya,prodykt,vessklad,state where state.name ='" + label3.Text + "' and state.id=vessklad.idstate and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows)
                    while (r.Read())
                    {
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[2].ToString(), "", r[3].ToString(), "", "" });
                        listView1.Items.Add(lvi);
                    }
                //lvi.BackColor = Color.Honeydew;
                conn.Close();
            }

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                listView1.Items[i].SubItems[4].Text = "0";
                if (listView1.Items[i].SubItems[5].Text != "")
                {
                    decimal ves = 0;
                    decimal tcena = 0;
                    string vess = listView1.Items[i].SubItems[3].Text.Replace('.', ',');
                    string tcenaa = listView1.Items[i].SubItems[5].Text.Replace('.', ',');
                    if (listView1.Items[i].SubItems[3].Text.Replace('.', ',') == "") vess = "0";
                    ves = Convert.ToDecimal(vess) - Convert.ToDecimal(listView1.Items[i].SubItems[2].Text.Replace('.', ','));
                    listView1.Items[i].SubItems[4].Text = ves.ToString();
                    if (listView1.Items[i].SubItems[5].Text.Replace('.', ',') == "") tcenaa = "0";
                    tcena = Convert.ToDecimal(vess) * Convert.ToDecimal(tcenaa);
                    listView1.Items[i].SubItems[6].Text = tcena.ToString();
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                listView1.SelectedItems[0].SubItems[3].Text = textBox1.Text;
                listView1.SelectedItems[0].SubItems[5].Text = textBox2.Text;
                decimal ves = 0;
                decimal tcena = 0;
                string vess = listView1.SelectedItems[0].SubItems[3].Text.Replace('.', ',');
                string tcenaa = listView1.SelectedItems[0].SubItems[5].Text.Replace('.', ',');
                if (listView1.SelectedItems[0].SubItems[3].Text.Replace('.', ',') == "") vess = "0";
                ves = Convert.ToDecimal(vess) - Convert.ToDecimal(listView1.SelectedItems[0].SubItems[2].Text.Replace('.', ','));
                listView1.SelectedItems[0].SubItems[4].Text = ves.ToString();
                if (listView1.SelectedItems[0].SubItems[5].Text.Replace('.', ',') == "") tcenaa = "0";
                tcena = Convert.ToDecimal(vess) * Convert.ToDecimal(tcenaa);
                listView1.SelectedItems[0].SubItems[6].Text = tcena.ToString();
                conn.Open();
                command.CommandText = "select id from partiya where name='"+listView1.SelectedItems[0].SubItems[0].Text+"'";
                int idprt = (int)command.ExecuteScalar();
                command.CommandText = "select id from prodykt where name='" + listView1.SelectedItems[0].SubItems[1].Text + "'";
                int idpod = (int)command.ExecuteScalar();
                command.CommandText = "update vessklad set tcena=" + tcenaa.ToString().Replace(',', '.') + " where idpartiya=" + idprt + " and idprodykt=" + idpod + " and ostatok=" + listView1.SelectedItems[0].SubItems[2].Text.Replace(',', '.');
                command.ExecuteNonQuery();
                conn.Close();
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                textBox1.Text = listView1.SelectedItems[0].SubItems[3].Text;
                textBox2.Text = listView1.SelectedItems[0].SubItems[5].Text;
            }
        }
        private void SetItogo()
        {
            decimal ves_prog=0;
            decimal ves_fakt=0; decimal raznitca=0; decimal summa=0;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].SubItems[2].Text!="")
                ves_prog += Convert.ToDecimal(listView1.Items[i].SubItems[2].Text);
            }
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].SubItems[3].Text!="")
                ves_fakt += Convert.ToDecimal(listView1.Items[i].SubItems[3].Text);
            }
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].SubItems[4].Text!="")
                raznitca += Convert.ToDecimal(listView1.Items[i].SubItems[4].Text);
            }
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].SubItems[6].Text!="")
                summa += Convert.ToDecimal(listView1.Items[i].SubItems[6].Text);
            }
            listView1.Items.Add(new ListViewItem(new string[] { "Итого", "", ves_prog.ToString(), ves_fakt.ToString(), raznitca.ToString(), "", summa.ToString(), "" }));
        }
        private void button2_Click(object sender, EventArgs e)
        {

            SetItogo();
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
            oPara1.Range.Text = "Утвердил\t\t\t";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            oPara1.Format.SpaceAfter = 12;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
            Word.Paragraph oPara11;
            oPara11 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara11.Range.Text = "нач. пр-ва _____________\t";
            oPara11.Range.Font.Bold = 1;
            oPara1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            oPara11.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara11.Range.InsertParagraphAfter();


            Word.Paragraph oPara2;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara2.Range.Text = label3.Text;
            oPara2.Range.Font.Bold = 1;
            oPara2.Format.SpaceAfter = 24;//24 pt spacing after paragraph.
            oPara2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oPara2.Range.InsertParagraphAfter();

            int r = 0;
            int c = 0;
            c = listView1.Columns.Count;
            r = listView1.Items.Count + 1;
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, r, c, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 12;
            //oTable.Range.InsertParagraphAfter();

            for (int i = 1; i <= c; i++)
            {
                oTable.Cell(1, i).Range.Text = listView1.Columns[i - 1].Text;
            }

            for (int i = 2; i <= r; i++)
            {
                for (int j = 1; j <= c; j++)
                {
                    //if (listView1.Items[i - 2].SubItems[j - 1].Text == "Итого:") oTable.Rows[i].Range.Font.Shadow = 5;
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

            Word.Paragraph oPara31;
            oPara31 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara31.Range.Text = "";
            oPara31.Range.Font.Bold = 1;
            oPara31.Format.SpaceAfter = 12;//24 pt spacing after paragraph.
            oPara31.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oPara31.Range.InsertParagraphAfter();

            Word.Paragraph oPara3;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara3.Range.Text = "Составил технолог _______________";
            oPara3.Range.Font.Bold = 1;
            oPara3.Format.SpaceAfter = 12;//24 pt spacing after paragraph.
            oPara3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.InsertParagraphAfter();

            //oPara3.Range.InsertParagraphBefore();
            Word.Paragraph oPara4;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara4.Range.Text = "Сдал кладовщик _______________";
            oPara4.Range.Font.Bold = 1;
            oPara4.Format.SpaceAfter = 12;//24 pt spacing after paragraph.
            oPara4.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oPara4.Range.InsertParagraphAfter();
            Word.Paragraph oPara5;
            oPara5 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara5.Range.Text = "Принял кладовщик _______________";
            oPara5.Range.Font.Bold = 1;
            oPara5.Format.SpaceAfter = 12;//24 pt spacing after paragraph.
            oPara5.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oPara5.Range.InsertParagraphAfter();

            oWord.Visible = true;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {

            SetItogo();
            object m_objOpt = System.Reflection.Missing.Value;
            Excel.Application m_objExcel = new Excel.Application();
            Excel.Workbooks m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            Excel._Workbook m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            ///////
            int r = 0;
            int c = 0;
            c = listView1.Columns.Count;
            r = listView1.Items.Count + 1;
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
                    rrt = listView1.Items[r2 - 1].SubItems[h].Text;
                    objData[r2 - 1, h] = rrt;
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
