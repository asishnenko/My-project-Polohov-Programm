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
    public partial class zadanie : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        Zapolnenie zap;
        Othodi ot;
        public int id;
        decimal summavzyato;
        decimal summapolycheno;
        //int idmusor;
        //int vesmusor;
        public zadanie(SqlConnection conn)
        {
            InitializeComponent();
            //summa = 0;
            this.conn = conn;
            command.Connection = conn;
            listView1.Columns.Add("Партия",50);
            listView1.Columns.Add("Наименование",100);
            listView1.Columns.Add("Количество(кг)",100);
            listView1.Columns.Add("", 0);
            listView1.Columns.Add("Взято(кг)",70);

            listView2.Columns.Add("Наименование", 100);
            listView2.Columns.Add("Кол.(кг)", 60);
            listView2.Columns.Add("Направить", 100);
            listView2.MultiSelect = true;
            listView2.FullRowSelect = true;

            comboBox2.Items.Add("Готовая продукция");
            comboBox2.Items.Add("Склад производства");

            comboBox1.Enabled = false;
            textBox2.Enabled = false;
            button4.Enabled = false;
            comboBox2.Enabled = false;
            button7.Enabled = false;

            conn.Open();
            command.CommandText = "select name from stanok";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                if ((string)r[0] != "")
                comboBox4.Items.Add((string)r[0]);
            }
            conn.Close();

            conn.Open();
            command.CommandText = "select surname from rabotnik";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                if ((string)r[0] != "")
                comboBox5.Items.Add((string)r[0]);
            }
            conn.Close();

            comboBox1.Items.Add("");
            conn.Open();
            command.CommandText = "select name from prodykt";
            r = command.ExecuteReader();
            while (r.Read() == true)
            {
                comboBox1.Items.Add((string)r[0]);
            }
            conn.Close();

        }
        public DialogResult ShowDialog(int id)
        {
            this.id = id;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (Convert.ToInt32(listView1.Items[i].SubItems[3].Text) == id)
                {
                    MessageBox.Show("Такой продукт уже добавлен. \nВыберите другой или добавьте на склад новый.");
                    return this.ShowDialog();
                }
            }
            conn.Open();
            command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok from vessklad, partiya, prodykt where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + id.ToString();
            r = command.ExecuteReader();
            string[] s = new string[r.FieldCount + 2];
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        s[i] = r[i].ToString();
                    }
                    s[r.FieldCount] = id.ToString();
                    s[r.FieldCount + 1] = "";
                    ListViewItem lvi = new ListViewItem(s);
                    listView1.Items.Add(lvi);
                }
            }
            else
            {
                MessageBox.Show("Ни одной записи не найдено!");
                conn.Close();
            }
            conn.Close();
            return this.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("Выберите продукт в главном окне")==DialogResult.OK)
            {
                this.Hide();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                
                if (listView1.SelectedItems.Count == 0&&listView1.Items.Count>1)
                {
                    MessageBox.Show("Выберите строку куда записать!");
                    return;
                }
                if (listView1.Items.Count == 1) listView1.Items[0].Selected = true;
                //else listView1.Items[0].Selected = false;
                if (Convert.ToDecimal(listView1.SelectedItems[0].SubItems[2].Text.Replace('.', ',')) >= Convert.ToDecimal(textBox1.Text.Replace('.', ',')))
                {
                    listView1.SelectedItems[0].SubItems[4].Text = textBox1.Text;
                    textBox1.Text = "";
                }
                else
                {
                    MessageBox.Show("Неверно указан вес!");
                    return;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].SubItems[4].Text == "")
                {
                    MessageBox.Show("Укажите сколько продукта взято!");
                    return;
                }
            }
            //проверка на будущую дату
            if (dateTimePicker1.Value > DateTime.Now)
            {
                MessageBox.Show("Неверно указана дата!");
                return;
            }
            //проверка на соответствие дат
            conn.Open();
            command.CommandText = "select data from vessklad where id=" + Convert.ToInt32(listView1.Items[0].SubItems[3].Text);
            DateTime one = Convert.ToDateTime(command.ExecuteScalar());
            conn.Close();
            if (one > dateTimePicker1.Value) { MessageBox.Show("Несоответствие даты прихода товара и даты его переработки"); return; }
            for(int i=0;i<listView1.Items.Count;i++)
            {
                conn.Open();
                decimal ves1 = 0;
                string data1 = dateTimePicker1.Value.ToShortDateString();
                command.CommandText = "select sobitie.ves,sobitie.data,balans.name,dvig.name from sobitie,balans,dvig where sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and dvig.name!='Окончилось' and sobitie.idsklad="+listView1.Items[i].SubItems[3].Text;
                r = command.ExecuteReader();
                if(r.HasRows)
                    while (r.Read())
                    {
                        DateTime dt = Convert.ToDateTime(r[1].ToString());
                        if (dt <= Convert.ToDateTime(data1))
                        {
                            if ((string)r[2] == "Приход") ves1 += Convert.ToDecimal(r[0].ToString());
                            if ((string)r[2] == "Расход") ves1 -= Convert.ToDecimal(r[0].ToString());
                        }
                    }
                conn.Close();
                if (ves1 < Convert.ToDecimal(listView1.Items[i].SubItems[4].Text.Replace('.', ',')))
                {
                    StreamReader sreader;
                    StreamWriter swriter;                    
                    //parol = "1111";
                    FileInfo err = new FileInfo("err.txt");
                    if (!err.Exists)
                    {
                        swriter = err.CreateText();
                        conn.Open();
                        command.CommandText = "select max(id) from proizv";
                        swriter.WriteLine("Нет наличия на складе на дату. Ид.склада:" + listView1.Items[0].SubItems[3].Text + ", Ид.произв.:" + command.ExecuteScalar().ToString());
                        conn.Close();
                        swriter.Close();
                    }
                    else
                    {
                        sreader = err.OpenText();
                        string tettt = sreader.ReadToEnd();
                        sreader.Close();
                        swriter = err.CreateText();
                        swriter.Write(tettt);
                        conn.Open();
                        command.CommandText = "select max(id) from proizv";
                        swriter.WriteLine("Нет наличия на складе на дату. Ид.склада:" + listView1.Items[0].SubItems[3].Text + ", Ид.произв.:" + command.ExecuteScalar().ToString());
                        conn.Close();
                    }
                }
            }

            if (comboBox4.Text != "" && comboBox5.Text != "" && dateTimePicker1.Text != "")
            {
                //comboBox4.Enabled = false;
                //comboBox5.Enabled = false;
                //dateTimePicker1.Enabled = false;
                //textBox3.Enabled = false;
                comboBox1.Enabled = true;
                textBox2.Enabled = true;
                button4.Enabled = true;
                comboBox2.Enabled = true;
                comboBox4.Enabled = false;
                comboBox5.Enabled = false;
                dateTimePicker1.Enabled = false;
                textBox3.Enabled = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && textBox2.Text != ""&&comboBox2.Text!="")
            {
                for (int i = 0; i < listView2.Items.Count; i++)
                {
                    if (comboBox1.Text == listView2.Items[i].SubItems[0].Text)
                    {
                        MessageBox.Show("Этот продукт уже есть в списке");
                        return;
                    }
                }
                decimal ssum = CheckSum() + Convert.ToDecimal(textBox2.Text.Replace('.',','));
                decimal ish = 0;
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    ish += Convert.ToDecimal(listView1.Items[i].SubItems[4].Text.Replace('.', ','));
                }
                
                
                if(ssum>ish)
                {
                    MessageBox.Show("Невозможно добавить(получено больше чем взято)");
                    return;
                }
                //summapolycheno=0;
                //for (int i = 0; i < listView2.Items.Count; i++)
                //{
                //    summapolycheno = summapolycheno + Convert.ToInt32(listView2.Items[i].SubItems[1].Text);
                //}
                //if (summapolycheno + Convert.ToInt32(textBox2.Text) > Convert.ToInt32(listView1.Items[0].SubItems[4].Text))
                //{
                //    MessageBox.Show("Суммарный вес полученных продуктов превышает \nвес взятого");
                //    return;
                //}
                //summapolycheno = summapolycheno + Convert.ToInt32(textBox2.Text);
                string []s=new string[3];
                s[0] = comboBox1.Text;
                s[1] = textBox2.Text;
                s[2] = comboBox2.Text;
                ListViewItem lvi = new ListViewItem(s);
                listView2.Items.Add(lvi);
                label9.Text= "Итого: " + CheckSum().ToString();
                comboBox1.SelectedIndex = 0;
                textBox2.Text = "";
                button7.Enabled = true;
            }
        }
        private decimal CheckSum()
        {
            decimal sum = 0;
            for (int i = 0; i < listView2.Items.Count; i++)
            {
                sum = sum + Convert.ToDecimal(listView2.Items[i].SubItems[1].Text.Replace('.',','));
            }
            return sum;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (listView2.SelectedItems.Count != 0)
            {
                listView2.SelectedItems[0].Remove();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listView2.Items.Clear();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
           m3: summavzyato = 0;
            summapolycheno = 0;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                summavzyato = summavzyato + Convert.ToDecimal(listView1.Items[i].SubItems[4].Text.Replace('.', ','));
            }
            for (int i = 0; i < listView2.Items.Count; i++)
            {
                summapolycheno = summapolycheno + Convert.ToDecimal(listView2.Items[i].SubItems[1].Text.Replace('.', ','));
            }
            if (summapolycheno < summavzyato)
            {
                if (MessageBox.Show("Потери при переработке составили: " + (summavzyato - summapolycheno) + " кг.\nВыберите Нет если вы хотите изменить список полученных продуктов,\n или Да если согласны с количеством потерь.", "Предупреждение", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    return;
                }
                else
                {
                    ot = new Othodi(conn, (summavzyato - summapolycheno));
                    ot.FormClosing += new FormClosingEventHandler(ot_FormClosing);
                    ot.ShowDialog();
                    goto m3;
                }
            }
            if (summapolycheno > summavzyato)
            {
                if (MessageBox.Show("Суммарный вес полученных продуктов больше взятого на " + (summapolycheno - summavzyato) + " кг.!\nВыберите в поле 'Взято' кнопку 'Добавить' для добавления продукта из склада\nили кнопку 'Новый' если продукта на складе нет", "Предупреждение", MessageBoxButtons.OK) == DialogResult.OK)
                {
                    return;
                }
            }
            conn.Open();
            command.CommandText = "select id from partiya where name='" + listView1.Items[0].SubItems[0].Text + "'";
            int idpartiya = (int)command.ExecuteScalar();
            conn.Close();

            //запись id партии гп
            //conn.Open();
            //command.CommandText = "select id from partiya where name='ГП'";
            //int idpartiyagp = (int)command.ExecuteScalar();
            //conn.Close();

            conn.Open();
            command.CommandText = "select id from rabotnik where surname='" + comboBox5.Text + "'";
            int idrabotnik = (int)command.ExecuteScalar();
            conn.Close();

            conn.Open();
            command.CommandText = "select id from stanok where name='" + comboBox4.Text + "'";
            int idstanok = (int)command.ExecuteScalar();
            conn.Close();

            conn.Open();
            command.CommandText = "select id from sost where name='Переработано'";
            int idsost = (int)command.ExecuteScalar();
            conn.Close();

            decimal z = 0;
            if (textBox3.Text == "") z = 12;
            if (textBox3.Text != "") z = Convert.ToDecimal(textBox3.Text.Replace('.', ','));
           
            int idpermax=0;
            
            try
            {
                conn.Open();
                command.CommandText = "select max(id) from proizv";
                idpermax= (int)command.ExecuteScalar();
                conn.Close();
            }
            catch (System.Exception)
            {
                conn.Close();
                conn.Open();
                command.CommandText = "insert into proizv(id) values(0)";
                command.ExecuteNonQuery();
                conn.Close();
                idpermax = 0;
            }
            conn.Close();
            idpermax++;
            conn.Open();
            command.CommandText = "insert into proizv(id)values("+idpermax+")";
            command.ExecuteNonQuery();
            conn.Close();

            //добавляем все сырье в табл. сырье
            //conn.Open();
            //for (int i = 0; i < listView1.Items.Count; i++)
            //{                
            //    command.CommandText = "insert into sirie values("+idpermax+","+listView1.Items[i].SubItems[3].Text+")";
            //    command.ExecuteNonQuery();
            //}
            //conn.Close();
                //перечисляем исходное и делаем с ним все
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    decimal ost = 0;// остаток сырья
                    ost = Convert.ToDecimal(listView1.Items[i].SubItems[2].Text.Replace('.', ',')) - Convert.ToDecimal(listView1.Items[i].SubItems[4].Text.Replace('.', ','));
                    conn.Open();
                    command.CommandText = "update vessklad set ostatok=" + ost.ToString().Replace(',', '.') + " where id=" + Convert.ToInt32(listView1.Items[i].SubItems[3].Text);
                    command.ExecuteNonQuery();
                    //conn.Close();

                    
                    //command.CommandText = "insert into sirie values(" + idpermax + "," + listView1.Items[i].SubItems[3].Text + ")";
                    //command.ExecuteNonQuery();
                    //conn.Open();
                    //command.CommandText = "update proizv set idsirie=" + Convert.ToInt32(listView1.Items[i].SubItems[3].Text) + " where id=" + idpermax;
                    //command.ExecuteNonQuery();
                    //conn.Close();

                    //conn.Open();
                    command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + Convert.ToInt32(listView1.Items[i].SubItems[3].Text) + "," + listView1.Items[i].SubItems[4].Text.Replace(',', '.') + ",2,2,2," + idpermax + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                    command.ExecuteNonQuery();
                    conn.Close();
                    if (ost == 0)
                    {
                        conn.Open();
                        command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime)values (" + Convert.ToInt32(listView1.Items[i].SubItems[3].Text) + "," + listView1.Items[i].SubItems[4].Text.Replace(',', '.') + ",2,4,2," + idpermax + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                        command.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            ///// продолжить!!!!! вставлять ид в событие
            //перечисляем получено и делаем ему все
                for (int i = 0; i < listView2.Items.Count; i++)
                {
                    if (listView2.Items[i].SubItems[2].Text != "Отходы")
                    {
                        conn.Open();
                        command.CommandText = "select id from prodykt where name='" + listView2.Items[i].SubItems[0].Text + "'";
                        int idprodykt = (int)command.ExecuteScalar();
                        conn.Close();

                        conn.Open();
                        command.CommandText = "select id from state where name='" + listView2.Items[i].SubItems[2].Text + "'";
                        int idstate = (int)command.ExecuteScalar();
                        conn.Close();
                        //объединение всего ГП
                        int prttt = idpartiya;
                        //if (listView2.Items[i].SubItems[2].Text == "Готовая продукция")
                        //{
                        //    prttt = idpartiyagp;
                        //}

                        conn.Open();
                        command.CommandText = "insert into vessklad(idpartiya,idprodykt,nachves,ostatok,idrabotnik,idstanok,data,idstate,idsost,recordtime) values (" + prttt + "," + idprodykt + "," + listView2.Items[i].SubItems[1].Text.Replace(',', '.') + "," + listView2.Items[i].SubItems[1].Text.Replace(',', '.') + "," + idrabotnik + "," + idstanok + ",'" + dateTimePicker1.Text + "'," + idstate + "," + idsost + ",'" + DateTime.Now.ToString() + "')";
                        if (command.ExecuteNonQuery() == 0)
                        {
                            conn.Close();
                            MessageBox.Show("Ошибка добавления");
                            return;
                        }
                        conn.Close();
                        conn.Open();
                        command.CommandText = "select max(id) from vessklad";
                        int idd = (int)command.ExecuteScalar();
                        conn.Close();
                        conn.Open();
                        command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + idd + "," + listView2.Items[i].SubItems[1].Text.Replace(',', '.') + ",2,2,1," + idpermax + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                        command.ExecuteNonQuery();
                        conn.Close();
                        conn.Open();
                        command.CommandText = "update proizv set idprodykt"+(i+1)+"="+idd+" where id="+idpermax;
                        command.ExecuteNonQuery();
                        conn.Close();
                        //проверка и объединение одинаковых позиций
                        conn.Open();
                        //int x1 = Convert.ToInt32(listView1.Items[i].SubItems[3].Text);                                             
                        command.CommandText = "select id,ostatok,sostav from vessklad where idpartiya=" + prttt + " and idprodykt=" + idprodykt + " and idstate=" + idstate + " and idsost=" + idsost;
                        r = command.ExecuteReader();
                        int idsostav = -1;
                        ArrayList same = new ArrayList();
                        if (r.HasRows == true)
                        {
                            while (r.Read() == true)
                            {
                                try
                                {
                                    if (Convert.ToDecimal(r[1].ToString().Replace('.', ',')) == 0 && !(bool)r[2]) continue;
                                }
                                catch (System.Exception)
                                {
                                    continue;
                                }
                                try
                                {
                                    if ((bool)r[2]) { idsostav = (int)r[0]; continue; }
                                }
                                catch (System.Exception)
                                {
                                    
                                }
                                decimal[] ss = new decimal[2];
                                ss[0] = Convert.ToDecimal(r[0].ToString().Replace('.',','));
                                ss[1] = Convert.ToDecimal(r[1].ToString().Replace('.', ','));
                                same.Add(ss);
                            }
                        }
                        conn.Close();

                        bool checker = false;
                        if (same.Count == 1 && idsostav == -1) checker = false;
                        if (same.Count > 1) checker = true;
                        if (same.Count >= 1 && idsostav != -1) checker = true;
                        if(checker)
                        {
                            int idnew = 0;//объединяющий элемент
                            int idsobun = 0;//айди события объединенного
                            int idpereun = 0;//айди переработки объединения
                            //int vessost = 0;//вес объединяющего
                            decimal sum = 0;
                            decimal[] sumar = new decimal[2];
                            if (idsostav == -1)//нет объдиняющего элемента
                            {
                                //добавление объдиняющего элемента
                                //sumar=(decimal[])same[0];
                                //conn.Open();
                                //command.CommandText = "select data from vessklad where id="+sumar[0];
                                //string datastart = (string)command.ExecuteScalar();
                                //conn.Close();
                                conn.Open();//вставляем пустой составной
                                command.CommandText = "insert into vessklad(idpartiya,idprodykt,nachves,ostatok,idrabotnik,idstanok,idstate,idsost,recordtime,sostav) values (" + prttt + "," + idprodykt + "," + 0 + "," + 0 + ","+idrabotnik+","+idstanok+"," + idstate + "," + idsost + ",'" + DateTime.Now.ToString() + "','1')";
                                command.ExecuteNonQuery();
                                //conn.Close();
                                //conn.Open();
                                command.CommandText = "select max(id) from vessklad";
                                idnew = (int)command.ExecuteScalar();
                                //заполнение таблицы производство
                                command.CommandText = "select max(id) from proizv";
                                idpereun = (int)command.ExecuteScalar();
                                idpereun++;
                                command.CommandText = "insert into proizv (id,idsirie) values(" + idpereun + ", " + idnew + ")";
                                command.ExecuteNonQuery();
                                //for (int k = 0; k < same.Count; k++)
                                //{
                                //    sumar = (decimal[])same[k];
                                //    command.CommandText = "update proizv set idprodykt" + (k + 1) + "=" + sumar[0] + " where id=" + (permax + 1);
                                //    command.ExecuteNonQuery();
                                //}
                                //command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + idnew + "," + sum.ToString().Replace(',', '.') + ",2,5,1," + idpereun + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                                //command.ExecuteNonQuery();
                                //command.CommandText = "select max(id) from sobitie";
                                //idsobun = (int)command.ExecuteScalar();//сщбытие создания объединенного нам пока не нужно
                                conn.Close();
                            }
                            else 
                            {                                
                                idnew = idsostav;
                                conn.Open();
                                command.CommandText = "select ostatok from vessklad where id="+idnew;
                                if (Convert.ToInt32(command.ExecuteScalar()) == 0)
                                {
                                    command.CommandText = "delete from sobitie where idsklad="+idnew+" and iddvig=4 and idbalans=2";
                                    try
                                    {
                                        command.ExecuteNonQuery();
                                    }
                                    catch (System.Exception)
                                    {
                                    }
                                }
                                conn.Close();
                                conn.Open();
                                command.CommandText = "select id, idproizv from sobitie where iddvigfrom=2 and iddvig=5 and idbalans=1 and idsklad="+idnew.ToString();
                                r = command.ExecuteReader();
                                if (r.HasRows)
                                {
                                    r.Read();
                                    idsobun = (int)r[0];
                                    idpereun = (int)r[1];
                                }
                                conn.Close();
                                conn.Open();
                                command.CommandText = "select ostatok from vessklad where id=" + idnew.ToString();
                                sum = Convert.ToDecimal(command.ExecuteScalar().ToString().Replace('.', ','));
                                conn.Close();
                            }
                            decimal ostpohogih = 0;
                            for (int u = 0; u < same.Count;u++ )
                            {
                                sumar = (decimal[])same[u];
                                conn.Open();
                                command.CommandText = "update vessklad set ostatok=0 where id="+sumar[0];//остаток=0
                                command.ExecuteNonQuery();
                                command.CommandText = "update vessklad set sostav=0 where id=" + sumar[0];//оно составное
                                command.ExecuteNonQuery();//вставка в событие что оно объединилось
                                command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + sumar[0] + "," + sumar[1].ToString().Replace(',', '.') + ",2,5,2," + idpereun + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                                command.ExecuteNonQuery();
                                //поиск последнего в производстве

                                int idprodMax = 0;
                                //выбираем максимальный из произв
                                command.CommandText = "select * from proizv where id=" + idpereun;
                                r = command.ExecuteReader();

                                if (r.HasRows)
                                {
                                    while (r.Read())
                                    {
                                        for (int kk = 0; kk < r.FieldCount; kk++)
                                        {
                                            try
                                            {
                                                int err = (int)r[kk + 2];
                                            }
                                            catch (System.Exception)
                                            {
                                                idprodMax = kk + 1;
                                                //conn.Close();
                                                break;
                                            }                                            
                                        }
                                    }
                                }
                                conn.Close();
                                conn.Open();
                                //добавление в производство
                                try
                                {

                                    command.CommandText = "update proizv set idprodykt" + idprodMax + "=" + sumar[0] + " where id=" + idpereun;
                                    command.ExecuteNonQuery();
                                }
                                catch (System.Exception)
                                {
                                    command.CommandText = "ALTER TABLE proizv ADD idprodykt"+idprodMax+" int CONSTRAINT fk_prodykt"+idprodMax+"_element FOREIGN KEY  REFERENCES vessklad(id)on delete no action on update no action";
                                    command.ExecuteNonQuery();
                                    command.CommandText = "update proizv set idprodykt" + idprodMax + "=" + sumar[0] + " where id=" + idpereun;
                                    command.ExecuteNonQuery();
                                }
                                ostpohogih += sumar[1];//суммируем остатки  остальных
                                //узнаем дату рождения каждого
                                command.CommandText = "select data from vessklad where id="+sumar[0].ToString();
                                string datakajdogo = (string)command.ExecuteScalar();
                                //добавляем в событие каждый из объединяемых что он добавился к объединяющему с датой
                                command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + idnew + "," + sumar[1].ToString().Replace(',', '.') + ",2,5,1," + idpereun + ",'" + datakajdogo + "','" + DateTime.Now.ToString() + "')";
                                command.ExecuteNonQuery();
                                string datapervogo = "";
                                if(u==0&&same.Count>1)
                                {
                                    command.CommandText = "select data from vessklad where id="+sumar[0];
                                    datapervogo = (string)command.ExecuteScalar();
                                    command.CommandText = "update vessklad set data='"+datapervogo+"' where id="+idnew;
                                    command.ExecuteNonQuery();
                                    command.CommandText = "update vessklad set nachves="+sumar[1].ToString().Replace(',','.')+" where id="+idnew;
                                    command.ExecuteNonQuery();
                                }
                                conn.Close();
                            }
                            //делаем все для объединяющего(изменяем остаток,вставляем в событие что к нему добавилось)
                            conn.Open();
                            decimal sumobch = sum + ostpohogih;
                            command.CommandText = "update vessklad set ostatok=" + sumobch.ToString().Replace(',', '.') + " where id=" + idnew;//остаток=0
                            command.ExecuteNonQuery();//в событии только изменять вес!
                            //command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + idnew + "," + ostpohogih.ToString().Replace(',', '.') + ",2,5,1," + idpereun + ",'" + dateTimePicker1.Text + "','" + DateTime.Now.ToString() + "')";
                            //command.ExecuteNonQuery();
                            //command.CommandText = "update vessklad";
                            conn.Close();

                        }
                        //конец нового
                    }                     
                    if (listView2.Items[i].SubItems[2].Text == "Отходы")
                    {
                        conn.Open();
                        command.CommandText = "select id from tipothoda where name='" + listView2.Items[i].SubItems[0].Text + "'";
                        int idtipothoda = (int)command.ExecuteScalar();
                        conn.Close();

                        conn.Open();
                        command.CommandText = "insert into othodi(idproizv,idtipothoda,ves,recordtime) values("+idpermax+"," + idtipothoda + "," + listView2.Items[i].SubItems[1].Text.Replace(',','.') + ",'" + DateTime.Now.ToString() + "')";
                        command.ExecuteNonQuery();
                        conn.Close();
                    }
                }
            MessageBox.Show("Переработка завершена!!");
            
            this.Close();
        }

        void ot_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ot.musor != "" && ot.vesmusor != 0)
            {
                string[] s = new string[3];
                s[0] = ot.musor;
                s[1] = ot.vesmusor.ToString();
                s[2] = "Отходы";
                ListViewItem lvi = new ListViewItem(s);
                listView2.Items.Add(lvi);
                return;
            }
            else return;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            zap = new Zapolnenie(conn);
            zap.Show();
            zap.FormClosing += new FormClosingEventHandler(zap_FormClosing);
        }

        void zap_FormClosing(object sender, FormClosingEventArgs e)
        {
            id = zap.id;
            if (id != 0)
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok from vessklad, partiya, prodykt where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + id.ToString();
                r = command.ExecuteReader();
                string[] s = new string[r.FieldCount + 2];
                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        for (int i = 0; i < r.FieldCount; i++)
                        {
                            s[i] = r[i].ToString();
                        }
                        s[r.FieldCount] = id.ToString();
                        s[r.FieldCount + 1] = "";
                        ListViewItem lvi = new ListViewItem(s);
                        listView1.Items.Add(lvi);
                    }
                }
                else
                {
                    MessageBox.Show("Ни одной записи не найдено!");
                    conn.Close();
                }
                conn.Close();
            }
        }

        private void listView2_EnabledChanged(object sender, EventArgs e)
        {
            MessageBox.Show("fgfghf");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                listView1.SelectedItems[0].Remove();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                listView1.Items[i].SubItems[4].Text = listView1.Items[i].SubItems[2].Text;
            }
        }
    }
}