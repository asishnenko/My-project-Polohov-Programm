using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Collections;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Polohov
{
    public partial class Form1 : Form
    {
        static public System.Data.SqlClient.SqlConnection conn;
        static public bool prava;
        static public bool fwd;
        static public string user;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        AddProd1 adpr1;
        AddProd1 addka;
        zadanie zad;
        Zapolnenie zap;
        Prodaja pr;
        Pasting ps;
        //Form2 f2;
        SostSklad sostSklad;
        Union un;
        Rabochie rab;
        delete del;
        SetPartiya setp;
        Password pwd;
        Tree tree;
        deletepartiya delp;
        Starting starting;
        changelogpwd chlog;
        changepwd chpwd;
        ResetSklad reset;
        int perekl;
        int perekl1;
        static public int idstanok;
        static public int idrabotnik;
        static public bool price;
        //int w;
        int krai;
        int iter;
        int idskladpartiya;
        //static public string parol;
        
        public Form1()
        {
            command.Connection = conn;
            perekl = 0;
            perekl1 = 0;
          
            InitializeComponent();
            label4.Text = "";
            //w = dataGridView1.Width / 2;
            //label3.Text = "";
            krai = dataGridView1.Location.X+15;
            //dataGridView1.Width = ClientRectangle.Width /2;
            fwd = false;
            prava = false;
            price = false;
            starting = new Starting();
            starting.FormClosed += new FormClosedEventHandler(starting_FormClosed);
            starting.ShowDialog();
            if (fwd) { Start(); Text = "Склад -"+user+"-"; }
            else
            {
                MessageBox.Show("Авторизуйтесь! Программа не сможет работать!");
            }
            

  
        }
        void starting_FormClosed(object sender, FormClosedEventArgs e)
        {
            //return;
        }
        public void Start()//начало инициализация
        {
            StreamReader sreader;
            StreamWriter swriter;
            //parol = "1111";
            FileInfo pwd = new FileInfo("oll");
            if (!pwd.Exists)
            {
                swriter = pwd.CreateText();
                swriter.WriteLine("1111");
                swriter.Close();
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

            FileInfo opt = new FileInfo("opt.txt");
            if (!opt.Exists)
            {
                хранитьЦенуToolStripMenuItem.Checked = false;
                price = false;
            }
            else
            {
                sreader = opt.OpenText();
                string op = sreader.ReadLine();
                if (op == "price:on")
                {
                    хранитьЦенуToolStripMenuItem.Checked = true;
                    price = true;
                }
                if (op == "price:off")
                {
                    хранитьЦенуToolStripMenuItem.Checked = false;
                    price = false;
                }
            }

           

            //первый раз делаем вставку в произв
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select max(id) from proizv";
            try
            {
                command.ExecuteScalar();
            }
            catch (System.Exception)
            {
                command.CommandText = "insert into proizv(id) values (1)";
                command.ExecuteScalar();
            }
            conn.Close();

            //делаем партию ГП
            conn.Open();
            command.CommandText = "select id from partiya where name='ГП'";
            r = command.ExecuteReader();
            bool rezz=r.HasRows;
            conn.Close();
            if (!rezz)
            {
                conn.Open();
                command.CommandText = "insert into partiya(name) values('ГП')";
                command.ExecuteNonQuery();
                conn.Close();
            }
            


            //Vhod v = new Vhod(conn);
            //v.ShowDialog();


            comboBox1.Items.Add("Весь склад");
            comboBox1.Items.Add("Готовая продукция");
            comboBox1.Items.Add("Склад производства");
            //conn.Open();
            //command.CommandText = "select name from state";
            //command.Connection = conn;
            //r = command.ExecuteReader();
            //if (r.HasRows == true)
            //{
            //    while (r.Read() == true)
            //    {

            //        if ((string)r[0] == "Продано" || (string)r[0] == "В работе") continue;
            //        comboBox1.Items.Add((string)r[0]);
            //    }
            //}
            //conn.Close();

            CheckPartiyaEnd();
            CheckPartiyaKon();

            //заполнение комбо партий для фильтра
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select name from partiya where net=0";
            comboBox3.Items.Add("Все");
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox3.Items.Add((string)r[0]);
                }
            }
            conn.Close();

            comboBox3.SelectedIndex = 0;
            comboBox1.SelectedIndex = 0;

            conn.Open();
            command.CommandText = "select id from stanok where name=''";
            try { idstanok = (int)command.ExecuteScalar(); }
            catch (System.Exception)
            {
                conn.Close();
                conn.Open();
                command.CommandText = "insert into stanok values('')";
                command.ExecuteScalar();
                conn.Close();
                conn.Open();
                command.CommandText = "select id from stanok where name=''";
                idstanok = (int)command.ExecuteScalar();
            }
            conn.Close();

            conn.Open();
            command.CommandText = "select id from rabotnik where name=''";
            try { idrabotnik = (int)command.ExecuteScalar(); }
            catch (System.Exception)
            {
                conn.Close();
                conn.Open();
                command.CommandText = "insert into rabotnik(name,surname) values('','')";
                command.ExecuteScalar();
                conn.Close();
                conn.Open();
                command.CommandText = "select id from rabotnik where name=''";
                idrabotnik = (int)command.ExecuteScalar();
            }
            conn.Close();

            comboBox2.Items.Add("Дерево");
            comboBox2.Items.Add("Подробно текст");
            comboBox2.Items.Add("Кратко текст");
        }
        public void CheckPartiyaEnd()//проверяет и устанавливает переключатель партии ..нет на складе производства ничего(но возможно есть на гп и продано)
        {
            //conn.Open();
            //command.CommandText = "select konetc from partiya where name ="+partiya;
            //if ((int)command.ExecuteScalar()==0)
            //{
            //    command.CommandText = "select vessklad.id from vessklad,state where state.name='Готовая продукция' and vessklad.idstate=state.id and vessklad.ostatok!=0";
            //    r = command.ExecuteReader();
            //    if (!r.HasRows)
            //    {
            //        conn.Close();
            //        conn.Open();
            //        command.CommandText = "update partiya set konetc=1 where name="+partiya;
            //    }
            //}
            //conn.Close();
            conn.Close();
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select name from partiya where konetc=0 and name!='ГП'";
            r = command.ExecuteReader();
            ArrayList ttm = new ArrayList();
            if (r.HasRows)
            {                
                while (r.Read())
                {
                    ttm.Add(r[0]);
                }
            }
            conn.Close();
            
            for (int i = 0; i < ttm.Count; i++)
            {
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,state,partiya where partiya.name='"+(string)ttm[i]+"' and state.name='Склад производства' and vessklad.idstate=state.id and vessklad.idpartiya=partiya.id and vessklad.ostatok!=0";
                r = command.ExecuteReader();
                if (!r.HasRows)
                {
                    conn.Close();
                    conn.Open();
                    command.CommandText = "update partiya set konetc=1 where name='" + (string)ttm[i]+"'";
                    command.ExecuteNonQuery();
                    //conn.Open();
                    int idp = 0;
                    command.CommandText = "select id from partiya where name='"+(string)ttm[i]+"'";
                    idp = (int)command.ExecuteScalar();
                    command.CommandText = "insert into sobitiepartii values(" + idp + ",2,'" + DateTime.Now.ToShortDateString() + "')";
                    command.ExecuteNonQuery();
                    //conn.Close();
                    conn.Close();
                }
                conn.Close();
            }
            //установка обратного
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select name from partiya where konetc=1 and name!='ГП'";
            r = command.ExecuteReader();
            ttm = new ArrayList();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    ttm.Add(r[0]);
                }
            }
            conn.Close();

            for (int i = 0; i < ttm.Count; i++)
            {
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,state,partiya where partiya.name='" + (string)ttm[i] + "' and state.name='Склад производства' and vessklad.idstate=state.id and vessklad.idpartiya=partiya.id and vessklad.ostatok!=0";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    conn.Close();
                    conn.Open();
                    command.CommandText = "update partiya set konetc=0 where name='" + (string)ttm[i] + "'";
                    command.ExecuteNonQuery();
                    int idp1 = 0;
                    command.CommandText = "select id from partiya where name='" + (string)ttm[i] + "'";
                    idp1 = (int)command.ExecuteScalar();
                    command.CommandText = "delete from sobitiepartii where idpartiya="+idp1+" and idnamesobpar=2";
                    command.ExecuteNonQuery();
                    conn.Close();
                }
                conn.Close();
            }
            
        }
        public void CheckPartiyaKon()//проверяет и устанавливает переключатель партии на конец, т.е. на гп и на складе произв ничего нет из этой партии
        {
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select name from partiya where net=0 and konetc=1 and name!='ГП'";
            r = command.ExecuteReader();
            ArrayList ttm = new ArrayList();
            if (r.HasRows)
            {                
                while (r.Read())
                {
                    ttm.Add(r[0]);
                }
            }
            conn.Close();

            for (int i = 0; i < ttm.Count; i++)
            {
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,partiya where partiya.name='" + (string)ttm[i] + "' and vessklad.idpartiya=partiya.id and vessklad.ostatok!=0";
                r = command.ExecuteReader();
                if (!r.HasRows)
                {
                    conn.Close();
                    conn.Open();
                    command.CommandText = "update partiya set net=1 where name='" + (string)ttm[i] + "'";
                    command.ExecuteNonQuery();
                    int idp = 0;
                    command.CommandText = "select id from partiya where name='" + (string)ttm[i] + "'";
                    idp = (int)command.ExecuteScalar();
                    try
                    {
                        command.CommandText = "insert into sobitiepartii values(" + idp + ",3,'" + DateTime.Now.ToShortDateString() + "')";
                        command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        command.CommandText = "update sobitiepartii set data=" + DateTime.Now.ToShortDateString() + " where idpartiya=" + idp + " and idnamesobpar=3";
                        command.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                conn.Close();
            }

            //Проверка на обратное
            conn.Open();
            command.Connection = conn;
            command.CommandText = "select name from partiya where net=1 and name!='ГП'";
            r = command.ExecuteReader();
            ttm = new ArrayList();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    ttm.Add(r[0]);
                }
            }
            conn.Close();
            for (int i = 0; i < ttm.Count; i++)
            {
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,partiya where partiya.name='" + (string)ttm[i] + "' and vessklad.idpartiya=partiya.id and vessklad.ostatok!=0";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    conn.Close();
                    conn.Open();
                    command.CommandText = "update partiya set net=0 where name='" + (string)ttm[i] + "'";
                    command.ExecuteNonQuery();
                    int idp = 0;
                    command.CommandText = "select id from partiya where name='" + (string)ttm[i] + "'";
                    idp = (int)command.ExecuteScalar();
                    command.CommandText = "update sobitiepartii set data='' where idpartiya=" + idp + " and idnamesobpar=3";
                    command.ExecuteNonQuery();
                    conn.Close();
                }
                conn.Close();
            }


        }
        public void SqlZapros(string comand, string label, bool id)
        {
            dataGridView1.Width = ClientRectangle.Width-krai;
            richTextBox1.Visible = false;
            
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            conn.Open();
            command.Connection = conn;            
            command.CommandText = comand;            
            r = command.ExecuteReader();
            dataGridView1.Columns.Add("", "");
            for (int i = 0; i < r.FieldCount; i++)
            {
                dataGridView1.Columns.Add(r.GetName(i), r.GetName(i));
            }
            int k = 1;
            string[] s = new string[r.FieldCount + 1];
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    s[0] = k.ToString();
                    for (int i = 0; i < r.FieldCount; i++)
                    {
                        s[i + 1] = r[i].ToString();
                    }
                    if (s[1] != "")
                    {

                        //try
                        //{
                            dataGridView1.Rows.Add(s);
                        //}
                        //catch (System.Exception)
                        //{
                        //    dataGridView1.Rows.SharedRow(dataGridView1.Rows.Count);
                        //    dataGridView1.Rows.Add(s);

                        //}
                        //if (dataGridView1.Rows.Count > 1)
                        //{
                        //    if (dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[dataGridView1.ColumnCount - 1].Value.ToString() == dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[dataGridView1.ColumnCount - 1].Value.ToString())
                        //        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 1);
                        //}
                    }
                    k++;
                }
            }
                
            else
            {
                //MessageBox.Show("Ни одной записи не найдено!");
                conn.Close();
                //return;
            }
            conn.Close();
            if(id==true)dataGridView1.Columns[dataGridView1.Columns.Count-1].Visible = false;
            label2.Text = label;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Enabled = true;
            comboBox3.Enabled = true;
            button3.Enabled = true;
            string s = "";
            string s1 = "";
            string s2 = "";
            CheckPartiyaEnd();
            CheckPartiyaKon();

            //заполнение комбо партий для фильтра
            conn.Open();
            command.Connection = conn;
            comboBox3.Items.Clear();
            command.CommandText = "select name from partiya where net=0";
            comboBox3.Items.Add("Все");
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    comboBox3.Items.Add((string)r[0]);
                }
            }
            conn.Close();
            //if (price == true) { s = " ,sobitie.price as Цена "; s1 = " ,sobitie "; s2 = " and sobitie.idsklad=vessklad.id "; }
            if ((string)comboBox1.SelectedItem == "Весь склад") 
            { 
                SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(переработки)], state.name as Положение  " + s + ",vessklad.id from partiya,prodykt,vessklad, state, sost" + s1 + " where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstate=state.id and vessklad.idsost=sost.id " + s2 + " and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                perekl = 0;
                perekl1 = 0;
                comboBox3.SelectedIndex = 0;
                textBox1.Text = "";
            }
            //if ((string)comboBox1.SelectedItem == "Продано") SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], stanok.name as Станок, rabotnik.surname as Рабочий, vessklad.zatracheno as [Затрачено времени(час)], sost.name as Положение, vessklad.id from partiya,prodykt,stanok,rabotnik,sost,vessklad where vessklad.idstate=" + (comboBox1.SelectedIndex) + " and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstanok=stanok.id and vessklad.idrabotnik=rabotnik.id and vessklad.idsost=sost.id order by partiya.name asc", comboBox1.Text, true);
            if ((string)comboBox1.SelectedItem == "Склад производства") 
            { 
                SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(переработки)], vessklad.id from partiya,prodykt,sost,vessklad,state where state.name ='" + (comboBox1.Text) + "' and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idsost=sost.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                perekl = 0;
                perekl1 = 0;
                comboBox3.SelectedIndex = 0;
                textBox1.Text = "";
            }
            if ((string)comboBox1.SelectedItem == "Готовая продукция")
            {
                SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата начала поступления на ГП], vessklad.id from partiya,prodykt,sost,vessklad,state where state.name='"+comboBox1.Text+"' and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idsost=sost.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                perekl = 0;
                perekl1 = 0;
                comboBox3.SelectedIndex = 0;
                textBox1.Text = "";
            }
            if ((string)comboBox1.SelectedItem == "Исходное сырье") { SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления], state.name as Положение, vessklad.id from partiya,prodykt,state,vessklad where vessklad.idstate= state.id and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and ostatok!=0 and state.name='" + comboBox1.Text + "' order by partiya.name asc", comboBox1.Text, true); }
            //if ((string)comboBox1.SelectedItem == "В работе") SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], stanok.name as Станок, rabotnik.surname as Рабочий, vessklad.zatracheno as [Затрачено времени(час)], sost.name as Положение, vessklad.id from partiya,prodykt,stanok,rabotnik,sost,vessklad where vessklad.idstate=" + (comboBox1.SelectedIndex) + " and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstanok=stanok.id and vessklad.idrabotnik=rabotnik.id and vessklad.idsost=sost.id order by partiya.name asc", comboBox1.Text, true);
            if ((string)comboBox1.SelectedItem == "Не выполнено") { SqlZapros("select zadanie.number as [Номер задания], zadanie.name as Описание, partiya.name as Партия,prodykt.name as Наименование, zadanie.datavidachi as [Дата выдачи], zadanie.id from zadanie, partiya, prodykt, vessklad where zadanie.idsklad = vessklad.id and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and sost=0", "Не выполнено", true); }
            if ((string)comboBox1.SelectedItem == "Выполнено") { SqlZapros("select zadanie.number as [Номер задания], zadanie.name as Описание, partiya.name as Партия,prodykt.name as Наименование, zadanie.datavidachi as [Дата выдачи], zadanie.datazaversheniya as [Дата завершения], zadanie.id from zadanie, partiya, prodykt, vessklad where zadanie.idsklad = vessklad.id and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and sost=1", "Выполнено", true);}
            if ((string)comboBox1.SelectedItem == "Не оконченные") {SqlZapros("select partiya.name as Партия,prodykt.name as Продукт from vessklad, partiya,prodykt,sost where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sost.id=vessklad.idsost and vessklad.ostatok!=0 and sost.name='Новое'", "Партия", false);}
            if ((string)comboBox1.SelectedItem == "Оконченные") {SqlZapros("select partiya.name as Партия,prodykt.name as Продукт from vessklad, partiya,prodykt,sost where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and sost.id=vessklad.idsost and vessklad.ostatok=0 and sost.name='Новое'", "Партия", false);}
            perekl = 1;
            perekl1 = 1;
        }
        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e) //добавить значит отобразить склад
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Весь склад");
            comboBox1.Items.Add("Готовая продукция");
            comboBox1.Items.Add("Склад производства");
            textBox1.Enabled = true;
            comboBox3.Enabled = true;
            button3.Enabled = true;
            //conn.Open();
            //command.CommandText = "select name from state";
            //command.Connection = conn;
            //r = command.ExecuteReader();
            //if (r.HasRows == true)
            //{
            //    while (r.Read() == true)
            //    {
            //        if (r[0].ToString() == "Продано" || r[0].ToString() == "В работе") continue;
            //        comboBox1.Items.Add(r[0].ToString());
            //    }
            //}
            //conn.Close();                 

            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = true;
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
        }
        void adr_FormClosed(object sender, FormClosedEventArgs e)
        {
            показатьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите строку!");
                return;
            }     
            if (button1.Text == "Удалить работника")
            {
                           
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    command.CommandText = "delete from rabotnik where working=0";
                    conn.Open();
                    try
                    {
                        command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        conn.Close();
                    }
                    conn.Close();
                    int ss = 0;
                    conn.Open();
                    command.CommandText = "delete from rabotnik where name='" + dataGridView1.SelectedRows[i].Cells[1].Value.ToString() + "' and surname='" + dataGridView1.SelectedRows[i].Cells[2].Value.ToString() + "'";
                    try
                    {
                        ss = command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        conn.Close();
                        command.CommandText = "update rabotnik set working=0 where name='" + dataGridView1.SelectedRows[i].Cells[1].Value.ToString() + "' and surname='" + dataGridView1.SelectedRows[i].Cells[2].Value.ToString() + "'";
                        conn.Open();
                        command.ExecuteNonQuery();
                    }
                    if (ss == 0)
                    {
                        conn.Close();
                        command.CommandText = "update rabotnik set working=0 where name='" + dataGridView1.SelectedRows[i].Cells[1].Value.ToString() + "' and surname='" + dataGridView1.SelectedRows[i].Cells[2].Value.ToString() + "'";
                        conn.Open();
                        command.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                показатьToolStripMenuItem_Click(dataGridView1, new EventArgs());
                if (MessageBox.Show("Удалено успешно.\nУдалить ещё?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    button1.Visible = false;
                }
            }
            if (button1.Text == "Удалить продукт")
            {
                int n = 0;
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "delete from prodykt where name='"+dataGridView1.SelectedRows[i].Cells[1].Value.ToString()+"'";
                    int ss = 0;
                    try
                    {
                        ss = command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        int ind = dataGridView1.SelectedRows[i].Index + 1;
                        conn.Close();
                        MessageBox.Show("Запись "+ind.ToString()+" нельзя удалить, т.к. она используется в записях.\nВозможно только ИЗМЕНЕНИЕ наименования этого продукта.");
                        continue;
                    }
                    if (ss != 0)
                    {
                        n++;
                    }
                    conn.Close();
                }
                показатьToolStripMenuItem1_Click(dataGridView1, new EventArgs());
                MessageBox.Show("Удалено "+n.ToString()+" записей");
            }

            if (button1.Text == "Удалить контрагент")
            {
                int n = 0;
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "delete from kAgent where name='" + dataGridView1.SelectedRows[i].Cells[1].Value.ToString() + "'";
                    int ss = 0;
                    try
                    {
                        ss = command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        int ind = dataGridView1.SelectedRows[i].Index + 1;
                        conn.Close();
                        MessageBox.Show("Запись " + ind.ToString() + " нельзя удалить, т.к. она используется в записях.\nВозможно только ИЗМЕНЕНИЕ наименования этого продукта.");
                        continue;
                    }
                    if (ss != 0)
                    {
                        n++;
                    }
                    conn.Close();
                }
                показатьToolStripMenuItem5_Click(dataGridView1, new EventArgs());
                MessageBox.Show("Удалено " + n.ToString() + " записей");
            }

            if (button1.Text == "Удалить оборудование")
            {
                int n = 0;
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "delete from stanok where name='" + dataGridView1.SelectedRows[i].Cells[1].Value.ToString() + "'";
                    int ss = 0;
                    try
                    {
                        ss = command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        int ind = dataGridView1.SelectedRows[i].Index + 1;
                        conn.Close();
                        MessageBox.Show("Запись " + ind.ToString() + " нельзя удалить, т.к. она используется в записях.\nВозможно только ИЗМЕНЕНИЕ наименования этого продукта.");
                        continue;
                    }
                    if (ss != 0)
                    {
                        n++;
                    }
                    conn.Close();
                }
                показатьToolStripMenuItem2_Click(dataGridView1, new EventArgs());
                MessageBox.Show("Удалено " + n.ToString() + " записей");
            }
            if (button1.Text == "Удалить виды отходов")
            {
                int n = 0;
                for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "delete from tipothoda where name='" + dataGridView1.SelectedRows[i].Cells[1].Value.ToString() + "'";
                    int ss = 0;
                    try
                    {
                        ss = command.ExecuteNonQuery();
                    }
                    catch (System.Exception)
                    {
                        int ind = dataGridView1.SelectedRows[i].Index + 1;
                        conn.Close();
                        MessageBox.Show("Запись " + ind.ToString() + " нельзя удалить, т.к. она используется в записях.\nВозможно только ИЗМЕНЕНИЕ наименования этого продукта.");
                        continue;
                    }
                    if (ss != 0)
                    {
                        n++;
                    }
                    conn.Close();
                }
                показатьToolStripMenuItem4_Click(dataGridView1, new EventArgs());
                MessageBox.Show("Удалено " + n.ToString() + " записей");
            }
            if (button1.Text == "Добавить")
            {
                zad.ShowDialog(Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[dataGridView1.Columns.Count-1].Value));
            }
            if (button1.Text == "Добавить ")
            {
                un.ShowDialog(Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[dataGridView1.Columns.Count - 1].Value));
            }
            if (button1.Text == "Переработать")
            {
                zad=new zadanie(conn);
                zad.LostFocus += new EventHandler(zad_LostFocus);
                zad.FormClosed += new FormClosedEventHandler(zad_FormClosed);
                zad.ShowDialog(Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[dataGridView1.Columns.Count-1].Value));
            }
            if (button1.Text == "Продать")
            {
                pr = new Prodaja(conn, Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[dataGridView1.Columns.Count - 1].Value));
                pr.FormClosed += new FormClosedEventHandler(pr_FormClosed);
                pr.ShowDialog();
            }
            if (button1.Text == "Переместить")
            {
                ps = new Pasting(conn, Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[dataGridView1.Columns.Count - 1].Value));
                ps.FormClosed += new FormClosedEventHandler(ps_FormClosed);
                ps.ShowDialog();
            }
            if (button1.Text == "Удалить")
            {
                del = new delete(conn, Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[dataGridView1.Columns.Count - 1].Value));
                del.FormClosed += new FormClosedEventHandler(del_FormClosed);
                del.ShowDialog();
            }
            
        }
        void del_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = true;
            //comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            button1.Visible = false;
            button2.Visible = false;
            CheckPartiyaEnd();
            CheckPartiyaKon();
            добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        void ps_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = true;
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            button1.Visible = false;
            button2.Visible = false;
            CheckPartiyaEnd();
            CheckPartiyaKon();
            добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        void pr_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            comboBox1.Enabled = true;
            button1.Visible = false;
            CheckPartiyaEnd();
            CheckPartiyaKon();
            добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        void zad_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = true;
            //comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            button1.Visible = false;
            button2.Visible = false;
            CheckPartiyaEnd();
            CheckPartiyaKon();
        }
        void zad_LostFocus(object sender, EventArgs e)
        {
            this.Show();
            button1.Text = "Добавить";
            button1.Visible = true;
            button2.Visible = true;
        }
        private void добавитьНовыйПродуктToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            AddProd adp = new AddProd(conn);
            adp.Show();
            adp.FormClosed += new FormClosedEventHandler(adp_FormClosed);
        }
        void adp_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
        }
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            button1.Visible = false;
            button2.Visible = false;
            comboBox2.Visible = false;
            label3.Visible = false;
            //label3.Text = "";
        }
        private void показатьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = false;
            SqlZapros("select name as Наименование from prodykt","Наименование продукции",false);
        }
        private void показатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = false;
            SqlZapros("select name as Имя, surname as Фамилия, data as [Время записи] from rabotnik where working=1", "Работники",false);
        }
        private void добавитьToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            показатьToolStripMenuItem_Click(dataGridView1, new EventArgs());
            AddRab adr = new AddRab(conn);
            adr.Show();
            adr.FormClosed += new FormClosedEventHandler(adr_FormClosed);
            adr.Invalidated += new InvalidateEventHandler(adr_Invalidated);
        }
        void adr_Invalidated(object sender, InvalidateEventArgs e)
        {
            показатьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        private void удалитьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            показатьToolStripMenuItem_Click(dataGridView1, new EventArgs());
            if (dataGridView1.RowCount != 0)
            {
                if (MessageBox.Show("Выделите записи для удаления и нажмите 'Удалить'\nДля выделения несколький записей держите нажатой кнопку Ctrl", "Удаление", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    button1.Text = "Удалить работника";
                    button1.Visible = true;
                }
            }
        }
        private void добавитьНовоеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem1_Click(dataGridView1, new EventArgs());
            adpr1 = new AddProd1(conn,"Новый продукт");
            adpr1.Show();
            adpr1.FormClosed += new FormClosedEventHandler(adpr1_FormClosed);
            adpr1.Invalidated += new InvalidateEventHandler(adpr1_Invalidated);
        }
        void adpr1_Invalidated(object sender, InvalidateEventArgs e)
        {
            показатьToolStripMenuItem1_Click(dataGridView1, new EventArgs());
        }
        void adpr1_FormClosed(object sender, FormClosedEventArgs e)
        {
            показатьToolStripMenuItem1_Click(dataGridView1, new EventArgs());
        }
        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem1_Click(dataGridView1, new EventArgs());
            if (dataGridView1.RowCount != 0)
            {
                if (MessageBox.Show("Выделите записи для удаления и нажмите 'Удалить'\nДля выделения несколький записей держите нажатой кнопку Ctrl", "Удаление", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    button1.Text = "Удалить продукт";
                    button1.Visible = true;
                }
            }
        }
        private void удалитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem2_Click(dataGridView1, new EventArgs());
            if (dataGridView1.RowCount != 0)
            {
                if (MessageBox.Show("Выделите записи для удаления и нажмите 'Удалить'\nДля выделения несколький записей держите нажатой кнопку Ctrl", "Удаление", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    button1.Text = "Удалить оборудование";
                    button1.Visible = true;
                }
            }
        }
        private void показатьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = false;
            SqlZapros("select name as Наименование from stanok", "Оборудование",false);
        }
        private void добавитьНовоеОборудованиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem2_Click(dataGridView1, new EventArgs());
            adpr1 = new AddProd1(conn,"Новое оборудование");
            adpr1.Show();
            adpr1.FormClosed += new FormClosedEventHandler(adpr1_FormClosed1);
            adpr1.Invalidated += new InvalidateEventHandler(adpr1_Invalidated1);
        }
        void adpr1_Invalidated1(object sender, InvalidateEventArgs e)
        {
            показатьToolStripMenuItem2_Click(dataGridView1, new EventArgs());
        }
        void adpr1_FormClosed1(object sender, FormClosedEventArgs e)
        {
            показатьToolStripMenuItem2_Click(dataGridView1, new EventArgs());
        }
        private void выдатьЗаданиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = true;
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Не выполнено");
            comboBox1.Items.Add("Выполнено");
            comboBox1.SelectedIndex = 0;
        }
        private void выполнитьЗаданиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Не выполнено");
            comboBox1.Items.Add("Выполнено");
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = false;
            //SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата обработки/прихода], stanok.name as Станок, rabotnik.surname as Рабочий, vessklad.zatracheno as [Затрачено времени(час)],state.name as Состояние, vessklad.id from partiya,prodykt,stanok,rabotnik,vessklad,state where vessklad.idstate in (2) and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstanok=stanok.id and vessklad.idrabotnik=rabotnik.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", "Выбор продукта для переработки",true);
            if (MessageBox.Show("Выберите продукт и нажмите 'Выбрать'", "Выбор продукта", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
            button1.Text = "Выбрать";
            button1.Visible = true;
        }
        private void выполнитьЗаданиеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            добавитьToolStripMenuItem_Click(sender, e);
            comboBox1.SelectedIndex = 2;
            comboBox1.Enabled = false;
            //SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата обработки/прихода], stanok.name as Станок, rabotnik.surname as Рабочий, vessklad.zatracheno as [Затрачено времени(час)],state.name as Состояние, vessklad.id from partiya,prodykt,stanok,rabotnik,vessklad,state where vessklad.idstate in (2) and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstanok=stanok.id and vessklad.idrabotnik=rabotnik.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", "Выбор продукта для переработки", true);
            if (MessageBox.Show("Выберите продукт и нажмите 'Переработать'", "Выбор продукта", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
            button1.Text = "Переработать";
            button1.Visible = true;
        }
        private void заполнитьСкладToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            comboBox1.SelectedIndex = 0;
            zap = new Zapolnenie(conn);
            zap.Show();
            zap.FormClosed += new FormClosedEventHandler(zap_FormClosed);
            zap.Invalidated += new InvalidateEventHandler(zap_Invalidated);
        }
        void zap_Invalidated(object sender, InvalidateEventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
        }
        void zap_FormClosed(object sender, FormClosedEventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Выбрать")
            {
                this.Hide();
                zad.Show();
            }
            if (button1.Text == "Выбрать ")
            {
                this.Hide();
                un.Show();
            }
        }
        private void показатьToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = false;
            label2.Text = "Партия";
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Add("Наименование", "Наименование");
            dataGridView1.Columns.Add("Дата начала", "Дата начала");
            dataGridView1.Columns.Add("Дата окончания переработки", "Дата окончания переработки");
            dataGridView1.Columns.Add("Дата окончания", "Дата окончания");
            dataGridView1.Columns.Add("", "");
            dataGridView1.Columns[4].Visible = false;

            conn.Open();
            command.CommandText = "select name,id from partiya where show=1 and name!='Не определено'";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    dataGridView1.Rows.Add(new string[] { r[0].ToString(), "", "", "", r[1].ToString() });
                }
            }
            conn.Close();
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                conn.Open();
                command.CommandText = "select data from sobitiepartii where idnamesobpar=1 and idpartiya="+dataGridView1.Rows[i].Cells[4].Value.ToString();
                dataGridView1.Rows[i].Cells[1].Value = command.ExecuteScalar();
                command.CommandText = "select data from sobitiepartii where idnamesobpar=2 and idpartiya=" + dataGridView1.Rows[i].Cells[4].Value.ToString();
                dataGridView1.Rows[i].Cells[2].Value = command.ExecuteScalar();
                command.CommandText = "select data from sobitiepartii where idnamesobpar=3 and idpartiya=" + dataGridView1.Rows[i].Cells[4].Value.ToString();
                dataGridView1.Rows[i].Cells[3].Value = command.ExecuteScalar();
                conn.Close();
            }
        }
        private void показатьToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = false;
            SqlZapros("select name as Отходы from tipothoda","Вид отходов",false);
        }       
        void adpr1_Invalidated2(object sender, InvalidateEventArgs e)
        {
            показатьToolStripMenuItem4_Click(dataGridView1, new EventArgs());
        }
        void adpr1_FormClosed2(object sender, FormClosedEventArgs e)
        {
            показатьToolStripMenuItem4_Click(dataGridView1, new EventArgs());
        }
        private void удалитьToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem4_Click(dataGridView1, new EventArgs());
            if (dataGridView1.RowCount != 0)
            {
                if (MessageBox.Show("Выделите записи для удаления и нажмите 'Удалить'\nДля выделения несколький записей держите нажатой кнопку Ctrl", "Удаление", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    button1.Text = "Удалить виды отходов";
                    button1.Visible = true;
                }
            }
        }
        private void FillRich(int idsklad,string s)
        {
            string line = "________________________________________________________";
            ArrayList idpererab = new ArrayList();
            ArrayList prih = new ArrayList();
            ArrayList idskladi = new ArrayList();
            try
            {
                conn.Open();
                command.CommandText = "select sobitie.idpererab from sobitie,dvig,balans where idsklad=" + idsklad + " and balans.name='Расход' and dvig.name='Переработка' and sobitie.idbalans=balans.id and sobitie.iddvig=dvig.id";
                r = command.ExecuteReader();
                if (r.HasRows == true)
                {
                    int k = 0;
                    while (r.Read() == true)
                    {                       
                        idpererab.Add((int)r[0]);
                        k++;
                    }
                }
                else
                {
                    //richTextBox1.Visible = false;
                    //dataGridView1.Width = w * 2;
                    //richTextBox1.Text = "";
                    conn.Close();
                    //MessageBox.Show("Продукт этой партии ещё не перерабатывался");
                    return;
                }
                conn.Close();

                for (int i = 0; i < idpererab.Count; i++)
                {
                    //richTextBox1.Text += "\r\n";
                    conn.Open();
                    //command.CommandText = "select partiya.name, prodykt.name, sobitie.ves,  sobitie.data from vessklad,partiya,prodykt,sobitie,balans,dvig where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id and sobitie.idbalans=balans.id and sobitie.iddvig=dvig.id and dvig.name='Переработка' and balans.name='Расход' and sobitie.idpererab="+idpererab[i].ToString();
                    command.CommandText = "select top 1 stanok.name, rabotnik.surname, sobitie.data from sobitie,vessklad,rabotnik,stanok,dvig,balans where vessklad.idrabotnik=rabotnik.id and vessklad.idstanok=stanok.id and sobitie.idsklad=vessklad.id and sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and balans.name='Приход' and dvig.name='Переработка' and sobitie.idpererab=" + idpererab[i].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows == true)
                    {
                        richTextBox1.Text += "\r\n";
                        richTextBox1.Text += s+line;
                        richTextBox1.Text += "\r\n";
                        while (r.Read() == true)
                        {
                            richTextBox1.Text += s + r[2].ToString() + "   Станок: " + r[0].ToString() + "   Рабочий: " + r[1].ToString();
                        }
                    }
                    conn.Close();
                    //string sel = "select prodykt.name as Продукт, sobitie.ves as Вес(кг), stanok.name as Станок, rabonik.surname as Рабочий, sobitie.data as Дата переработки, vessklad.";
                    conn.Open();
                    //command.CommandText = "select top 1 stanok.name, rabotnik.surname from sobitie,vessklad,rabotnik,stanok,dvig,balans where vessklad.idrabotnik=rabotnik.id and vessklad.idstanok=stanok.id and sobitie.idsklad=vessklad.id and sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and balans.name='Приход' and dvig.name='Переработка' and sobitie.idpererab="+idpererab[i].ToString();
                    command.CommandText = "select partiya.name,prodykt.name, sobitie.ves from vessklad,partiya,prodykt,sobitie,balans,dvig where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id and sobitie.idbalans=balans.id and sobitie.iddvig=dvig.id and dvig.name='Переработка' and balans.name='Расход' and sobitie.idpererab=" + idpererab[i].ToString();
                    r = command.ExecuteReader();
                    richTextBox1.Text += "\r\n";
                    richTextBox1.Text += s+"Взято:";
                    if (r.HasRows == true)
                    {
                        while (r.Read() == true)
                        {
                            richTextBox1.Text += "\r\n   ";
                            richTextBox1.Text += s+r[1].ToString() + "\t" + r[2].ToString() + " кг.";
                        }
                    }
                    conn.Close();
                    conn.Open();
                    prih.Clear();
                    idskladi.Clear();
                    command.CommandText = "select prodykt.name, sobitie.ves, sobitie.idsklad from vessklad, prodykt, sobitie,dvig, balans where vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id and sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and balans.name='Приход' and dvig.name='Переработка' and sobitie.idpererab=" + idpererab[i].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows == true)
                    {
                        richTextBox1.Text +="\r\n"+s+ "Получено:";
                        while (r.Read() == true)
                        {
                            //richTextBox1.Text += "\r\n";
                            //richTextBox1.Text +=s+ "   " + r[0].ToString() + "  " + r[1].ToString() + " кг.";
                            prih.Add(s + "   " + r[0].ToString() + "\t" + r[1].ToString() + " кг.");
                            idskladi.Add((int)r[2]);
                        }
                    }
                    conn.Close();

                    for (int k = 0; k < prih.Count; k++)
                    {
                        richTextBox1.Text += "\r\n";
                        richTextBox1.Text += (string)prih[k];
                        //s=s+"  ";
                        FillRich((int)idskladi[k],s+"\t");
                    }

                    conn.Open();
                    command.CommandText = "select tipothoda.name, othodi.ves from tipothoda, othodi where othodi.idtipothoda=tipothoda.id and othodi.idpererab=" + idpererab[i].ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        richTextBox1.Text += "\r\n" + s + "Отходы:";
                        while (r.Read())
                        {
                            richTextBox1.Text += "\r\n";
                            richTextBox1.Text +=s+ "   " + r[0].ToString() + "\t" + r[1].ToString() + " кг.";
                        }
                    }
                    conn.Close();
                    richTextBox1.Text += "\r\n"+s+line;


                }

            }
            catch (System.Exception e)
            {
                richTextBox1.Text = "";
                conn.Close();
                MessageBox.Show(e.Message);
                return;
            }

        }
        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (label2.Text == "Партия")
            {
                conn.Open();
                command.CommandText = "select id from vessklad where idrabotnik is null and idstanok is null and idsost=1 and idpartiya="+dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                int idsklad = (int)command.ExecuteScalar();
                conn.Close();
                
                tree = new Tree(conn,  idsklad);
                tree.ShowDialog();
            }
            if (label2.Text == "Партия1")
            {
                //dataGridView1.Width = w;
                conn.Open();
                command.CommandText = "select id from partiya where name='"+dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()+"'";
                int idpart = (int)command.ExecuteScalar();
                conn.Close();
                conn.Open();
                command.CommandText = "select vessklad.id from vessklad,sost where idpartiya=" + idpart + "and idrabotnik=" + idrabotnik + " and idstanok=" + idstanok + " and sost.name='Новое' and vessklad.idsost=sost.id";
                idskladpartiya = (int)command.ExecuteScalar();
                conn.Close();

                conn.Open();
                command.CommandText = "select nachves, ostatok from vessklad where id="+idskladpartiya;
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {         
                        if ((int)r[0] == (int)r[1])
                        {
                            dataGridView1.Width = ClientRectangle.Width - krai;
                            richTextBox1.Visible = false;                            
                            comboBox2.Visible = false;
                            label3.Visible = false;
                            MessageBox.Show("Выбранная партия еще не перерабатывалась");
                            conn.Close();
                            idskladpartiya = 0;
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Информации о выбранной записи не существует");
                    conn.Close();
                    return;
                }
                conn.Close();
                iter = 0;

                comboBox2.Visible = true;
                label3.Visible = true;
                dataGridView1.Width = ClientRectangle.Width / 2 - krai;
                comboBox2.SelectedIndex = 0;               
                //comboBox2_SelectedIndexChanged(dataGridView1, new EventArgs());
               // FillRich(idsklad, "");
            }
        }
        void FillRich2(int idsklad,string s)
        {
            string line = "________________________________________________________";
            ArrayList idpererab = new ArrayList();
            ArrayList prih = new ArrayList();
            ArrayList idskladi = new ArrayList();
            conn.Open();
            command.CommandText = "select nachves, ostatok from vessklad where id=" + idsklad;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    if ((int)r[0] == (int)r[1])
                    {
                        conn.Close();
                        return;
                    }
                }
            }
            conn.Close();
            iter++;
            if (iter > 1)
            {
                richTextBox1.Text += "\t";
                conn.Open();
                command.CommandText = "select ostatok from vessklad where id="+idsklad;
                richTextBox1.Text+= "Осталось: "+command.ExecuteScalar().ToString()+" кг";
                conn.Close();
            }
            richTextBox1.Text += "\r\n"+s+line;
            conn.Open();
            string sprice = "";
            if (price) sprice = ",sobitie.price";
            command.CommandText = "select sobitie.data,prodykt.name,sobitie.ves "+sprice+" from sobitie,prodykt,vessklad,dvig where sobitie.idsklad=vessklad.id and prodykt.id=vessklad.idprodykt and dvig.id=sobitie.iddvig and dvig.name='Продажа' and sobitie.idsklad="+idsklad;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                richTextBox1.Text +="\r\n"+ s+"Продано:";
                while (r.Read())
                {
                    richTextBox1.Text += "\r\n" + s + "\t" + r[0].ToString() + " " + r[1].ToString() + " " + r[2].ToString() + " кг";
                    if (price) richTextBox1.Text += " Цена: " + r[3].ToString();
                }
            }
            conn.Close();

            conn.Open();
            //idpererab.Clear();
            command.CommandText = "select sobitie.idpererab from sobitie,dvig,balans where idsklad=" + idsklad + " and balans.name='Расход' and dvig.name='Переработка' and sobitie.idbalans=balans.id and sobitie.iddvig=dvig.id";
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                richTextBox1.Text += "\r\n" + s + "Переработано:";
                while (r.Read())
                {
                    idpererab.Add((int)r[0]);
                }
            }
            else { conn.Close(); return; }
            conn.Close();
            for (int i = 0; i < idpererab.Count; i++)
            {
                conn.Open();
                command.CommandText = "select top 1 stanok.name, rabotnik.surname, sobitie.data from sobitie,vessklad,rabotnik,stanok,dvig,balans where vessklad.idrabotnik=rabotnik.id and vessklad.idstanok=stanok.id and sobitie.idsklad=vessklad.id and sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and balans.name='Приход' and dvig.name='Переработка' and sobitie.idpererab=" + idpererab[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                    //richTextBox1.Text += "\r\n";
                {
                    while (r.Read())
                    {
                        richTextBox1.Text += "\r\n" + "    "+s +r[2].ToString()+"  Станок: "+r[0].ToString()+"  Работник: "+r[1].ToString();
                    }
                }
                conn.Close();
                
                conn.Open();
                command.CommandText = "select partiya.name,prodykt.name, sobitie.ves from vessklad,partiya,prodykt,sobitie,balans,dvig where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id and sobitie.idbalans=balans.id and sobitie.iddvig=dvig.id and dvig.name='Переработка' and balans.name='Расход' and sobitie.idpererab=" + idpererab[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    richTextBox1.Text += "\r\n" + s + "     "+"Взято:";
                    while (r.Read())
                    {
                        richTextBox1.Text += "\r\n" + s + "\t" + r[1].ToString() + "\t" + r[2].ToString() + " кг.";
                    }
                }
                conn.Close();

                prih.Clear();
                idskladi.Clear();
                conn.Open();
                command.CommandText = "select prodykt.name, sobitie.ves, sobitie.idsklad from vessklad, prodykt, sobitie,dvig, balans where vessklad.idprodykt=prodykt.id and sobitie.idsklad=vessklad.id and sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and balans.name='Приход' and dvig.name='Переработка' and sobitie.idpererab=" + idpererab[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    richTextBox1.Text += "\r\n" + s + "     "+"Получено:";
                    while (r.Read() == true)
                    {
                        prih.Add("\r\n" + s + "\t" + r[0].ToString() + "\t" + r[1].ToString() + " кг.");
                        idskladi.Add((int)r[2]);
                    }
                }
                conn.Close();

                for (int k = 0; k < prih.Count; k++)
                {
                    richTextBox1.Text += (string)prih[k];
                    FillRich2((int)idskladi[k], s + "\t");
                }

                conn.Open();
                command.CommandText = "select tipothoda.name, othodi.ves from tipothoda, othodi where othodi.idtipothoda=tipothoda.id and othodi.idpererab=" + idpererab[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    richTextBox1.Text += "\r\n" + s + "     " + "Отходы:";
                    while (r.Read())
                    {

                        richTextBox1.Text += "\r\n" + s + "\t" + r[0].ToString() + "\t\t" + r[1].ToString() + " кг.";
                    }
                }
                conn.Close();
                richTextBox1.Text += "\r\n" + s+line;

            }
        }
        private void FillTree(int idskladpartiya)
        {

        }
        private void хранитьЦенуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (хранитьЦенуToolStripMenuItem.Checked == true) { price = true; MessageBox.Show("Включена функция хранения и отображения цены"); }
            if (хранитьЦенуToolStripMenuItem.Checked == false) {price = false; MessageBox.Show("Отключена функция хранения и отображения цены");}
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //FileInfo fi;
            //fi = new FileInfo("opt.txt");
            //StreamWriter sw = fi.CreateText();
            //if (price == true) sw.Write("price:on");
            //if (price == false) sw.Write("price:off");
            //sw.Close();
        }
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (richTextBox1.Visible)
            {
                dataGridView1.Width = ClientRectangle.Width / 2 - krai;
                richTextBox1.Location = new System.Drawing.Point(dataGridView1.Right + krai, dataGridView1.Location.Y);
                richTextBox1.Width = dataGridView1.Width;
                richTextBox1.Height = dataGridView1.Height;
            }
        }
        private void продажаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Готовая продукция");
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = false;
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            //добавитьToolStripMenuItem_Click(sender, new EventArgs());
            button1.Text = "Продать";
            button1.Visible = true;
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            //conn.Open();
            //command.CommandText = "select id from partiya where name='" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() + "'";
            //int idpart = (int)command.ExecuteScalar();
            //conn.Close();
            //conn.Open();
            //command.CommandText = "select vessklad.id from vessklad,sost where idpartiya=" + idpart + "and idrabotnik=" + idrabotnik + " and idstanok=" + idstanok + " and sost.name='Новое' and vessklad.idsost=sost.id";
            //idskladpartiya = (int)command.ExecuteScalar();
            //conn.Close();

            if ((string)comboBox2.SelectedItem == "Дерево")
            {

            }
            if ((string)comboBox2.SelectedItem == "Подробно текст")
            {
                richTextBox1.Location = new System.Drawing.Point(dataGridView1.Right + krai, dataGridView1.Location.Y);
                richTextBox1.Width = dataGridView1.Width;
                richTextBox1.Height = dataGridView1.Height;
                richTextBox1.Visible = true;
                richTextBox1.Text = "";
                iter = 0;
                conn.Open();
                string s = "";
                string s1 = "";
                string s2 = "";
                if (price == true) { s = ",sobitie.price"; s1 = ",sobitie"; s2 = "and sobitie.idsklad=vessklad.id"; }
                command.CommandText = "select partiya.name, prodykt.name, vessklad.nachves, vessklad.ostatok " + s + " from vessklad,partiya,prodykt" + s1 + " where partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt " + s2 + " and vessklad.id=" + idskladpartiya;
                r = command.ExecuteReader();

                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        richTextBox1.Text += r[0].ToString() + "  " + r[1].ToString() + "   Нач.вес: " + r[2].ToString() + "кг.   Осталось: " + r[3].ToString() + "кг.";
                        if (price == true) { richTextBox1.Text += "    Цена: " + r[4].ToString(); }
                        break;
                    }
                }
                conn.Close();
                richTextBox1.Text += "\r\n";


                FillRich2(idskladpartiya, "");
            }
            if ((string)comboBox2.SelectedItem == "Кратко текст")
            {

            }
        }
        private void показатьToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = false;
            SqlZapros("select name as[Наименование] from kAgent","Контрагенты", false);
        }
        private void добавитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem4_Click(dataGridView1, new EventArgs());
            adpr1 = new AddProd1(conn, "Вид отходов");
            adpr1.Show();
            adpr1.FormClosed += new FormClosedEventHandler(adpr1_FormClosed2);
            adpr1.Invalidated += new InvalidateEventHandler(adpr1_Invalidated2);
        }
        private void добавитьНовоеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem5_Click(dataGridView1, new EventArgs ());
            addka= new AddProd1(conn,"Контрагент");
            addka.Show();
            addka.FormClosed += new FormClosedEventHandler(addka_FormClosed);
            addka.Invalidated += new InvalidateEventHandler(addka_Invalidated);
        }
        void addka_Invalidated(object sender, InvalidateEventArgs e)
        {
            показатьToolStripMenuItem5_Click(dataGridView1, new EventArgs());
        }
        void addka_FormClosed(object sender, FormClosedEventArgs e)
        {
            показатьToolStripMenuItem5_Click(dataGridView1, new EventArgs());
        }
        private void удалитьToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            показатьToolStripMenuItem5_Click(dataGridView1, new EventArgs());
            if (dataGridView1.RowCount != 0)
            {
                if (MessageBox.Show("Выделите записи для удаления и нажмите 'Удалить'\nДля выделения несколький записей держите нажатой кнопку Ctrl", "Удаление", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    button1.Text = "Удалить контрагент";
                    button1.Visible = true;
                }
            }
        }
        private void переработатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripMenuItem3_Click(sender, e);
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            добавитьToolStripMenuItem_Click(sender, e);//добавить - значит показать по ходу
            comboBox1.SelectedIndex = 2; //выбор склада
            comboBox1.Enabled = false;
            //SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата обработки/прихода], stanok.name as Станок, rabotnik.surname as Рабочий, vessklad.zatracheno as [Затрачено времени(час)],state.name as Состояние, vessklad.id from partiya,prodykt,stanok,rabotnik,vessklad,state where vessklad.idstate in (2) and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstanok=stanok.id and vessklad.idrabotnik=rabotnik.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", "Выбор продукта для переработки", true);
            if (MessageBox.Show("Выберите продукт и нажмите 'Переработать'", "Выбор продукта", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
            button1.Text = "Переработать";
            button1.Visible = true;
        }
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Не выполнено");
            comboBox1.Items.Add("Выполнено");
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = false;
            //SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата обработки/прихода], stanok.name as Станок, rabotnik.surname as Рабочий, vessklad.zatracheno as [Затрачено времени(час)],state.name as Состояние, vessklad.id from partiya,prodykt,stanok,rabotnik,vessklad,state where vessklad.idstate in (2) and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstanok=stanok.id and vessklad.idrabotnik=rabotnik.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", "Выбор продукта для переработки",true);
            if (MessageBox.Show("Выберите продукт и нажмите 'Выбрать'", "Выбор продукта", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
            button1.Text = "Выбрать";
            button1.Visible = true;
        }
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = true;
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Не выполнено");
            comboBox1.Items.Add("Выполнено");
            comboBox1.SelectedIndex = 0;
        }
        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {

        }
        private void перемещениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            добавитьToolStripMenuItem_Click(sender, new EventArgs());
            button1.Text = "Переместить";
            button1.Visible = true;
        }
        private void заполнитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rab = new Rabochie(conn);
            rab.ShowDialog();

            //f2 = new Form2(conn);
            //f2.ShowDialog();
            label4.Text = "";
        }
        private void label2_TextChanged(object sender, EventArgs e)
        {
            if (label2.Text == "Весь склад" || label2.Text == "Склад производства" || label2.Text == "Готовая продукция")
            {
                conn.Close();
                //SqlConnection sk = conn;
                conn.Open();
                //sk.Open();
                //SqlCommand komm = new SqlCommand();
                //komm.Connection = sk;
                //command = new SqlCommand();
                command.CommandText = "select idsklad,ves,data,idproizv,id from sobitie where iddvig=2 and idbalans=2 and price is null";
                r=command.ExecuteReader();
                //int x = (int)komm.ExecuteScalar();
                if (r.HasRows)
                {
                    label4.Text = "Не забудьте выдать зарплату!";
                }
                conn.Close();
            }
            else
            {
                label4.Text = "";
            }
        }
        private void готоваяПродукцияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sostSklad = new SostSklad(conn,"Готовая продукция");
            sostSklad.ShowDialog();
        }
        private void открытьФормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckPartiyaEnd();
            sostSklad = new SostSklad(conn);
            sostSklad.ShowDialog();
        }
        void un_FormClosed(object sender, FormClosedEventArgs e)
        {
            button2.Visible = false;
            button1.Visible = false;
            добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        void un_LostFocus(object sender, EventArgs e)
        {
            this.Show();
            button1.Text = "Добавить ";
            button1.Visible = true;
            button2.Visible = true;
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (perekl != 0)
            {
                if (comboBox3.Text == "Все" && comboBox1.Text != "Весь склад")
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], vessklad.id from partiya,prodykt,sost,vessklad,state where state.name ='" + (comboBox1.Text) + "' and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idsost=sost.id and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    perekl1 = 0;
                    textBox1.Text = "";
                    perekl1 = 1;
                    return;
                }
                if (comboBox3.Text == "Все" && comboBox1.Text == "Весь склад")
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(переработки)], state.name as Положение  ,vessklad.id from partiya,prodykt,vessklad, state, sost where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstate=state.id and vessklad.idsost=sost.id  and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    perekl1 = 0;
                    textBox1.Text = "";
                    perekl1 = 1;
                    return;
                }
                if (comboBox1.Text == "Весь склад")
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(переработки)], state.name as Положение  ,vessklad.id from partiya,prodykt,vessklad, state, sost where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.idstate=state.id and vessklad.idsost=sost.id and partiya.name='" + comboBox3.Text + "' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    perekl1 = 0;
                    textBox1.Text = "";
                    perekl1 = 1;
                    return;
                }
                else
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], vessklad.id from partiya,prodykt,sost,vessklad,state where state.name ='" + (comboBox1.Text) + "' and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idsost=sost.id and vessklad.idstate=state.id and partiya.name='" + comboBox3.Text + "' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    perekl1 = 0;
                    textBox1.Text = "";
                    perekl1 = 1;
                }
                
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (perekl != 0&&perekl1!=0)
            {
                if (textBox1.Text.Length >= 3)
                {
                    if (comboBox3.Text == "Все"&&comboBox1.Text=="Весь склад")
                    {
                        SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)],state.name as Положение, vessklad.id from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and prodykt.name like '%" + textBox1.Text + "%'  and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                        return;
                    }
                    if (comboBox3.Text == "Все" && comboBox1.Text != "Весь склад")
                    {
                        SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)] ,vessklad.id from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and prodykt.name like '%" + textBox1.Text + "%'  and vessklad.idstate=state.id and state.name='"+comboBox1.Text+"' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                        return;
                    }
                    if (comboBox3.Text != "Все" && comboBox1.Text == "Весь склад")
                    {
                        SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], state.name as Положение ,vessklad.id from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and prodykt.name like '%" + textBox1.Text + "%'  and vessklad.idstate=state.id and partiya.name='" + comboBox3.Text + "' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                        return;
                    }
                    else
                    {
                        SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], vessklad.id from partiya,prodykt,sost,vessklad,state where state.name ='" + (comboBox1.Text) + "' and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idsost=sost.id and vessklad.idstate=state.id and partiya.name='" + comboBox3.Text + "' and prodykt.name like '%" + textBox1.Text + "%' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    }
                }
            }
        }
        private void button3_MouseHover(object sender, EventArgs e)
        {
            toolTip1.Show("Сбросить результаты поиска", button3);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                textBox1.Text = "";
                if (comboBox3.Text == "Все" && comboBox1.Text == "Весь склад")
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)],state.name as Положение, vessklad.id from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idstate=state.id and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    return;
                }
                if (comboBox3.Text == "Все" && comboBox1.Text != "Весь склад")
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)] , state.name as Положение, vessklad.id from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idstate=state.id and state.name='"+comboBox1.Text+"' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    return;
                }
                if (comboBox3.Text != "Все" && comboBox1.Text == "Весь склад")
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)], state.name as Положение ,vessklad.id from partiya,prodykt,vessklad,state where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idstate=state.id and partiya.name='" + comboBox3.Text + "' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                    return;
                }
                else
                {
                    SqlZapros("select partiya.name as Партия, prodykt.name as Продукт, vessklad.ostatok as [Осталось(кг)], vessklad.data as [Дата поступления(обработки)],state.name as Положение, vessklad.id from partiya,prodykt,vessklad,state where state.name ='" + (comboBox1.Text) + "' and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id  and vessklad.idstate=state.id and partiya.name='" + comboBox3.Text + "' and ostatok!=0 order by partiya.name asc", comboBox1.Text, true);
                }
            }
        }
        private void базаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBox1.Enabled = false;
            comboBox3.Enabled = false;
            button3.Enabled = false;
        }
        private void показатьСделанноеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckZadanie chZad = new CheckZadanie(conn);
            chZad.ShowDialog();
        }
        private void показатьToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            Zarplata zarp = new Zarplata(conn);
            zarp.ShowDialog();
        }
        private void продатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            продажаToolStripMenuItem_Click(sender, e);
        }
        private void показатьToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            checkProdaja chPr = new checkProdaja(conn,"Показать проданное");
            chPr.ShowDialog();
        }
        private void button4_Click(object sender, EventArgs e)
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

            Word.Paragraph oPara2;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara2.Range.Text = label2.Text;
            oPara2.Range.Font.Bold = 1;
            oPara2.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara2.Range.InsertParagraphAfter();

           

            int r = 0;
            int c = 0;
            c = dataGridView1.ColumnCount-1;
            if(label2.Text=="Весь склад"||label2.Text == "Склад производства"||label2.Text == "Готовая продукция")c = dataGridView1.ColumnCount-2;
            r = dataGridView1.RowCount+1;
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, r, c, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;

            for (int i = 1; i <= c; i++)
            {
                oTable.Cell(1, i).Range.Text = dataGridView1.Columns[i].Name;
            }

            for (int i = 2; i <= r; i++)
            {
                for (int j = 1; j <= c; j++)
                {
                    //if (listView1.Items[i - 2].SubItems[j - 1].Text == "Итого:") oTable.Rows[i].Range.Font.Shadow = 5;
                    oTable.Cell(i, j).Range.Text = dataGridView1.Rows[i-2].Cells[j].Value.ToString();
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
        private void удалитьToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());            
            button1.Text = "Удалить";
            button1.Visible = true;
        }
        private bool UnionSame(int idpartiya,int idprodykt,int idstate, int idsost)
        {
            //проверка и объединение одинаковых позиций
            string date = DateTime.Now.ToLongDateString();
            conn.Open();
            command.CommandText = "select id,ostatok,sostav from vessklad where idpartiya=" + idpartiya + " and idprodykt=" + idprodykt + "  and idstate=" + idstate + " and idsost=" + idsost + " and ostatok!=0";
            r = command.ExecuteReader();
            int idsostav = -1;
            ArrayList same = new ArrayList();
            if (r.HasRows == true)
            {
                while (r.Read() == true)
                {
                    try
                    {
                        if ((bool)r[2]) { idsostav = (int)r[0]; continue; }
                    }
                    catch (System.Exception)
                    {

                    }
                    decimal[] ss = new decimal[2];
                    ss[0] = Convert.ToDecimal(r[0].ToString().Replace('.', ','));
                    ss[1] = Convert.ToDecimal(r[1].ToString().Replace('.', ','));
                    same.Add(ss);
                }
            }
            conn.Close();
            bool checker = false;
            if (same.Count == 1 && idsostav == -1) checker = false;
            if (same.Count > 1) checker = true;
            if (same.Count >= 1 && idsostav != -1) checker = true;

            if (checker)
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
                    sumar = (decimal[])same[0];
                    conn.Open();
                    command.CommandText = "select data from vessklad where id=" + sumar[0];
                    string datastart = (string)command.ExecuteScalar();
                    conn.Close();

                    conn.Open();
                    command.CommandText = "insert into vessklad(idpartiya,idprodykt,nachves,ostatok,idrabotnik,idstanok,data,idstate,idsost,recordtime,sostav) values (" + idpartiya + "," + idprodykt + "," + sum.ToString().Replace(',', '.') + "," + sum.ToString().Replace(',', '.') + "," + idrabotnik + "," + idstanok + ",'" + datastart + "'," + idstate + "," + idsost + ",'" + DateTime.Now.ToString() + "','1')";
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
                    command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + idnew + "," + sum.ToString().Replace(',', '.') + ",2,5,1," + idpereun + ",'" + date + "','" + DateTime.Now.ToString() + "')";
                    command.ExecuteNonQuery();
                    command.CommandText = "select max(id) from sobitie";
                    idsobun = (int)command.ExecuteScalar();
                    conn.Close();
                }
                else
                {
                    idnew = idsostav;
                    conn.Open();
                    command.CommandText = "select id, idproizv from sobitie where iddvigfrom=2 and iddvig=5 and idbalans=1 and idsklad=" + idnew.ToString();
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
                for (int u = 0; u < same.Count; u++)
                {
                    sumar = (decimal[])same[u];
                    conn.Open();
                    command.CommandText = "update vessklad set ostatok=0 where id=" + sumar[0];//остаток=0
                    command.ExecuteNonQuery();
                    command.CommandText = "update vessklad set sostav=0 where id=" + sumar[0];//оно составное
                    command.ExecuteNonQuery();//вставка в событие что оно объединилось
                    command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + sumar[0] + "," + sumar[1].ToString().Replace(',', '.') + ",2,5,2," + idpereun + ",'" + date + "','" + DateTime.Now.ToString() + "')";
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
                    command.CommandText = "update proizv set idprodykt" + idprodMax + "=" + sumar[0] + " where id=" + idpereun;
                    command.ExecuteNonQuery();
                    conn.Close();
                    ostpohogih += sumar[1];//суммируем остатки  остальных
                }
                //делаем все для объединяющего(изменяем остаток,вставляем в событие что к нему добавилось)
                conn.Open();
                decimal sumobch = sum + ostpohogih;
                command.CommandText = "update vessklad set ostatok=" + sumobch.ToString().Replace(',', '.') + " where id=" + idnew;//остаток=0
                command.ExecuteNonQuery();
                command.CommandText = "insert into sobitie(idsklad,ves,iddvigfrom,iddvig,idbalans,idproizv,data,recordtime) values (" + idnew + "," + ostpohogih.ToString().Replace(',', '.') + ",2,5,1," + idpereun + ",'" + date + "','" + DateTime.Now.ToString() + "')";
                conn.Close();
                return true;
            }
            return false;
            //конец нового
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Весь склад" || comboBox1.Text == "Готовая продукция" || comboBox1.Text == "Склад производства")
            {
                comboBox1.Enabled=false;
                int idi=5;
                int prt = 0;
                int prd = 0;
                int stt = 0;
                int sott = 0;
                bool ret = true;
                if(comboBox1.Text=="Весь склад")idi=6;        
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    conn.Open();
                    command.CommandText = "select idpartiya,idprodykt,idstate,idsost from vessklad where id=" + dataGridView1.Rows[i].Cells[idi].Value.ToString();
                    r = command.ExecuteReader();
                    if (r.Read())
                    {
                        prt = (int)r[0]; prd = (int)r[1]; stt = (int)r[2]; sott = (int)r[3];
                    }
                    conn.Close();
                    if (UnionSame(prt, prd, stt, sott))
                    {
                        break;
                    }
                }
                добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
            }           
            comboBox1.Enabled = true;
        }
        private void настройкиПартийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setp = new SetPartiya(conn);
            setp.ShowDialog();
        }
        private void парольНаУдалениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pwd = new Password();
            pwd.ShowDialog();
        }
        private void заполнениеОстатковToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1_SelectedIndexChanged(dataGridView1, new EventArgs());
            comboBox1.SelectedIndex = 0;
            zap = new Zapolnenie(conn);
            zap.Show();
            zap.FormClosed += new FormClosedEventHandler(zap_FormClosed);
            zap.Invalidated += new InvalidateEventHandler(zap_Invalidated);
        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void удалитьПартиюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            delp = new deletepartiya(conn);
            delp.FormClosed += new FormClosedEventHandler(delp_FormClosed);
            delp.ShowDialog();
        }
        void delp_FormClosed(object sender, FormClosedEventArgs e)
        {
            добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        private void войтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            starting = new Starting();
            starting.ShowDialog();
            if (fwd) { Start(); this.Text = "Склад -"+user+"-"; }
            else
            {
                MessageBox.Show("Авторизуйтесь! Программа не сможет работать!");
            }
        }
        private void изменитьЛогинToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fwd)
            {
                chlog = new changelogpwd();
                chlog.ShowDialog();
                Text = "Склад -" + user + "-";
            }
        }
        private void изменитьПарольToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fwd)
            {
                chpwd = new changepwd();
                chpwd.ShowDialog();
            }
        }
        private void очиститьСоздатьНовуюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            reset = new ResetSklad(conn);
            reset.FormClosed += new FormClosedEventHandler(reset_FormClosed);
            reset.ShowDialog();
        }
        void reset_FormClosed(object sender, FormClosedEventArgs e)
        {
            добавитьToolStripMenuItem_Click(dataGridView1, new EventArgs());
        }
        private void объединитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            un = new Union(conn);
            un.LostFocus += new EventHandler(un_LostFocus);
            un.FormClosed += new FormClosedEventHandler(un_FormClosed);
            un.ShowDialog();
        }
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //MessageBox.Show("tt");
            //if (comboBox1.Text == "Весь склад")
            //{
            //    if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "Готовая продукция")
            //    {
            //        dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Готовая продукция";
            //        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //        {
            //            //if (dataGridView1.Rows[i].Cells[0].Value.ToString()=="93") MessageBox.Show(dataGridView1.Rows[i].Cells[2].Value.ToString());
            //            if (dataGridView1.Rows[i].Cells[5].Value.ToString() == "Готовая продукция" && dataGridView1.Rows[i].Cells[2].Value.ToString() == dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString() && dataGridView1.Rows[e.RowIndex].Visible == true)
            //            {
            //                //MessageBox.Show(dataGridView1.Rows[i].Cells[2].Value.ToString());
            //                dataGridView1.Rows[i].Cells[3].Value = Convert.ToDecimal(dataGridView1.Rows[i].Cells[3].Value) + Convert.ToDecimal(dataGridView1.Rows[e.RowIndex].Cells[3].Value);
            //                if (Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].Value) > Convert.ToDateTime(dataGridView1.Rows[e.RowIndex].Cells[4].Value))
            //                {
            //                    dataGridView1.Rows[i].Cells[4].Value = dataGridView1.Rows[e.RowIndex].Cells[4].Value;
            //                }
            //                dataGridView1.Rows[e.RowIndex - 1].Visible = false;
            //                dataGridView1.Rows[e.RowIndex - 1].Cells[0].Value = Convert.ToInt16(dataGridView1.Rows[e.RowIndex - 1].Cells[0].Value) - 1;
            //                //dataGridView1.Rows.Clear();
            //                //dataGridView1.Rows.SharedRow(e.RowIndex+1);
            //                //e.RowIndex--;
            //                //continue;
            //            }

            //        }
            //    }

            //}
        }
        private void показатьToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            checkProdaja chPr = new checkProdaja(conn, "Показать объединенное");
            chPr.ShowDialog();
        }

        private void историяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkProdaja chPr = new checkProdaja(conn, "История");
            chPr.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Otchet otchet = new Otchet(conn, comboBox1.Text);
            otchet.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Start a new workbook in Excel.
            object m_objOpt = System.Reflection.Missing.Value;
            Excel.Application m_objExcel = new Excel.Application();
            Excel.Workbooks m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            Excel._Workbook m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            ///////
            int r = 0;
            int c = 0;
            c = dataGridView1.ColumnCount - 1;
            r = dataGridView1.RowCount + 1;
            if (label2.Text == "Весь склад" || label2.Text == "Склад производства" || label2.Text == "Готовая продукция") c = dataGridView1.ColumnCount - 2;
            
           // Add data to cells in the first worksheet in the new workbook.
            Excel.Sheets m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            Excel._Worksheet m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));

            object[] objHeaders = new object[c];
            for (int i = 1; i <= c; i++)
            {
                objHeaders[i - 1] = dataGridView1.Columns[i].Name;
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
                    rrt=dataGridView1.Rows[r2-1].Cells[h+1].Value.ToString();
                    objData[r2-1, h] = rrt;
                }
            }
            m_objRange = m_objSheet.get_Range("A2", m_objOpt);
            m_objRange = m_objRange.get_Resize(r, c);
            m_objRange.Value = objData;
            m_objRange.NumberFormat = "@";
            m_objExcel.Visible = true;
        }

        private void показатьПриходТовараToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkProdaja chPrihod = new checkProdaja(conn, "Показать приход товара");
            chPrihod.ShowDialog();
        }


        
    }
}