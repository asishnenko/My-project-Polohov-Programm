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
    public partial class SostSklad : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        ArrayList arr;
        public SostSklad(SqlConnection conn)
        {
            this.conn = conn;
            command.Connection = conn;
            InitializeComponent();
            richTextBox1.Visible = false;
            listView1.Visible = true;
            //��������� ��������� ��������� ������
            ViborVarianta.Items.Add("������ ������������� �������������� ������");
            ViborVarianta.Items.Add("����� ������� ��������� �� "+DateTime.Today.ToShortDateString());
            ViborVarianta.Items.Add("����� ������������ �� ������� " + DateTime.Today.ToShortDateString());
            ViborVarianta.Items.Add("���� ����� �� ������� " + DateTime.Today.ToShortDateString());
            ViborVarianta.Items.Add("�������� �����");
            ViborVarianta.Items.Add("��������� ������ �� ����");
        }
        public SostSklad(SqlConnection conn,string windowName)
        {
            this.conn = conn;
            command.Connection = conn;            
            InitializeComponent();            
            this.Text = windowName;
            if (windowName == "���������� �������")
            {

            }
            if (windowName == "qq")
            {
                arr = new ArrayList();
                richTextBox1.Text = "����� ������� ��������� �� ��������� �� " + DateTime.Today.ToShortDateString() + "\r\n";
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,vessklad.data from vessklad, partiya,prodykt where vessklad.idstate=1 and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        ArrayList temp = new ArrayList();
                        for (int i = 0; i < 4; i++)
                        {
                            temp.Add(r[i]);
                        }
                        arr.Add(temp);
                    }
                }
                conn.Close();
                string partiya = "";
                string text = "";
                string text1 = "";
                for (int i = 0; i < arr.Count; i++)
                {
                    ArrayList temp = new ArrayList();
                    temp = (ArrayList)arr[i];
                    if (partiya == "" || partiya != (string)temp[0])
                    {
                        text = "������: " + (string)temp[0] + "\r\n\t";
                    }
                    else text = "\t";
                    int ves = (int)temp[2];
                    text1 = (string)temp[1] + ":\t\t" + ves.ToString() + "��.\t\t����: " + (string)temp[3] + "\r\n";
                    richTextBox1.Text += text + text1;
                    partiya = (string)temp[0];
                }
            }
        }
        public void VizibleSebest(bool x)
        {
            if (x == true)
            {
                //Info.Visible = true;
                label2.Visible = true;
                label2.Text = "������";
                label3.Visible = true;
                label3.Text = "������� �� ������. ������(���.)";
                comboBox2.Visible = true;
                textBox1.Visible = true;
                button1.Visible = true;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                label5.Visible = false;
                comboBox1.Visible = false;
            }
            if (x == false)
            {
                //Info.Visible = false;
                label2.Visible = false;
                label3.Visible = false;
                comboBox2.Visible = false;
                textBox1.Visible = false;
                button1.Visible = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                label5.Visible = false;
                comboBox1.Visible = false;
            }
        }
        public void VizibleMoving(bool x)
        {
            if (x == true)
            {
                //Info.Visible = true;
                label2.Visible = true;
                label2.Text = "�����";
                label3.Visible = true;
                label3.Text = "��������";
                comboBox2.Visible = true;
                textBox1.Visible = false;
                button1.Visible = true;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                label5.Visible = true;
                comboBox1.Visible = true;
            }
            if (x == false)
            {
                //Info.Visible = false;
                label2.Visible = false;
                label3.Visible = false;
                comboBox2.Visible = false;
                textBox1.Visible = false;
                button1.Visible = false;
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            VizibleSebest(false);
            richTextBox1.Visible=false;
            label4.Text = "";
            if (ViborVarianta.Text == "��������� ������ �� ����")
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                VizibleMoving(true);
                dateTimePicker2.Visible = false;
                listView1.Columns.Add("������",80);
                listView1.Columns.Add("�������", 120);
                listView1.Columns.Add("�����", 120);
                listView1.Columns.Add("�������", 120);
                listView1.Columns.Add("id", 0);
                label3.Text = "����";                                
                comboBox2.Items.Clear();
                comboBox2.Items.Add("���");
                comboBox2.Items.Add("����� ������������");
                comboBox2.Items.Add("������� ���������");
                comboBox1.Items.Clear();
                comboBox1.Items.Add("���");
                conn.Open();
                command.CommandText = "select name from partiya where show=1 and name!='�� ����������'";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox1.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();
                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
            }
            if (ViborVarianta.Text == "����� ������������ �� ������� " + DateTime.Today.ToShortDateString())
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                ListViewItem lvi;           
                listView1.Columns.Add("������",100);
                listView1.Columns.Add("�������", 100);
                listView1.Columns.Add("���(��)", 100);
                label4.Text = ViborVarianta.Text;
                decimal sum = 0;
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,vessklad.data from vessklad, partiya,prodykt where vessklad.idstate=2 and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        sum += Convert.ToDecimal(r[2].ToString());
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString() });
                        listView1.Items.Add(lvi);
                    }
                }
                conn.Close();
                lvi = new ListViewItem(new string[] { "�����", "", sum.ToString() });
                lvi.BackColor = Color.FromArgb(240, 240, 240);
                listView1.Items.Add(lvi);
            }
            if (ViborVarianta.Text == "����� ������� ��������� �� " + DateTime.Today.ToShortDateString())
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                ListViewItem lvi;
                listView1.Columns.Add("������", 100);
                listView1.Columns.Add("�������", 100);
                listView1.Columns.Add("���(��)", 100);
                label4.Text = ViborVarianta.Text;
                decimal sum = 0;
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,vessklad.data from vessklad, partiya,prodykt where vessklad.idstate=1 and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        sum += Convert.ToDecimal(r[2].ToString());
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString() });
                        listView1.Items.Add(lvi);
                    }
                }
                conn.Close();
                lvi = new ListViewItem(new string[] { "�����", "", sum.ToString() });
                lvi.BackColor = Color.FromArgb(240, 240, 240);
                listView1.Items.Add(lvi);
                
            }
            if (ViborVarianta.Text == "���� ����� �� ������� " + DateTime.Today.ToShortDateString())
            {
                listView1.Columns.Clear();
                listView1.Items.Clear();
                ListViewItem lvi;
                listView1.Columns.Add("������", 100);
                listView1.Columns.Add("�������", 100);
                listView1.Columns.Add("���(��)", 100);
                listView1.Columns.Add("�����", 120);
                //listView1.Columns.Add("id", 120);
                //listView1.Columns.Add("sostav", 120);
                label4.Text = ViborVarianta.Text;
                decimal sum = 0;
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.ostatok,state.name,vessklad.id,vessklad.sostav from vessklad, partiya,prodykt,state where state.id=vessklad.idstate and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and ostatok!=0 order by partiya.name asc";
                r = command.ExecuteReader();
                if (r.HasRows == true)
                {
                    while (r.Read() == true)
                    {
                        sum += Convert.ToDecimal(r[2].ToString());
                        lvi = new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString(), r[5].ToString() });
                        listView1.Items.Add(lvi);
                    }
                }
                conn.Close();
                lvi = new ListViewItem(new string[] { "�����", "", sum.ToString(),"" ,"",""});                
                lvi.BackColor = Color.FromArgb(240,240,240);
                listView1.Items.Add(lvi);

            }
            if (ViborVarianta.Text == "������ ������������� �������������� ������")
            {
                VizibleSebest(true);
                comboBox2.Items.Clear();
                listView1.Columns.Clear();
                listView1.Items.Clear();
                //��������� ��������� ������
                conn.Open();
                command.CommandText = "select name from partiya where konetc=1 and sbor=0 and show=1 and partiya.name!='�� ����������' and partiya.ostatki!=1";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox2.Items.Add((string)r[0]);
                    }
                }
                conn.Close();
                //Info.Text = "�������� ������ ��� ����������� ������!";
                MessageBox.Show("�������� ������ ��� ����������� ������!","�����!",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
            if (ViborVarianta.Text == "�������� �����")
            {
                VizibleMoving(true);
                listView1.Items.Clear();
                listView1.Columns.Clear();
                listView1.Columns.Add("������", 80);
                listView1.Columns.Add("�������",120);
                listView1.Columns.Add("�����", 120);
                listView1.Columns.Add("����.�������", 80);
                listView1.Columns.Add("������", 80);
                listView1.Columns.Add("������", 80);
                listView1.Columns.Add("�����.�������", 80);
                listView1.Columns.Add("id", 0);
                //listView1.Columns.Add("id", 80);
                comboBox2.Items.Clear();
                comboBox2.Items.Add("���");
                comboBox2.Items.Add("����� ������������");
                comboBox2.Items.Add("������� ���������");
                comboBox1.Items.Clear();
                comboBox1.Items.Add("���");
                conn.Open();
                command.CommandText = "select name from partiya where show=1 and name!='�� ����������'";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        comboBox1.Items.Add(r[0].ToString());
                    }
                }
                conn.Close();
                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
            }
        }
        string partiya;
        string nomenkl;
        decimal ves=0;
        decimal tcena=0;
        decimal symma=0;
        int idsiriya=0;
        ArrayList vihod;
        private void button1_Click(object sender, EventArgs e)
        {
            //�������� ����������!!!
            richTextBox1.Text = "";
            if (ViborVarianta.Text == "��������� ������ �� ����")
            {
                string data1 = dateTimePicker1.Value.ToShortDateString();
                //string data2 = dateTimePicker2.Value.ToShortDateString();
                listView1.Items.Clear();
                //if (dateTimePicker1.Value >= dateTimePicker2.Value) { MessageBox.Show("�������������� ���!"); return; }
                string part = "";
                string state = "";
                if (comboBox2.Text != "���") state = " and state.name='" + comboBox2.Text + "' ";
                if (comboBox1.Text != "���") part = " and partiya.name='" + comboBox1.Text + "' ";
                conn.Open();
                command.CommandText = "select vessklad.id,vessklad.data,vessklad.sostav from vessklad,partiya,state where vessklad.idpartiya=partiya.id and vessklad.sostav is null and vessklad.idstate!=6 and vessklad.idstate=state.id " + part + state;
                //ArrayList a_id = new ArrayList();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        if (Convert.ToDateTime(r[1].ToString()) > Convert.ToDateTime(data1)) continue;//�� ������� ���� ���� ��������>����2
                        //a_id.Add((int)r[0]);
                        listView1.Items.Add(new ListViewItem(new string[] { "", "", "", "",  r[0].ToString() }));


                    }
                }
                conn.Close();

                conn.Open();
                command.CommandText = "select vessklad.id,vessklad.data,vessklad.sostav from vessklad,partiya,state where vessklad.idpartiya=partiya.id and vessklad.sostav=1 and vessklad.idstate!=6 and vessklad.idstate=state.id " + part + state;
                //a_id = new ArrayList();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        if (Convert.ToDateTime(r[1].ToString()) > Convert.ToDateTime(data1)) continue;//�� ������� ���� ���� ��������>����2
                        listView1.Items.Add(new ListViewItem(new string[] { "", "", "", "",  r[0].ToString() }));

                    }
                }
                conn.Close();

                ////�������� �� �������������� ����(���� ������<���� ������ � ���� ��������>���� �����)
                //conn.Open();
                //for (int i = 0; i < listView1.Items.Count; i++)
                //{
                //    command.CommandText = "select sobitie.data from sobitie where sobitie.idsklad=" + listView1.Items[i].SubItems[4].Text + " and iddvig=4";
                //    string sss = (string)command.ExecuteScalar();
                //    if (sss != null)
                //    {
                //        if (Convert.ToDateTime(sss) < Convert.ToDateTime(data1)) { listView1.Items[i].Remove(); i--; }
                //    }

                //}
                //conn.Close();


                //��������� ������� ��������� �������, ���������� ���� �����
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "select partiya.name,prodykt.name,state.name from partiya,prodykt,vessklad,state where vessklad.idstate=state.id and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + listView1.Items[i].SubItems[4].Text;
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            listView1.Items[i].SubItems[0].Text = (string)r[0];
                            listView1.Items[i].SubItems[1].Text = (string)r[1];
                            listView1.Items[i].SubItems[2].Text = (string)r[2];
                        }
                    conn.Close();
                    decimal ves1 = 0;
                    conn.Open();
                    command.CommandText = "select sobitie.ves,sobitie.data,balans.name,dvig.name from sobitie,balans,dvig where sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and dvig.name!='����������' and sobitie.idsklad=" + listView1.Items[i].SubItems[4].Text;
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            DateTime dt = Convert.ToDateTime(r[1].ToString());
                            if (dt < Convert.ToDateTime(data1))
                            {
                                if ((string)r[2] == "������") ves1 += Convert.ToDecimal(r[0].ToString());
                                if ((string)r[2] == "������") ves1 -= Convert.ToDecimal(r[0].ToString());
                            }
                        }
                    //r.Close();
                    listView1.Items[i].SubItems[3].Text = ves1.ToString();
                    conn.Close();
                   
                }
                //������� �������
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    if (listView1.Items[i].SubItems[3].Text == "0" || listView1.Items[i].SubItems[3].Text == "0,0")
                    {
                        listView1.Items[i].Remove();
                        i--;
                    }
                }

            }
            if (ViborVarianta.Text == "�������� �����")
            {
                //string comand = "";
                string data1 = dateTimePicker1.Value.ToShortDateString();
                string data2 = dateTimePicker2.Value.ToShortDateString();
                listView1.Items.Clear();
                if (dateTimePicker1.Value >= dateTimePicker2.Value) { MessageBox.Show("�������������� ���!"); return; }
                string part = "";
                string state = "";
                if (comboBox2.Text != "���") state = " and state.name='"+comboBox2.Text+"' ";
                if (comboBox1.Text != "���") part = " and partiya.name='"+comboBox1.Text+"' ";
                conn.Open();
                command.CommandText = "select vessklad.id,vessklad.data,vessklad.sostav from vessklad,partiya,state where vessklad.idpartiya=partiya.id and vessklad.sostav is null and vessklad.idstate!=6 and vessklad.idstate=state.id " + part + state;
                //ArrayList a_id = new ArrayList();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        if (Convert.ToDateTime(r[1].ToString()) > Convert.ToDateTime(data2)) continue;//�� ������� ���� ���� ��������>����2
                        //a_id.Add((int)r[0]);
                        listView1.Items.Add(new ListViewItem(new string[] { "", "", "", "", "", "", "", r[0].ToString() }));

                        
                    }
                }
                conn.Close();

                conn.Open();
                command.CommandText = "select vessklad.id,vessklad.data,vessklad.sostav from vessklad,partiya,state where vessklad.idpartiya=partiya.id and vessklad.sostav=1 and vessklad.idstate!=6 and vessklad.idstate=state.id " + part + state;
                //a_id = new ArrayList();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    while (r.Read())
                    {
                        if (Convert.ToDateTime(r[1].ToString()) > Convert.ToDateTime(data2)) continue;//�� ������� ���� ���� ��������>����2
                        listView1.Items.Add(new ListViewItem(new string[] { "", "", "", "", "", "", "", r[0].ToString() }));

                    }
                }
                conn.Close();

                //�������� �� �������������� ����(���� ������<���� ������ � ���� ��������>���� �����)
                conn.Open();
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    command.CommandText = "select sobitie.data from sobitie where sobitie.idsklad="+listView1.Items[i].SubItems[7].Text+" and iddvig=4";
                    string sss = (string)command.ExecuteScalar();
                    if (sss != null)
                    {
                        if (Convert.ToDateTime(sss) < Convert.ToDateTime(data1)) { listView1.Items[i].Remove(); i--; }
                    }
                    
                }
                conn.Close();

                //��������� ������� ��������� �������, ���������� ���� �����
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    conn.Open();
                    command.CommandText = "select partiya.name,prodykt.name,state.name from partiya,prodykt,vessklad,state where vessklad.idstate=state.id and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + listView1.Items[i].SubItems[7].Text;
                    r = command.ExecuteReader();
                    if (r.HasRows)
                        while (r.Read())
                        {
                            listView1.Items[i].SubItems[0].Text = (string)r[0];
                            listView1.Items[i].SubItems[1].Text = (string)r[1];
                            listView1.Items[i].SubItems[2].Text = (string)r[2];
                        }
                    conn.Close();
                    decimal ves1 = 0;
                    decimal ves2 = 0;
                    decimal ves3 = 0;
                    decimal ves4 = 0;
                    conn.Open();                
                    command.CommandText = "select sobitie.ves,sobitie.data,balans.name,dvig.name from sobitie,balans,dvig where sobitie.iddvig=dvig.id and sobitie.idbalans=balans.id and dvig.name!='����������' and sobitie.idsklad=" + listView1.Items[i].SubItems[7].Text;
                    r = command.ExecuteReader();
                    if(r.HasRows)
                        while (r.Read())
                        {
                            DateTime dt=Convert.ToDateTime(r[1].ToString());
                            if (dt < Convert.ToDateTime(data1))
                            {
                                if ((string)r[2] == "������") ves1 += Convert.ToDecimal(r[0].ToString());
                                if ((string)r[2] == "������") ves1 -= Convert.ToDecimal(r[0].ToString());
                            }
                            //if (dt > Convert.ToDateTime(data2))
                            //{
                            //    if ((string)r[2] == "������") ves4 += Convert.ToDecimal(r[0].ToString());
                            //    if ((string)r[2] == "������") ves4 -= Convert.ToDecimal(r[0].ToString());
                            //}
                            if (dt >= Convert.ToDateTime(data1) && dt <= Convert.ToDateTime(data2))
                            {
                                if ((string)r[2] == "������") ves2 += Convert.ToDecimal(r[0].ToString());
                                if ((string)r[2] == "������") ves3 += Convert.ToDecimal(r[0].ToString());
                            }

                        }
                    //r.Close();
                    listView1.Items[i].SubItems[3].Text = ves1.ToString();
                    listView1.Items[i].SubItems[4].Text = ves2.ToString();
                    listView1.Items[i].SubItems[5].Text = ves3.ToString();
                    ves4=Convert.ToDecimal(listView1.Items[i].SubItems[3].Text) + Convert.ToDecimal(listView1.Items[i].SubItems[4].Text) - Convert.ToDecimal(listView1.Items[i].SubItems[5].Text);
                    listView1.Items[i].SubItems[6].Text = ves4.ToString();
                    conn.Close();

                }
         
            }
            if (ViborVarianta.Text == "������ ������������� �������������� ������")
            {
                if (comboBox2.Text != "" && textBox1.Text != "")
                {
                    decimal comyn=0;
                    //�������� �� �����
                    try
                    {
                        comyn = Convert.ToDecimal(textBox1.Text);
                    }
                    catch (System.Exception)
                    {
                        MessageBox.Show("� ���� ������� ������� ������ �����!");
                        return;
                    }
                    listView1.Visible = true;
                    richTextBox1.Visible = false;
                    //������������� �������
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("",80);
                    listView1.Columns.Add("������������",120);
                    listView1.Columns.Add("�������(��)", 80);
                    listView1.Columns.Add("�������(��)", 80);
                    listView1.Columns.Add("�����(��)", 80);
                    listView1.Columns.Add("����(��/��)",100);
                    listView1.Columns.Add("�����(��)",100);
                    //listView1.Columns.Add("���������",80);
                    listView1.Columns.Add("id", 0);
                    partiya = comboBox2.Text;
                    conn.Open();
                    command.CommandText = "select vessklad.id,prodykt.name,vessklad.nachves,sobitie.price from vessklad,sobitie,partiya,prodykt where sobitie.iddvig=1 and sobitie.iddvigfrom is null and vessklad.id=sobitie.idsklad and partiya.id=vessklad.idpartiya and prodykt.id=vessklad.idprodykt and partiya.name='"+partiya+"'";
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        r.Read();
                        idsiriya = (int)r[0];
                        nomenkl = (string)r[1];
                        ves = Convert.ToDecimal(r[2].ToString());
                        tcena = Convert.ToDecimal( r[3].ToString());
                    }
                    symma = ves * tcena;
                    conn.Close();
                    conn.Open();
                    ArrayList temp;
                    vihod = new ArrayList();//�������� � ������ ���������
                    command.CommandText = "select prodykt.name,vessklad.ostatok,vessklad.id from vessklad,prodykt,state,partiya where partiya.id=vessklad.idpartiya and partiya.name='" + partiya + "' and state.id=vessklad.idstate and state.name in ('������� ���������') and prodykt.id=vessklad.idprodykt and vessklad.ostatok!=0 and vessklad.id!=" + idsiriya.ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {
                            temp = new ArrayList();
                            temp.Add(r[0]);
                            temp.Add(r[1]);
                            temp.Add(r[2]);
                            vihod.Add(temp);
                        }
                    }
                    conn.Close();
                    //string line = "\r\n_____________________________________________________________________\r\n";
                    string line = "\r\n";
                    string shapka1 = "������:"+line+"�������.\t\t���(��)\t����\t�����(��)\t����.";
                    string shapka2 = "����� ������� ���������:" + line + "�������.\t\t���(��)";
                    string line1 = nomenkl+"\t\t"+ves.ToString()+"\t"+tcena.ToString()+"\t"+symma.ToString()+"\t\t"+partiya;
                    //string tab = "\t\t\t\t\t\t\t";
                    listView1.Items.Add(new ListViewItem(new string[] { "������:", "", "", "", "", "", "","" }));
                    listView1.Items.Add(new ListViewItem( new string[] {"",partiya+" "+nomenkl,"","",ves.ToString(), tcena.ToString(),symma.ToString(),""}));
                    listView1.Items.Add(new ListViewItem(new string[] { "����� �.�.:", "", "", "", "", "","","" }));
                    richTextBox1.Text = shapka1 + line + line1;
                    richTextBox1.Text += line+line+shapka2+line;
                    decimal summagot = 0;
                    for (int i = 0; i < vihod.Count; i++)
                    {
                        temp = new ArrayList();
                        temp = (ArrayList)vihod[i];
                        decimal ves1 = Convert.ToDecimal( temp[1].ToString());
                        summagot += ves1;
                        string nom=(string)temp[0];
                        string lines="";
                        int id = 0;
                        id = (int)temp[2];
                        if(nom.Length>15) lines = nom+"\t"+ves1.ToString()+"\t";
                        else  lines = nom+"\t\t"+ves1.ToString()+"\t";
                        lines += line;                        
                        richTextBox1.Text += lines;
                        listView1.Items.Add(new ListViewItem(new string[] { "", nom, ves1.ToString(),"","", "", "", id.ToString()}));
                    }
                    //��������� �� ���������� � ������
                    conn.Open();
                    command.CommandText = "select prodykt.name,vessklad.ostatok,vessklad.id from vessklad,prodykt,state,partiya,sobitie where sobitie.idsklad=vessklad.id and sobitie.iddvig=3 and partiya.id=vessklad.idpartiya and partiya.name='" + partiya + "' and state.id=vessklad.idstate and state.name in ('������� ���������') and prodykt.id=vessklad.idprodykt and vessklad.ostatok=0 and vessklad.id!=" + idsiriya.ToString();
                    r = command.ExecuteReader();
                    if (r.HasRows)
                    {
                        while (r.Read())
                        {
                            listView1.Items.Add(new ListViewItem(new string[] { "", r[0].ToString(), r[1].ToString(), "", "", "", "", r[2].ToString() }));
                        }
                    }
                    conn.Close();
                    //������� ������� ������� �������
                    summagot = 0;
                    int ipereb=0;
                    conn.Open();
                    decimal prodano = 0;
                    while (ipereb < listView1.Items.Count)
                    {
                        if (listView1.Items[ipereb].SubItems[7].Text != "")
                        {
                            
                            command.CommandText = "select sum(sobitie.ves) from sobitie where idsklad="+listView1.Items[ipereb].SubItems[7].Text+" and iddvig=3 and idbalans=2";
                            try //�� ������ ���� ������ �� �������
                            {
                                prodano = Convert.ToDecimal(command.ExecuteScalar().ToString());
                            }
                            catch (System.Exception)
                            {
                                prodano = 0;
                            }
                            decimal itogo = Convert.ToDecimal(listView1.Items[ipereb].SubItems[2].Text) + prodano;
                            listView1.Items[ipereb].SubItems[3].Text = prodano.ToString();
                            listView1.Items[ipereb].SubItems[4].Text = itogo.ToString();
                            summagot += itogo;
                        }
                        ipereb++;
                    }
                    conn.Close();
                    //������
                    decimal othod = ves - summagot;
                    richTextBox1.Text += "�����\t\t\t"+othod.ToString()+line;
                    listView1.Items.Add(new ListViewItem(new string[] { "������:", "", "", "", othod.ToString(), "", "", "" }));

                    //��������� ������������                    
                    conn.Open();
                    decimal zatratitr = 0;
                    command.CommandText = "select sum(sobitie.price) from sobitie, partiya, vessklad where partiya.id=vessklad.idpartiya and sobitie.idsklad=vessklad.id and partiya.name='"+partiya+"' and iddvigfrom=2 and iddvig=2 and idbalans=2 and price is not null";
                    zatratitr = Convert.ToDecimal(command.ExecuteScalar().ToString());
                    zatratitr = Math.Round(zatratitr, 2);
                    conn.Close();
                    listView1.Items.Add(new ListViewItem(new string[] { "�����������:", "", "", "", "", "", zatratitr.ToString(), "" }));
                    listView1.Items.Add(new ListViewItem(new string[] { "������������:", "", "", "", "", "", comyn.ToString(), "" }));
                    richTextBox1.Text += line+"������� ����� �� �����������(��): " + zatratitr.ToString();
                    decimal sebest = 0;
                    //���� ������ �� ����������������
                    if (zatratitr == 0)
                    {
                        //int sebest = 0;
                        richTextBox1.Text += line + "�������������(��/��): " + sebest.ToString();
                        listView1.Items.Add(new ListViewItem(new string[] { "�������������:", "", "", "", "", sebest.ToString(), "", "" }));
                        return;
                    }
                    try
                    {
                        sebest = (symma + zatratitr + comyn) / summagot;
                        sebest = Math.Round(sebest, 2);
                        richTextBox1.Text += line + "�������������(��/��): " + sebest.ToString();
                        listView1.Items.Add(new ListViewItem(new string[] { "�������������:", "", "", "", "", sebest.ToString(), "", "" }));
                    }
                    catch (System.Exception)
                    {
                        listView1.Items.Clear();
                        MessageBox.Show("�������� ����������� ������!");
                    }
                }
                else MessageBox.Show("��������� ��� ����!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
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
            oPara1.Range.Text = label4.Text;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();


            int r = 0;
            int c = 0;
            c = listView1.Columns.Count;
            r = listView1.Items.Count+1;
            if (ViborVarianta.Text == "������ ������������� �������������� ������"||ViborVarianta.Text == "�������� �����")
            {
                c--;
                oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            }            
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
                    if (listView1.Items[i - 2].SubItems[j - 1].Text == "�����:") oTable.Rows[i].Range.Font.Shadow = 5;
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

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            string part = "";
            string state = "";
            if (comboBox2.Text != "���") state = " and state.name='" + comboBox2.Text + "' ";
            if (comboBox1.Text != "���") part = " and partiya.name='" + comboBox1.Text + "' ";
            conn.Open();
            command.CommandText = "select vessklad.id,vessklad.data,vessklad.sostav,state.name from vessklad,partiya,state where vessklad.idpartiya=partiya.id and vessklad.sostav is null and vessklad.idstate!=6 and vessklad.idstate=state.id " + part + state;
            ArrayList a_id = new ArrayList();
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    //if (Convert.ToDateTime(r[1].ToString()) > dateTimePicker2.Value) continue;//�� ������� ���� ���� ��������>����2
                    a_id.Add((int)r[0]);
                    listView1.Items.Add(new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(),r[3].ToString(),""}));

                }
            }
            conn.Close();
            conn.Open();
            command.CommandText = "select vessklad.id,vessklad.data,vessklad.sostav,state.name from vessklad,partiya,state where vessklad.idpartiya=partiya.id and vessklad.sostav=1 and vessklad.idstate!=6 and vessklad.idstate=state.id " + part + state;
            //ArrayList a_id = new ArrayList();
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    //if (Convert.ToDateTime(r[1].ToString()) > dateTimePicker2.Value) continue;//�� ������� ���� ���� ��������>����2
                    a_id.Add((int)r[0]);
                    listView1.Items.Add(new ListViewItem(new string[] { r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(),"" }));

                }
            }
            conn.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (Convert.ToDateTime(listView1.Items[i].SubItems[1].Text) > dateTimePicker2.Value) { listView1.Items[i].Remove(); i--; continue; }
                if (listView1.Items[i].SubItems[4].Text != "")
                {
                    if (Convert.ToDateTime(listView1.Items[i].SubItems[4].Text) < dateTimePicker1.Value) { listView1.Items[i].Remove(); i--; }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            conn.Open();
            for (int i = 0; i < listView1.Items.Count; i++)
            {

                command.CommandText = "select sobitie.data from sobitie where sobitie.idsklad=" + listView1.Items[i].SubItems[0].Text + " and iddvig=4";
                string sss = "";
                sss = (string)command.ExecuteScalar();
                if (sss == null) sss = "";
                listView1.Items[i].SubItems[4].Text = sss;
            }
            conn.Close();
        }
    }
}