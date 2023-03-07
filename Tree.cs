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
    public partial class Tree : Form
    {
        SqlConnection conn;
        SqlCommand command = new SqlCommand();
        SqlDataReader r;
        ArrayList idlist;
        int idsklad;
        int it = 0;
        public Tree(SqlConnection conn, int id)
        {
            this.conn = conn;
            command.Connection = conn;
            idsklad = id;
            InitializeComponent();
            TreeNode root = new TreeNode();
            treeView1.Nodes.Add(root);            
            Research(id, true, root);            
        }
        public void Research( int id,bool first,TreeNode root)
        {
            it++;
            bool key = first;
            int ido = id;
            //TreeNode root = node;
            ArrayList l1 = new ArrayList();
            ArrayList l2 = new ArrayList();
            ArrayList l3 = new ArrayList();
            string s1 = "";
            if (key)
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.nachves,vessklad.ostatok,vessklad.data from partiya,prodykt,vessklad where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + ido;
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    r.Read();
                    s1 = r[0].ToString() + " " + r[1].ToString() + ". Нач.вес(кг):" + r[2].ToString() + ". Остаток(кг):" + r[3].ToString() + ". Дата поступления:" + r[4].ToString();
                    l1.Add(s1);
                    root.Text = s1;
                    //MessageBox.Show(s1);
                }
                conn.Close();
            }

            conn.Open();
            command.CommandText = "select data,ves,idproizv from sobitie where iddvigfrom=2 and iddvig=2 and idsklad=" + ido;
            r = command.ExecuteReader();
            if (r.HasRows)
            {
                while (r.Read())
                {
                    l2.Add(r[2].ToString());
                    s1 = "-ПЕРЕРАБОТКА-Дата:" + r[0].ToString() + " Взято(кг):" + r[1].ToString();
                    l2.Add(s1);
                    root.Nodes.Add(new TreeNode(s1));
                }
            }
            conn.Close();

            //decimal[] nves = new decimal[l2.Count / 2];
            //decimal[] ost = new decimal[l2.Count / 2];
            ArrayList prodykt = new ArrayList();
            string partiya = "";
            for (int i = 0; i < l2.Count; i = i + 2)
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.nachves,vessklad.ostatok,vessklad.id,stanok.name, rabotnik.surname from partiya,prodykt,vessklad,sobitie,stanok,rabotnik where stanok.id=vessklad.idstanok and rabotnik.id=vessklad.idrabotnik and vessklad.id=sobitie.idsklad and vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and sobitie.idproizv=" + l2[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    bool ch = true;
                    int k = 0;
                    while (r.Read())
                    {
                        //nves[k] += Convert.ToDecimal(r[2].ToString().Replace('.', ','));
                        //ost[k] += Convert.ToDecimal(r[3].ToString().Replace('.', ','));
                        if (ch)
                        {
                            root.Nodes[i / 2].Text += " Станок:" + r[5].ToString() + " Работник:" + r[6].ToString();
                            partiya = r[0].ToString();
                            prodykt.Add(r[1].ToString());
                        }
                        l2[i + 1] += " Станок:" + r[5].ToString() + " Работник:" + r[6].ToString();
                        l3.Add(r[4].ToString());
                        s1 = r[0].ToString() + " " + r[1].ToString() + "| Вес(кг):" + r[2].ToString();
                        root.Nodes[i / 2].Nodes.Add(new TreeNode(s1));
                        ch = false;
                        k++;
                    }
                }
                conn.Close();
            }
            ArrayList l4 = new ArrayList();//id склада составное
            ArrayList l5 = new ArrayList();//id склада не составного
            for (int i = 0; i < l3.Count; i++)
            {
                conn.Open();
                command.CommandText = "select distinct idsklad from sobitie where iddvigfrom=2 and iddvig=5 and idbalans=1 and idproizv=(select idproizv from sobitie where iddvigfrom=2 and iddvig=5 and idbalans=2 and idsklad=" + l3[i].ToString() + ")";
                r = command.ExecuteReader();
                if (r.HasRows)
                {
            
                    while (r.Read())
                    {
                        if (!l4.Contains(r[0].ToString()))
                        {
                            l4.Add(r[0].ToString());
                        }
                    }
                }
                else
                {
                    if (!l5.Contains(l3[i].ToString()))
                    {
                        l5.Add(l3[i].ToString());
                    }
                }
                conn.Close();
            }
            ArrayList sves = new ArrayList();
            ArrayList ves=new ArrayList();
            for (int i = 0; i < root.Nodes.Count; i++)
            {
                for (int k = 0; k < root.Nodes[i].Nodes.Count; k++)
                {
                    string t=root.Nodes[i].Nodes[k].Text.Substring(0,root.Nodes[i].Nodes[k].Text.IndexOf('|'));
                    if (!sves.Contains(t)) sves.Add(t);
                    for (int tt = 0; tt < sves.Count; tt++)
                    {
                        if ((string)sves[tt] == t)
                        {
                            decimal yy=0;
                            try
                            {
                                yy = (decimal)ves[tt];
                            }
                            catch (System.Exception)
                            {
                                string tre = root.Nodes[i].Nodes[k].Text.Substring(root.Nodes[i].Nodes[k].Text.IndexOf(':')+1);
                                ves.Add(Convert.ToDecimal(tre));
                                continue;
                            }                            
                            string ch = root.Nodes[i].Nodes[k].Text.Substring(root.Nodes[i].Nodes[k].Text.IndexOf(':')+1);
                            yy += Convert.ToDecimal(ch.Replace('.', ','));
                            ves[tt] = yy;                            
                        }
                    }
                }
            }

            for (int i = 0; i < l4.Count; i++)
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.nachves,vessklad.ostatok,vessklad.data from partiya,prodykt,vessklad where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + l4[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    r.Read();
                    s1 = r[0].ToString() + " " + r[1].ToString() + ". Нач.вес(кг):" + ves[i].ToString() + ". Остаток(кг):" + r[3].ToString() + ". Дата поступления:" + r[4].ToString() + " /" + l4[i].ToString();
                    l1.Add(s1);
                    root.Nodes.Add(s1);                                        
                }
                conn.Close();
            }
            for (int i = 0; i < l5.Count; i++)
            {
                conn.Open();
                command.CommandText = "select partiya.name, prodykt.name, vessklad.nachves,vessklad.ostatok,vessklad.data from partiya,prodykt,vessklad where vessklad.idpartiya=partiya.id and vessklad.idprodykt=prodykt.id and vessklad.id=" + l5[i].ToString();
                r = command.ExecuteReader();
                if (r.HasRows)
                {
                    r.Read();
                    s1 = r[0].ToString() + " " + r[1].ToString() + ". Нач.вес(кг):" + r[2].ToString() + ". Остаток(кг):" + r[3].ToString() + ". Дата поступления:" + r[4].ToString()+" /"+l5[i].ToString();
                    l1.Add(s1);
                    TreeNode trr = new TreeNode(s1);                    
                    
                    root.Nodes.Add(trr);                                 
                }
                conn.Close();
            }
            //for (int i = 0; i < l5.Count; i++)
            //{
            //    l4.Add((int)l5[i]);
            //}
            //l5.Clear();
            //return l4;            
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
           
        }

        private void treeView1_NodeMouseHover(object sender, TreeNodeMouseHoverEventArgs e)
        {
            if (e.Node.Nodes.Count == 0 && e.Node.Text.Contains("/"))
            {
                TreeNode root = e.Node;
                //int startind=0;
                //for (int i = 0; i < root.Nodes.Count; i++)
                //{
                //    if (root.Nodes[i].Nodes.Count == 0) { startind = i; break; }
                //}
                //idlist=Research(
                int id = Convert.ToInt32(root.Text.Substring(root.Text.IndexOf('/') + 1));
                root.Text = root.Text.Remove(root.Text.IndexOf('/'));
                Research(id, false, root);
                //root.Expand();
            }
        }
    }
}