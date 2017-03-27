using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
//using Inventor;
using System.Collections;

#region forDll
public class FormTreeView : Form
{
    private TreeView treeView1;
    private Button btn;
    //private Button button1;
    private Dictionary<string, string> dic;
    private Dictionary<string, string> fsDic;
    private Dictionary<string, string> NuM;
    private List<KeyValuePair<int, int>> sortDic;
    List<string> val;
    bool check;
    string sb = "";
    int Offs = 64;

    public FormTreeView(ref List<string> lst)
    {
        treeView1 = new TreeView();
        btn = new Button();
        dic = new Dictionary<string, string>();
        fsDic = new Dictionary<string, string>();
        sortDic = new List<KeyValuePair<int, int>>();
        val = InvAddIn.PropBtn.m_Prop.sort_val;
        NuM = new Dictionary<string, string>();
        check = InvAddIn.PropBtn.m_Prop.part.Checked;

        this.SuspendLayout();

        // Initialize treeView1.
        InitializeComponent();
        treeView1.AllowDrop = true;
        treeView1.Dock = DockStyle.Fill;

        // Initialize Button

        // Add nodes to treeView1.
        TreeNode node;
        node = treeView1.Nodes.Add("Сборки");


        addElem(lst, node, 0);
        //for (int x = 0; x < 3; ++x)
        //{
        //    // Add a root node to treeView1.
        //    node = treeView1.Nodes.Add(String.Format("Node{0}", x * 4));
        //    for (int y = 1; y < 4; ++y)
        //    {
        //        // Add a child node to the previously added node.
        //        node = node.Nodes.Add(String.Format("Node{0}", x * 4 + y));
        //    }
        //}

        // Add event handlers for the required drag events.
        treeView1.ItemDrag += new ItemDragEventHandler(treeView1_ItemDrag);
        treeView1.DragEnter += new DragEventHandler(treeView1_DragEnter);
        treeView1.DragOver += new DragEventHandler(treeView1_DragOver);
        treeView1.DragDrop += new DragEventHandler(treeView1_DragDrop);

        // Initialize the form.
        Rectangle bns = Screen.PrimaryScreen.Bounds;
        bns.Y = 24;
        bns.Width = bns.Width / 4;
        bns.Height = bns.Height;
        btn.Location = new System.Drawing.Point(bns.Width - btn.Width - 1, 0);
        this.Bounds = bns;
        this.Controls.Add(treeView1);
        //treeView1.TreeViewNodeSorter = new nodeSort();
        nSort(treeView1.Nodes[0].Nodes, 0);
        treeView1.ExpandAll();

        this.ResumeLayout(false);
    }

    private void nSort(TreeNodeCollection nodes, int num)
    {
        List<string> val = InvAddIn.PropBtn.m_Prop.sort_val;
        for (int i = 0; i < nodes.Count; i++)
        {
            if (nodes[i].Nodes.Count != 0)
            {
                nSort(nodes[i].Nodes, num++);
                if (nodes.Count > 1 && num > 0)
                {
                    sortD(nodes);
                    i += nodes.Count - 1;
                }
            }
            else if (nodes.Count > 0)
            {
                sortD(nodes);
                i += nodes.Count - 1;
            }
        }
    }

    private void sortD(TreeNodeCollection nodes)
    {
        sortDic = addDic(nodes);
        sortDic.Sort(
            delegate(KeyValuePair<int, int> firstPair,
            KeyValuePair<int, int> nextPair)
            {
                return firstPair.Key >= nextPair.Key ? 1 : -1;
            }
            );
        int i = 0;
        TreeNode[] tmpnodes = new TreeNode[nodes.Count];
        //nodes.Add("End");
        foreach (KeyValuePair<int, int> kvp in sortDic)
        {
            //TreeNode tmpnode = nodes[kvp.Value];
            //tmpnodes.Add(tmpnode);
            tmpnodes[i] = nodes[kvp.Value];
            //nodes.Remove(nodes[kvp.Value]);
            //nodes.Add(nodes[kvp.Value]);
            i++;
            //nodes.RemoveAt(kvp.Value);
        }
        nodes.Clear();
        i = nodes.Count;
        //for (int j = 0; j < i / 2; j++)
        //{
        //    nodes.RemoveAt(0);
        //}
        nodes.AddRange(tmpnodes);
        //tmpnodes.Clear();
    }

    private List<KeyValuePair<int, int>> addDic(TreeNodeCollection nodes)
    {
        sortDic.Clear();
        for (int i = 0; i < nodes.Count; i++)
        {
            int ind1 = val.FindIndex(delegate(string str)
            {
                return nodes[i].Text.ToUpper().IndexOf(str.ToUpper()) != -1;
            });
            ind1 *= 2;
            if (nodes[i].Text.IndexOf("01") != -1) ind1++;
            if (nodes[i].Text.IndexOf("02") != -1) ind1 = +2;
            if (nodes[i].Text.IndexOf(".iam:") != -1) ind1 = 100 + val.Count;
            if (nodes[i].Text.IndexOf(".ipt") != -1) ind1 = +200;
            if (nodes[i].Text.IndexOf(".ipt:") != -1) ind1 = +250;
            if (ind1 == -1) ind1 = 1000 + val.Count;
            sortDic.Add(new KeyValuePair<int, int>(ind1, i));
        }
        return sortDic;
    }

    public void InsertionSort(int[] array)
    {
        for (int i = 1; i < array.Length; i++)
        {
            int cur = array[i];
            int j = i;
            while (j > 0 && cur < array[j - 1])
            {
                array[j] = array[j - 1];
                j--;
            }
            array[j] = cur;
        }
    }

    private void button1_Click(object sender, EventArgs e)
    {
        DataGridView dgv = InvAddIn.PropBtn.m_Prop.GetDGV();
        if (check == false)
        {
            dec(treeView1.Nodes, "", 0);
            updateDGV(dgv);
        }
        else
        {
            numb(treeView1.Nodes, "", 0);
            itemN(dgv);
        }
        this.Hide();
    }

    private void updateDGV(DataGridView dgv)
    {
        try
        {
            int num;
            DataGridViewColumn col = dgv.Columns["DecNumber"];
            if (col == null)
                dgv.Columns.Add("DecNumber", "Переменная часть децимального номера");
            col = dgv.Columns["Type"];
            if (col == null)
                dgv.Columns.Add("Type", "Постоянная часть децимального номера");
            col = dgv.Columns["Part Number"];
            if (col == null)
                dgv.Columns.Add("Part Number", "Обозначение");
            col = dgv.Columns["forSort"];
            if (col == null)
            {
                num = dgv.Columns.Add("forSort", "Сортировка");
                dgv.Columns[num].Visible = false;
            }
        }
        catch (Exception)
        {

        }
        for (int i = 0; i < dgv.Rows.Count - 1; i++)
        {
            string name = dgv[0, i].Value.ToString().Trim();
            if (dgv[dgv.Columns["DecNumber"].Index, i].Value.ToString() == "")
            {
                dgv[dgv.Columns["DecNumber"].Index, i].Value = dic[name];
                dgv[dgv.Columns["Part Number"].Index, i].Value = "=<Type>.<DecNumber>";
            }
            try
            { dgv[dgv.Columns["forSort"].Index, i].Value = fsDic[name]; }
            catch (Exception) { dgv[dgv.Columns["forSort"].Index, i].Value = fsDic[name + ":"]; };
        }
        dgv.Sort(dgv.Columns["forSort"], System.ComponentModel.ListSortDirection.Ascending);
    }

    private void itemN(DataGridView dgv)
    {
        int num;
        try
        {
            DataGridViewColumn col = dgv.Columns["forSort"];
            if (col == null)
            {
                num = dgv.Columns.Add("forSort", "Сортировка");
                dgv.Columns[num].Visible = false;
            }
        }
        catch (Exception) { }
        for (int i = 0; i < dgv.RowCount - 1; i++)
        {
            string name = dgv[0, i].Value.ToString().Trim();
            dgv[dgv.Columns["forSort"].Index, i].Value = fsDic[name];
            dgv[0, i].Value = NuM[name];
        }
        dgv.Sort(dgv.Columns["forSort"], System.ComponentModel.ListSortDirection.Ascending);
    }

    private void dec(TreeNodeCollection node, string number, int num)
    {
        int i = 0, offset = 0;

        foreach (TreeNode n1 in node)
        {
            if (n1.Text.IndexOf(".iam") != -1) i++;
        }
        foreach (TreeNode n in node)
        {
            if (n.Nodes.Count != 0)
            {
                if (n.Text != "Сборки" && num == 0 && n.Text.IndexOf(".iam") != -1)
                {
                    Offs++;
                    sb = ((char)Offs).ToString();
                }
                //if (n.Text.IndexOf(".iam") != -1)
                dec(n.Nodes, number + (n.Index + 1).ToString(), i);
                if (n.Text != "Сборки")
                {
                    if (n.Text.IndexOf(".iam") != -1)
                    {
                        string f = (number + (n.Index + 1)).ToString().Substring(2);
                        while (f.Length != 5)
                        {
                            f += "0";
                        }
                        //if (f.StartsWith("1")) f = "0" + f.Substring(1);
                        dic.Add(n.Text, f.Insert(2, "."));
                        fsDic.Add(n.Text, sb + f + "A");
                    }
                    else
                    {
                        string f = number.Substring(2);
                        if (f.Length < 3)
                        {
                            while (f.Length != 3)
                            {
                                f += "0";
                            }
                        }
                        //if (f.StartsWith("1")) f = "0" + f.Substring(1);
                        f += String.Format("{0:00}", (n.Index + 1 - i));
                        dic.Add(n.Text, f.Insert(2, "."));

                        fsDic.Add(n.Text, sb + f + "B");
                    }
                }
            }
            else
            {
                if (n.Text.IndexOf(".iam") != -1)
                {
                    string f = (number + (n.Index + 1)).ToString().Substring(2);
                    while (f.Length != 5)
                    {
                        f += "0";
                    }
                    dic.Add(n.Text, f.Insert(2, "."));
                    fsDic.Add(n.Text, sb + f + "A");
                }
                else
                {
                    string f = number.Substring(2);
                    if (f.Length < 3)
                    {
                        while (f.Length != 3)
                        {
                            f += "0";
                        }
                    }
                    if (n.Text.IndexOf("01") != -1)
                    {
                        offset++;
                        f += String.Format("{0:00}", (n.Index + 1 - i - offset)) + "-" + "01";
                    }
                    else
                        f += String.Format("{0:00}", (n.Index + 1 - i - offset));
                    dic.Add(n.Text, f.Insert(2, "."));
                    if (f.StartsWith("0"))
                        fsDic.Add(n.Text, sb + "B" + f + "B");
                    else
                        fsDic.Add(n.Text, sb + f + "B");
                }
            }
        }
    }

    public void numb(TreeNodeCollection node, string number, int num)
    {
        //int i = 0, offset = 0;

        foreach (TreeNode n in node)
        {
            if (n.Nodes.Count != 0)
            {
                if (n.Text != "Сборки" && num == 0 && n.Text.IndexOf(".iam") != -1)
                {
                    Offs++;
                    sb = ((char)Offs).ToString();
                }
                //if (n.Text.IndexOf(".iam") != -1)
                //TreeNode par = n.Parent;

                //if (par != null && par.Text == "Сборки")
                //{
                //    numb(n.Nodes, number + "." + (5).ToString(), num++);
                //}
                //else
                numb(n.Nodes, number + "." + (n.Index + 1).ToString(), num++);
                if (n.Text != "Сборки")
                {
                    string f = "";
                    try
                    { f = (number + "." + (n.Index + 1)).ToString().Substring(3); }
                    catch (Exception) { };
                    try
                    { NuM.Add(n.Text, f); }
                    catch (Exception)
                    {
                        n.Text += f;
                        NuM.Add(n.Text, f);
                    }
                    string[] tmp = f.Split('.');
                    //if (f == "1")
                    //    f = "A.";
                    //else
                    //{
                    f = "";
                    for (int j = 0; j < tmp.Length; j++)
                    {
                        if (Convert.ToInt16(tmp[j]) < 27)
                            f += ((char)(64 + Convert.ToInt16(tmp[j]))).ToString() + ".";
                        else f += ((char)(1014 + Convert.ToInt16(tmp[j]))).ToString() + ".";
                    }
                    //}
                    if (f.StartsWith(".")) f = f.Substring(1);
                    fsDic.Add(n.Text, sb + f);
                }
            }
            else
            {
                string f = (number + "." + (n.Index + 1)).ToString().Substring(3);
                try
                { NuM.Add(n.Text, f); }
                catch (Exception)
                {
                    n.Text = n.Text + f;
                    NuM.Add(n.Text, f);
                }
                string[] tmp = f.Split('.');
                f = "";
                for (int j = 0; j < tmp.Length; j++)
                {
                    if (Convert.ToInt16(tmp[j]) < 27)
                        f += ((char)(64 + Convert.ToInt16(tmp[j]))).ToString() + ".";
                    else f += ((char)(1014 + Convert.ToInt16(tmp[j]))).ToString() + ".";
                }
                if (f.StartsWith(".")) f = f.Substring(1);
                fsDic.Add(n.Text, sb + f);
            }
        }

    }

    public Dictionary<string, string> GetDic()
    {
        return dic;
    }

    private int GetCountTabs(string node)
    {
        string[] tmp = node.Split(new string[] { "    " }, StringSplitOptions.None);
        return tmp.Length - 1;
    }
    private List<string> GetItem(List<string> arr, int start, int num)
    {
        List<string> tmplst = new List<string>();
        for (int i = start; i < num; i++)
        {
            tmplst.Add(arr[i]);
        }
        return tmplst;
    }

    private void addElem(List<string> lst, TreeNode node, int count)
    {
        for (int i = 0; i < lst.Count; i++)
        {
            int num = GetCountTabs(lst[i]);
            if (count == num)
            {
                node.Nodes.Add(lst[i].Trim());
            }
            else
            {
                int start = i;
                try
                {
                    while (GetCountTabs(lst[i++]) >= num)
                    {
                        //if (i == lst.Count) break;
                    }
                }
                catch (Exception) { };
                i--;
                //if (node.PrevNode != null)
                //    node = node.PrevNode;
                try
                { node.LastNode.Remove(); }
                catch (Exception) { };
                addElem(GetItem(lst, start, i), node.Nodes.Add(lst[start - 1].Trim()), num);
                i--;
            }
        }
    }

    private void treeView1_ItemDrag(object sender, ItemDragEventArgs e)
    {
        // Move the dragged node when the left mouse button is used.
        if (e.Button == MouseButtons.Left)
        {
            DoDragDrop(e.Item, DragDropEffects.Move);
        }

        // Copy the dragged node when the right mouse button is used.
        else if (e.Button == MouseButtons.Right)
        {
            DoDragDrop(e.Item, DragDropEffects.Copy);
        }
    }

    // Set the target drop effect to the effect 
    // specified in the ItemDrag event handler.
    private void treeView1_DragEnter(object sender, DragEventArgs e)
    {
        e.Effect = e.AllowedEffect;
    }

    // Select the node under the mouse pointer to indicate the 
    // expected drop location.
    private void treeView1_DragOver(object sender, DragEventArgs e)
    {
        // Retrieve the client coordinates of the mouse position.
        System.Drawing.Point targetPoint = treeView1.PointToClient(new System.Drawing.Point(e.X, e.Y));

        // Select the node at the mouse position.
        treeView1.SelectedNode = treeView1.GetNodeAt(targetPoint);
    }

    private void treeView1_DragDrop(object sender, DragEventArgs e)
    {
        // Retrieve the client coordinates of the drop location.
        System.Drawing.Point targetPoint = treeView1.PointToClient(new System.Drawing.Point(e.X, e.Y));

        // Retrieve the node at the drop location.
        TreeNode targetNode = treeView1.GetNodeAt(targetPoint);

        // Retrieve the node that was dragged.
        TreeNode draggedNode = (TreeNode)e.Data.GetData(typeof(TreeNode));

        // Confirm that the node at the drop location is not 
        // the dragged node or a descendant of the dragged node.
        if (!draggedNode.Equals(targetNode) && !ContainsNode(draggedNode, targetNode))
        {
            // If it is a move operation, remove the node from its current 
            // location and add it to the node at the drop location.
            if (e.Effect == DragDropEffects.Move)
            {
                draggedNode.Remove();
                //targetNode.Nodes.Add(draggedNode);
                targetNode.Parent.Nodes.Insert(targetNode.Index, draggedNode);
            }

            // If it is a copy operation, clone the dragged node 
            // and add it to the node at the drop location.
            else if (e.Effect == DragDropEffects.Copy)
            {
                targetNode.Nodes.Add((TreeNode)draggedNode.Clone());
            }

            // Expand the node at the location 
            // to show the dropped node.
            targetNode.Expand();
        }
    }

    // Determine whether one node is a parent 
    // or ancestor of a second node.
    private bool ContainsNode(TreeNode node1, TreeNode node2)
    {
        // Check the parent node of the second node.
        if (node2.Parent == null) return false;
        if (node2.Parent.Equals(node1)) return true;

        // If the parent node is not null or equal to the first node, 
        // call the ContainsNode method recursively using the parent of 
        // the second node.
        return ContainsNode(node1, node2.Parent);
    }

    private void InitializeComponent()
    {
        this.btn = new System.Windows.Forms.Button();
        this.SuspendLayout();
        // 
        // btn
        // 
        this.btn.Location = new System.Drawing.Point(205, 234);
        this.btn.Name = "btn";
        this.btn.Size = new System.Drawing.Size(89, 41);
        this.btn.TabIndex = 0;
        if (!check)
            this.btn.Text = "Децимальные номера";
        else
            this.btn.Text = "Порядковые номера";
        this.btn.UseVisualStyleBackColor = true;
        this.btn.Click += new System.EventHandler(this.button1_Click);
        // 
        // FormTreeView
        // 
        this.ClientSize = new System.Drawing.Size(292, 273);
        this.Controls.Add(this.btn);
        this.Name = "FormTreeView";
        this.ResumeLayout(false);

    }

}
public class nodeSort : IComparer
{
    static TreeNode tx;
    static TreeNode ty;
    public int Compare(object x, object y)
    {
        tx = x as TreeNode;
        ty = y as TreeNode;

        List<string> val = InvAddIn.PropBtn.m_Prop.sort_val;
        // Compare the length of the strings, returning the difference.
        int ind1 = val.FindIndex(delegate(string str)
        {
            return tx.Text.ToUpper().IndexOf(str.ToUpper()) != -1;
        });
        //int ind1 = val.IndexOf(ind1str);
        int ind2 = val.FindIndex(delegate(string str)
        {
            return ty.Text.ToUpper().IndexOf(str.ToUpper()) != -1;
        });
        //int ind2 = val.IndexOf(ind2str);
        //int ind1 = val.FindIndex(indexX);
        //int ind2 = val.FindIndex(indexY);
        if (ind1 == -1 || ind2 == -1 || ind1 == ind2) return 0;
        return ind1 < ind2 ? -1 : 1;
    }
#endregion
    private static bool indexX(string str)
    {
        return tx.ToString().ToUpper().IndexOf(str) == -1 ? true : false;
    }
    private static bool indexY(string str)
    {
        return ty.ToString().ToUpper().IndexOf(str) == -1 ? true : false;
    }
}
