using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using mshtml;
using System.IO;
using System.Data.OleDb;
using System.Collections;
using System.Net;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class frmOpenWebBrowser : Form
    {
        int intStep = 3;
        DateTime dts = DateTime.Now;
        TreeView DOMTreeView4 = new TreeView();
        StringBuilder sb = new StringBuilder();
        StringBuilder sa = new StringBuilder();
        String newsText = null;
        RichTextBox richTextBox1 = new RichTextBox();
        RichTextBox richTextBox2 = new RichTextBox();
        RichTextBox richTextBox3 = new RichTextBox();
        int inttotlevel = 0;
        #region Instance
        private static frmOpenWebBrowser instance;

        public static frmOpenWebBrowser Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new frmOpenWebBrowser();
                }
                instance.Activate();
                return instance;
            }
        }
        #endregion

        #region 屬性
        private static string IsPath = "";
        private static bool IsBusy = false;
        private static bool IsError = false;
        private static string IsContect = "";
        private static string IsMinutes = "";
        private static string IsSeconds = "";
        private static string IsFullPath = "";
        private static int IsDepth = 0;

        #region Path
        /// <summary>
        /// 目前路徑
        /// </summary>
        public static string Path
        {
            get { return IsPath; }
            set { IsPath = value; }
        }
        #endregion

        #region Error
        /// <summary>
        /// 網頁是否有問題
        /// </summary>
        public static bool Error
        {
            get { return IsError; }
            set { IsError = value;}
        }
        #endregion

        #region Busy
        /// <summary>
        /// 處理是否在忙錄
        /// </summary>
        public static bool Busy
        {
            get { return IsBusy; }
            set { IsBusy = value;}
        }
        #endregion

        #region Contect
        /// <summary>
        /// 網頁新聞內容-txt
        /// </summary>
        public static string Contect
        {
            get { return IsContect; }
            set { IsContect = value; }
        }
        #endregion

        #region Minutes
        /// <summary>
        /// 逾時分的設定
        /// </summary>
        public static string Minutes
        {
            get { return IsMinutes; }
            set { IsMinutes = value; }
        }
        #endregion

        #region Seconds
        /// <summary>
        /// 逾時秒的設定
        /// </summary>
        public static string Seconds
        {
            get { return IsSeconds; }
            set { IsSeconds = value; }
        }
        #endregion

        #region FullPath
        /// <summary>
        /// 路徑
        /// </summary>
        public static string FullPath
        {
            get { return IsFullPath; }
            set { IsFullPath = value; }
        }
        #endregion

        #region Depth
        /// <summary>
        /// 深度
        /// </summary>
        public static int Depth
        {
            get { return IsDepth; }
            set { IsDepth = value; }
        }
        #endregion
        #endregion

        #region frmOpenWebBrowser
        /// <summary>
        /// 畫面元件建置子
        /// </summary>
        public frmOpenWebBrowser()
        {
            InitializeComponent();
        }
        #endregion

        #region frmOpenWebBrowser(string strUrl)
        /// <summary>
        /// 畫面元件建置子
        /// </summary>
        public frmOpenWebBrowser(string strUrl)
        {
            InitializeComponent();

            if (!strUrl.Equals(""))
            {
                webBrowser1.ScriptErrorsSuppressed = true;
                //webBrowser1.Navigate(strUrl);
                
                if (Ping(strUrl))
                {
                    webBrowser1.Navigate(strUrl);
                    IsError = false;
                    IsBusy = true;
                }
                else
                {
                    IsError = true;
                }
                 
            }

        }
        #endregion

        #region frmOpenWebBrowser_Load
        /// <summary>
        /// 畫面載入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmOpenWebBrowser_Load(object sender, EventArgs e)
        {
            txtErrorCode.Text = "";
            //txtDataString.Text = "";
            timer1.Enabled = true;
            timer1.Interval = 360;
            progressBar1.Maximum = 100;
            progressBar1.Minimum = 0;
        }
        #endregion

        #region frmOpenWebBrowser_FormClosed
        /// <summary>
        /// 畫面關閉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmOpenWebBrowser_FormClosed(object sender, FormClosedEventArgs e)
        {
            instance = null;
            timer1.Enabled = false;
        }
        #endregion

        #region webBrowser1_DocumentCompleted
        /// <summary>
        /// 輸入網址載入web
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            ArrayList aryError = new ArrayList();
            aryError.Add("您沒有檢視此網頁的授權");
            aryError.Add("必須在安全通道上檢視此網頁");
            aryError.Add("無法顯示資源");
            aryError.Add("需要 Proxy 驗證");
            aryError.Add("網頁不存在");
            aryError.Add("找不到網站");
            aryError.Add("找不到伺服器");
            aryError.Add("沒有網頁可顯示");

            try
            {
                webBrowser1.Document.Window.Error += new HtmlElementErrorEventHandler(webBrowser1_WindowError);
                if (webBrowser1.ReadyState == WebBrowserReadyState.Complete )
                {
                   
                    //Thread.Sleep(10);
                    DOMTreeView4.Nodes.Clear();
                    IHTMLDocument3 doc = webBrowser1.Document.DomDocument as IHTMLDocument3;
                    IHTMLDOMNode rootDomNode = (IHTMLDOMNode)doc.documentElement;
                    TreeNode root = DOMTreeView4.Nodes.Add("<HTML>");
                    root.Tag = rootDomNode;
                    //ProcessElement(htmlElement, root3);
                    inttotlevel = 1;
                    InsertDOMNodes(rootDomNode, root, inttotlevel, 1);
                    DOMTreeView4.SelectedNode = DOMTreeView4.Nodes[0];
                    DOMTreeView4.Select();
                    //this.webBrowser1.DocumentCompleted -= new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted); //停止webBrowser3_DocumentCompleted事件
                    //Parser的函數 BtnDeleteTag()為去除不相關Tag BtnExcuteDOMTree()載入編輯過後的DOMTree BtnInfon_Block() 區塊處理  Btn_Group_Info() 區塊判斷
                    //this.webBrowser1.DocumentCompleted -= new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted); //停止webBrowser2_DocumentCompleted事件
                    BtnDeleteTag();
                    intStep = 3;
                    BtnExcuteDOMTree();
                    BtnInfon_Block();
                    intStep = 2;
                    btn_Group_Info();
                    intStep = 1;
                    //newsText = richTextBox3.Text;
                    IsContect = newsText;
                    //webBrowser1.Document.ExecCommand("SaveAs", false, IsPath);
                    //處理完成                   
                    progressBar1.Value = progressBar1.Maximum;
                    IsBusy = false;
                    this.Close();
                    return;
                }
                
            }
            catch (InvalidOperationException ex)
            {
                IsError = true;
                this.Close();
            }
            catch (OleDbException ex)
            {
                IsError = true;
                this.Close();
            }
            catch (Exception ex)
            {
                IsError = true;
                this.Close();
            }
        }
        #endregion

        #region webBrowser1_NewWindow
        /// <summary>
        /// 避免打開其它廣告視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_NewWindow(object sender, CancelEventArgs e)
        {
            txtErrorCode.Text = "關閉廣告視窗";
            e.Cancel = true;
        }
        #endregion

        #region webBrowser1_WindowError
        /// <summary>
        /// 避免有錯誤碼出現
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_WindowError(object sender, HtmlElementErrorEventArgs e)
        {
            txtErrorCode.Text = "關閉錯誤視窗";
            e.Handled = true;
        }
        #endregion

        #region timer1_Tick
        /// <summary>
        /// 處理速度
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (IsError == true)
            {
                this.Close();
                return;
            }

            if (progressBar1.Value < 100)
            {
                progressBar1.Value = progressBar1.Value + 1;
            }
            else
            {
                if (treeView1.Nodes.Count == 0)
                {
                    System.TimeSpan diff = UsedTime(dts);

                    //判斷是否超過逾時設定
                    if (diff.Minutes * 60 + diff.Seconds >= Convert.ToInt16(IsMinutes) * 60 + Convert.ToInt16(IsSeconds))
                    {
                        IsError = false;
                        IsBusy = false;
                        this.Close();
                        return;
                    }
                    else
                    {
                        progressBar1.Value = progressBar1.Maximum - intStep * 15;
                    }
                }
                else
                {
                    progressBar1.Value = progressBar1.Maximum - intStep * 15;
                }
            }
        }
        #endregion

        #region CompTime
        /// <summary>
        /// 計算耗用時間
        /// </summary>
        /// <returns></returns>
        private TimeSpan UsedTime(DateTime dts)
        {
            //讀出新聞終止時間
            DateTime dte = DateTime.Now;

            //計算共費多少分秒
            System.TimeSpan diff = dte - dts;

            return (diff);
        }
        #endregion

        #region Ping
        /// <summary>
        /// 通指定的網頁是否連通
        /// </summary>
        /// <param name="ip"></param>
        /// <returns></returns>
        public bool Ping(string strUrl)
        {
            //Uri myUri = new Uri(strUrl);
            var decodeurl = Uri.UnescapeDataString(strUrl);
            System.Net.HttpWebRequest myWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(decodeurl);
            myWebRequest.Timeout = 1200000;
            myWebRequest.Proxy = null;
            System.Net.HttpWebResponse myWebResponse = (System.Net.HttpWebResponse)myWebRequest.GetResponse();
            myWebResponse.Close();
            if (myWebResponse.StatusCode == System.Net.HttpStatusCode.OK)
                return true;
            else
                return false;
        }
        #endregion

        #region 自動化處理新聞網頁之函數合集
        void BtnDeleteTag()
        {
            try
            {
                Global com = new Global();
                //com.DeleteNode(DOMTreeView3, DOMTreeView3.Nodes[0]);
                com.DeleteNode2(DOMTreeView4, DOMTreeView4.Nodes[0]);
                com.DeleteDataArea1(DOMTreeView4, DOMTreeView4.Nodes[0]);
                com = null;
                //Thread.Sleep(1000);
                //MessageBox.Show("去除不相關標記成功", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        void BtnExcuteDOMTree()
        {
            try
            {
                if (treeView1.Nodes == null)
                {
                    MessageBox.Show("請載入TreeNode");
                }
                else
                {
                    this.DOMTreeView4.CollapseAll();    //閉合(縮合)DOM_Tree以確保下行指令能抓到根節點
                    TreeNode node = (TreeNode)this.DOMTreeView4.TopNode.Clone();
                    this.treeView1.Nodes.Clear();
                    this.treeView1.Nodes.Add(node);
                    this.treeView1.ExpandAll();
                    //sa.Clear();
                    //sb.Clear();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        int Minweighnum = 0;
        ArrayList weightlist = new ArrayList();
        void BtnInfon_Block()
        {
            weightlist.Clear();
            txtlist.Clear();
            Minweighnum = 0;
            txtleath2 = 0;
            leath = 0;
            size = 0;
            PrintRecursive2(treeView1.Nodes[0]);
        }
        private void PrintRecursive2(TreeNode treeNode)
        {
            int weightnum = 0;
            if (treeNode.NextVisibleNode != null)
            {
                if (treeNode.NextVisibleNode.Nodes.Count == 0)//如果是葉節點
                {
                    weightnum = TreenodeGroup(treeNode, weightnum, Minweighnum);
                    if (weightnum > Minweighnum && leath > txtleath2)
                    {
                        txtleath2 = leath;
                        Minweighnum = weightnum;
                        weightlist.Add(Minweighnum);
                        treeNode.Parent.Text = weightnum.ToString();
                        //treeNode.Parent.Parent.Tag = weightnum.ToString();
                    }
                }
            }
            // 遞迴,依序尋找節點
            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive2(tn);
            }
        }
        /// <summary>
        /// 計算權重值
        /// </summary>
        /// <param name="node"></param>
        /// <param name="weight"></param>
        /// <returns></returns>
        int txtleath2 = 0;
        ArrayList txtlist = new ArrayList();
        int leath = 0;
        int size = 0;
        public int TreenodeGroup(TreeNode node, int weight = 0, int Minweigh = 0, String txtnode = null)
        {
            //往上尋找
            if (node.Parent != null)
            {
                mshtml.IHTMLDOMNode domnode = null;
                int nodeCount = node.Parent.Nodes.Count;
                TreeNode parentnode = node.Parent;
                int index = node.Index;
                int txtleath = 0;
                foreach (TreeNode childnode in parentnode.Nodes)
                {
                    if (childnode.Parent.Equals(node.Parent)) //如果 兩個node路徑相同時回傳 權重值 
                    {
                        if (childnode.Tag is mshtml.IHTMLDOMNode) // 如果此節點是mshtml
                        {
                            if (childnode.NextVisibleNode != null)
                                domnode = childnode.NextVisibleNode.Tag as mshtml.IHTMLDOMNode;
                            else
                                domnode = childnode.Tag as mshtml.IHTMLDOMNode;
                            if (childnode.Nodes.Count == 1)
                            {
                                if (domnode.nodeName == "#text") //如果是葉節點(text node)
                                {

                                    if (childnode.NextVisibleNode.Text.IndexOf("\n") != -1)
                                    {
                                        txtnode = childnode.NextVisibleNode.Text;
                                        txtlist.Add(txtnode);
                                        String[] sections = null;
                                        String sectionstxt = null;
                                        sectionstxt = childnode.NextVisibleNode.Text.Replace("\n", "~");
                                        sections = sectionstxt.Split(new char[] { '~' }); // 藉由,分割字串成為陣列
                                        for (int i = 0; i < sections.Length; i++)
                                        {
                                            sectionstxt = sections[i];
                                            sectionstxt = sectionstxt.Trim();
                                            txtleath = sectionstxt.Length;
                                        }
                                    }
                                    if (childnode.NextVisibleNode.Text.IndexOf("\n") == -1)
                                        if (!txtlist.Contains(childnode.NextVisibleNode.Text))
                                        {
                                            txtnode = childnode.NextVisibleNode.Text;
                                            txtlist.Add(txtnode);
                                            if (childnode.Index == 0)
                                                txtleath = childnode.NextVisibleNode.Text.Length;
                                            if (childnode.Index != 0)
                                                txtleath += childnode.NextVisibleNode.Text.Length;
                                        }
                                    /*
                                    weight = node.NextVisibleNode.Text.Length / parentnode.Nodes.Count; // weight為權重值
                                    if(weight > Minweigh)
                                            Minweigh = weight;
                                     */
                                    //return TreenodeGroup(node.Parent, weight,Minweigh,txtnode);
                                }
                                else
                                {
                                    nodeCount--;
                                }
                            }
                            else
                            {
                                if (childnode.Nodes.Count > 1)
                                {
                                    foreach (TreeNode tt in childnode.Nodes)
                                    {
                                        mshtml.IHTMLDOMNode domnode1 = tt.Tag as mshtml.IHTMLDOMNode;
                                        if (domnode1.nodeName == "A")
                                        {
                                            continue;
                                        }
                                        if (domnode1.nodeName == "#text") //如果是葉節點(text node)
                                        {
                                            if (tt.Index == 0)
                                            {

                                                if (!txtlist.Contains(tt.Text))
                                                {
                                                    txtlist.Add(tt.Text);
                                                    txtleath += tt.Text.Length;
                                                }
                                            }
                                            else
                                            {

                                                if (!txtlist.Contains(tt.Text))
                                                {
                                                    txtlist.Add(tt.Text);
                                                    txtleath += tt.Text.Length;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            nodeCount--;
                                        }
                                    }

                                }
                            }
                            if(childnode .NextVisibleNode != null)
                            if (childnode.NextVisibleNode.Nodes.Count == 0 && childnode.NextVisibleNode.Index == 0 && childnode.NextVisibleNode.Text != txtnode && !txtlist.Contains(childnode.NextVisibleNode.Text))
                            {
                                txtlist.Add(childnode.NextVisibleNode.Text);
                                txtleath += childnode.NextVisibleNode.Text.Length;
                            }
                        }
                    }
                    else  //不同時 則 往另一個路徑
                    {

                    }
                }
                if (nodeCount != 0)
                {
                    weight = txtleath / nodeCount;
                    weight = Math.Abs(weight);
                }
                    /*
                else
                {
                    weight = txtleath;
                }
                     */ 

                if (weight > Minweigh && txtleath > leath)
                {
                    leath = txtleath;
                    Minweigh = weight;

                }
                return weight;
            }
            //直到上層節點沒有為止
            else
            {
                return weight;
            }
        }
        StringBuilder SB = new StringBuilder();
        void btn_Group_Info()
        {
            SB.Clear();
           // PrintRecursive3(treeView1.Nodes[0]);
            GetInfoGroup();
        }
        private void PrintRecursive3(TreeNode treeNode)
        {
            foreach (TreeNode tn in treeNode.Nodes)
            {
                mshtml.IHTMLDOMNode domnode = tn.Tag as mshtml.IHTMLDOMNode;

                if (treeNode.Text.IndexOf("<") == -1 && domnode.nodeName != "#text")
                {
                    weightlist.Add(treeNode.Text);
                }
                else
                {
                    PrintRecursive3(tn);
                }

            }
        }
        String max = null;
        public void GetInfoGroup()
        {
            
            for (int i = 0; i < weightlist.Count; i++)
            {
                for (int j = i + 1; j < weightlist.Count; j++)
                {
                    if (Convert.ToInt32(weightlist[i]) > Convert.ToInt32(weightlist[j]))
                    {
                        int temp = Convert.ToInt32(weightlist[i]);
                        weightlist[i] = weightlist[j];
                        weightlist[j] = temp;
                    }
                }
            }
            max = Convert.ToString(weightlist[weightlist.Count - 1]);
            PrintRecursive4(treeView1.Nodes[0]);
        }
        private void PrintRecursive4(TreeNode treeNode)
        {
            mshtml.IHTMLDOMNode domnode = null;
            //StringBuilder SB = new StringBuilder();
            if (treeNode.Text.Contains(max))
            {
                foreach (TreeNode tt in treeNode.Nodes)
                {

                    if (tt.Nodes.Count > 1)
                    {
                        foreach (TreeNode tn in tt.Nodes)
                        {
                            domnode = tn.Tag as mshtml.IHTMLDOMNode;
                            if (tn.Text.IndexOf(">") == -1 && domnode.nodeName == "#text")
                                SB.AppendLine(tn.Text);
                            else
                            {
                                domnode = tn.NextVisibleNode.Tag as mshtml.IHTMLDOMNode;
                                if (tn.NextVisibleNode.Text.IndexOf(">") == -1 && domnode.nodeName == "#text")
                                    SB.AppendLine(tn.NextVisibleNode.Text);
                            }
                        }
                    }
                    else
                    {
                        mshtml.IHTMLDOMNode domnode1 = tt.NextVisibleNode.Tag as mshtml.IHTMLDOMNode;
                        if (tt.NextVisibleNode.Text.IndexOf(">") == -1 && domnode1.nodeName == "#text")
                            SB.AppendLine(tt.NextVisibleNode.Text);
                    }
                }
                newsText = SB.ToString();
            }
            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive4(tn);
            }
        }
        #endregion
        /*
        void BtnInfon_Block()
        {
            richTextBox1.Clear();
            richTextBox2.Clear();
            sa.Clear();
            sb.Clear();
            CallRecursive(this.treeView1); //遞迴
        }
        void btn_Group_Info()
        {
            try
            {
                String[] sections = null;       //儲存每一行字串之分段
                String[] sections_next = null;  //儲存每一行字串之分段
                String[] sections_next1 = null;  //儲存每一行字串之分段
                int words = 0;      //存每行字串後之文字的字數
                int arrayleath = 0;
                int dup = 0;
                int startleath = 0, endleath = 0, weight = 0, Max_weight = 0;
                sa.Clear();
                String domString = sb.ToString();
                String[] lines = domString.Split(new char[] { '\n' }); //根據尾端的空白字元切割成字串陣列
                int lineNo = 0;     //設定起始行號
                while (lineNo < lines.Length)    //從第一行開始至文件最後一行(lines.Length)
                {
                    if (lines[lineNo].IndexOf('︶') != -1) //當這行有數字時
                    {
                        sections = lines[lineNo].Split(new char[] { '↖' }); // 藉由,分割字串成為陣列
                        dup = Int32.Parse(sections[0].Substring(sections[0].IndexOf('︵') + 1, (sections[0].IndexOf('︶') - sections[0].IndexOf('︵') - 1))); //擷取出數字
                        if (sections.Length == 3)
                        {
                            words = sections[2].Length;
                        }
                        else if (sections.Length == 2)
                        {
                            words = sections[1].Length;
                        }
                        else if (sections.Length == 1)
                        {
                            words = sections[0].Length;
                        }
                        for (int i = 1; i <= dup; i++)
                        {
                            arrayleath = lineNo + i + 1;
                            sections_next = lines[lineNo + i].Split(new char[] { '↖' });

                            if (arrayleath < lines.Length)
                            {
                                sections_next1 = lines[lineNo + i + 1].Split(new char[] { '↖' });
                            }

                            //sections_next1 = lines[lineNo + i + 1].Split(new char[] { '/' });
                            //if (sections_next[0].IndexOf(sections [1]) >0 )
                            if (sections[1].IndexOf(sections_next[0]) != -1)
                            //if (sections[1].Equals(sections_next[0]))
                            {
                                if (sections_next.Length == 2)
                                {
                                    words += sections_next[1].Length;
                                }
                                if (sections_next.Length == 1)
                                {
                                    words += sections_next[0].Length;
                                }
                            }
                            else
                            {

                                if (sections_next1[0].IndexOf(sections_next[0]) != -1)
                                {
                                    if (sections_next.Length == 2)
                                    {
                                        words += sections_next[1].Length;
                                    }
                                    if (sections_next.Length == 1)
                                    {
                                        words += sections_next[0].Length;
                                    }
                                    if (sections_next1.Length == 2)
                                    {
                                        words += sections_next1[1].Length;
                                    }
                                    if (sections_next1.Length == 1)
                                    {
                                        words += sections_next1[0].Length;
                                    }
                                    continue;
                                }
                                else
                                {
                                    if (sections_next.Length == 2)
                                    {
                                        words += sections_next[1].Length;
                                    }
                                    if (sections_next.Length == 1)
                                    {
                                        words += sections_next[0].Length;
                                    }
                                    if (sections_next1.Length == 2)
                                    {
                                        words += sections_next1[1].Length;
                                        continue;
                                    }
                                    if (sections_next.Length == 1)
                                    {
                                        words += sections_next1[0].Length;
                                        continue;
                                    }
                                }
                                weight = words / dup; //權值計算
                                if (weight > Max_weight)
                                {
                                    startleath = lineNo;
                                    endleath = lineNo + dup;
                                    Max_weight = weight;
                                }
                                break;
                            }
                        }
                        weight = words / dup; //權值計算
                        if (weight > Max_weight)
                        {
                            startleath = lineNo;
                            endleath = lineNo + dup;
                            Max_weight = weight;
                        }
                        lineNo++;
                    }
                    else    //該行無數字,換處理下一行
                    {
                        lineNo++;
                    }
                }

                while (startleath < endleath)
                {
                    String Strnew = lines[startleath];
                    String[] Strnew_split = lines[startleath].Split(new char[] { '↖' });
                    if (Strnew_split.Length == 3)
                    {
                        Strnew = Strnew_split[2];
                        if (Strnew.IndexOf("<") != -1)
                            Strnew = Strnew.Replace(Strnew, String.Empty);
                    }
                    else if (Strnew_split.Length == 2)
                    {
                        Strnew = Strnew_split[1];
                        if (Strnew.IndexOf("<") != -1)
                            Strnew = Strnew.Replace(Strnew, String.Empty);
                    }
                    else if (Strnew_split.Length == 1)
                    {
                        Strnew = Strnew_split[0];
                        if (Strnew.IndexOf("<") != -1)
                            Strnew = Strnew.Replace(Strnew, String.Empty);
                    }
                    sa.AppendLine("↖" + Strnew + "↖");
                    startleath++;
                }
                richTextBox3.Text = sa.ToString();
                newsText = richTextBox3.Text;
                //richTextBox5.Text = newsTitle + "\n" + newsSource + newsDate + "\n" + newsText;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion

        #region 遞迴
        private void CallRecursive(TreeView treeView)           // Call the procedure using the TreeView.
        {
            TreeNodeCollection nodes = treeView.Nodes;  //Nodes 屬性可以用於取得包含樹狀結構中所有根節點的 TreeNodeCollection 物件
            foreach (TreeNode n in nodes)
            {
                PrintRecursive(n);      // Print each node recursively.
            }
        }

        /// <summary>
        /// 遞迴依序尋找節點並列印出路徑
        /// </summary>
        /// <param name="treeNode"></param>
        private void PrintRecursive(TreeNode treeNode)
        {
            if (treeNode.Tag is mshtml.IHTMLDOMNode)    //如果該節點非葉節點
            {
                mshtml.IHTMLDOMNode domnode = treeNode.Tag as mshtml.IHTMLDOMNode;
                if (domnode.nodeName != "#text")
                {
                    if (treeNode.Nodes.Count > 1)   //該節點的子節點(分支數)大於1
                    {
                        sa.Append(string.Format("{0}~︵{1}︶↖", treeNode.Text, treeNode.Nodes.Count));  //取得該節點的Tag_name及該節點分支數
                        sb.Append(string.Format("{0}~︵{1}︶↖", treeNode.Text, treeNode.Nodes.Count));  //取得該節點的Tag_name及該節點分支數
                    }
                    else
                    {
                        sa.Append(string.Format("{0}", treeNode.Text));    //取得該節點的Tag_name
                        sb.Append(string.Format("{0}", treeNode.Text));    //取得該節點的Tag_name
                    }
                }
                else
                {
                    sa.Append(string.Format("↖{0} \n", treeNode.Text));  //每一分支結尾,取得該葉節點上的文字
                    sb.Append(string.Format("↖{0} \n", treeNode.Text));  //每一分支結尾,取得該葉節點上的文字
                }
            }
            /*
        else          //該節點為葉節點
        {
            sa.Append(string.Format(",{0} \n", treeNode.Text));  //每一分支結尾,取得該葉節點上的文字
            sb.Append(string.Format(",{0} \n", treeNode.Text));  //每一分支結尾,取得該葉節點上的文字
        }
             

            // 遞迴,依序尋找節點
            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive(tn);
            }

            //結果輸出
            richTextBox1.Text = sa.ToString();
            sb.Replace("~", "\n");
            richTextBox2.Text = sb.ToString();
        }

        #endregion
        */
        #region InsertDOMNodes
        private void InsertDOMNodes(IHTMLDOMNode parentnode, TreeNode tree_node, int parentnode_level, int node_level)
        {
            string strtotlevel = "";
            string strparentlevel = "";
            string strnowlevel = "";

            if (parentnode.hasChildNodes())
            {
                IHTMLDOMChildrenCollection allchild = (IHTMLDOMChildrenCollection)parentnode.childNodes;
                int length = allchild.length;
                string strNode = "";

                for (int i = 0; i < length; i++)
                {
                    IHTMLDOMNode child_node = (IHTMLDOMNode)allchild.item(i);

                    if (child_node.nodeName != "#comment")
                    {
                        if (child_node.nodeName != "#text")
                        {
                            strNode = "<" + (string)child_node.nodeName.ToString().ToUpper().Trim() + ">";

                        }
                        else
                        {
                            if (child_node.nodeValue != null)
                            {
                                strNode = (string)child_node.nodeValue.ToString().ToUpper().Trim();
                            }
                            else
                            {
                                continue;
                            }
                        }
                        if (strNode != "")
                        {
                            TreeNode tempnode = tree_node.Nodes.Add(strNode);
                            tempnode.Tag = child_node;
                            strtotlevel = inttotlevel.ToString().PadLeft(4, '0');				//從根到目前層級
                            strparentlevel = (parentnode_level - 1).ToString().PadLeft(4, '0');	//父層級
                            strnowlevel = node_level.ToString().PadLeft(4, '0');					//目前層級序號

                            inttotlevel += 1;
                            //tempnode.Tag = strtotlevel + "\t" + strparentlevel + "\t" + strnowlevel;
                            node_level += 1;

                            //tempnode.Text = tempnode.Tag + tempnode.Text;
                            InsertDOMNodes(child_node, tempnode, inttotlevel, 1);
                        }
                    }

                }
            }
        }
        #endregion
    }
}
