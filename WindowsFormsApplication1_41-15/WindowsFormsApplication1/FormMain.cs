using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;
using System.Collections;
using System.Diagnostics;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using mshtml;

namespace WindowsFormsApplication1
{
    public partial class FormMain : Form
    {
        #region 變數宣告
        /// <summary>紀錄上次的Html Element</summary>
        HtmlElement _lastNode = null;
        string _hightlightStyle = "border:2px #FF0000 solid", _lastNodeStyle = "";
        static int TotalDocNum = 0, TotalTermNum = 0;
        
        String strMessage;
        int inttotlevel = 0;
        private int intTotCount = 1;
        private int intPageCount = 1;
        private int intCount = 1;
        //int intAddCount = 0;
        private static readonly string strCurrentPath = AppDomain.CurrentDomain.BaseDirectory; //專案本身的基底位址 為Bin底下
        WebBrowser webBrowser3 = new WebBrowser();
        private bool bolExecuteFlag = false;
        TreeView DOMTreeView4 = new TreeView();
        private String strSelectDir = "";
        private String strSelectDir1 = "";
        #region 新聞內容變數
        ///<sumary>新聞內容變數</sumary>
        String newsText = null;
        String newsHref = null;
        String newsTitle = null;
        String newsSource = null;
        String newsDate = null;
        StringBuilder newssb = new StringBuilder();
        #endregion
        /// <summary>資料表</summary>>
        DataTable dt = new DataTable();
        #endregion

        #region 畫面元件建構子
        public FormMain()
        {
            InitializeComponent();
        }
        #endregion

        #region FormMain_Load
        private void FormMain_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
            toolStripStatusLabel2.Text = "";
            toolTip1.SetToolTip(this.txtKeyword1, "請輸入關鍵字");
            timer1.Enabled = true;
            Form.CheckForIllegalCrossThreadCalls = false;

            ResetdataGridVeiw();
            dt.Columns.Add(new DataColumn("序號", typeof(string)));             //序號
            dt.Columns.Add(new DataColumn("網址", typeof(string)));             //網址
            dt.Columns.Add(new DataColumn("標題", typeof(string)));             //標題
            dt.Columns.Add(new DataColumn("來源", typeof(string)));             //來源
            dt.Columns.Add(new DataColumn("日期", typeof(string)));             //日期
            dt.Columns.Add(new DataColumn("狀態", typeof(string)));             //狀態
        }
        #endregion

        #region btnSearch
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtKeyword1.Text.ToString() == "")
                {
                     MessageBox.Show("請輸入關鍵字", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (!comboBox1.Items.Contains(txtKeyword1.Text))
                    {
                        comboBox1.Items.Add(txtKeyword1.Text);
                    }
                    if (rdoGoogle.Checked)
                    {
                        webBrowser1.Navigate(txtnewsAddress.Text + txtKeyword1.Text);
                    }
                    if (rdoyahoonews.Checked)
                    {
                        webBrowser1.Navigate(txtnewsAddress.Text + Uri.EscapeUriString(txtKeyword1.Text));
                    }
                    /*
                    if (rdomsnnews.Checked)
                    {
                        webBrowser1.Navigate(txtnewsAddress.Text + Uri.EscapeUriString(txtKeyword1.Text));
                    }
                     */ 
                }
            }
            catch (Exception ex)
            {
                string strMessage = String.Format("載入網頁失敗 \n{0}", ex.ToString());
                MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        //上面程式"網頁載入"成功後即觸發此事件,轉換網頁相對應的DOM_Tree
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                DOMTreeView1.Nodes.Clear();
                HtmlDocument dom = webBrowser1.Document; 
                TreeNode root = DOMTreeView1.Nodes.Add("<HTML>");
                //HtmlElement HtmlElement = dom.GetElementsByTagName("HTML")[0];
                HtmlElement el = webBrowser1 .Document.ActiveElement.Parent ;
                //HtmlElementCollection emc = webBrowser1.Document.All;
                root.Tag = el;
                ProcessElement(el, root);
                //ProcessElement1(emc, DOMTreeView1.Nodes);
                DOMTreeView1.SelectedNode = DOMTreeView1.Nodes[0];
                DOMTreeView1.Select();
                DOMTreeView1.EndUpdate();
                if (bolExecuteFlag)
                {
                    if(rdoyahoonews .Checked)
                    YahooRecursive(dt);
                    if (rdoGoogle.Checked)
                        GoogleRecursive1(dt);

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        public void ProcessElement(HtmlElement parentelement, TreeNode nodes)
        {
            foreach (HtmlElement element in parentelement.Children)
            {
                element.Enabled = true;
                TreeNode node = new TreeNode();
                node.Text = "<" + element.TagName + ">";
                nodes.Nodes.Add(node);
                node.Tag = element;
                if (!string.IsNullOrEmpty(element.TagName) && element.TagName.ToLower() == "script")
                    continue;
                if (!string.IsNullOrEmpty(element.TagName) && element.TagName.ToLower() == "style")
                    continue;
                if (!string.IsNullOrEmpty(element.TagName) && element.TagName.ToLower() == "noscript")
                    continue;
                if ((element.Children.Count == 0) && (element.InnerText != null))
                {
                    node.Nodes.Add(element.InnerText );
                }
                else
                {
                    ProcessElement(element , node);   
                }
            }
        }

        private void ProcessElement1(HtmlElementCollection elements, TreeNodeCollection nodes)
        {
            foreach (HtmlElement element in elements)
            {
                TreeNode node = new TreeNode("<" + element.TagName + ">");
                nodes.Add(node);
                if ((element.Children.Count == 0) && (element.TagName != null))
                {
                    node.Nodes.Add(element.OuterText);
                }
                else
                {
                    ProcessElement1(element.Children, node.Nodes);
                }
            }
        }

        #endregion

        #region btnDeleteTag_Click
        private void btnDeleteTag_Click(object sender, EventArgs e)
        {
            try
            {
                if (DOMTreeView1.Nodes.Count == 0)
                {
                    MessageBox.Show("尚未讀出檔案", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    Global com = new Global();
                    com.MarkDeleteNode(DOMTreeView1, DOMTreeView1.Nodes[0]);
                    com = null;
                    toolStripStatusLabel1.Text = "刪除完成";
                    Thread.Sleep(1000);
                    MessageBox.Show("去除不相關標記成功", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    HtmlDocument hd = webBrowser1.Document;
                    
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region timer1.Tick
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = System.DateTime.Now.ToString();
        }
        #endregion

        #region btnExcute_Click
        /// <summary>
        /// 載入網頁,需使用者先輸入網址.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void btnExcute_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtHttp.Text.ToString() == "")      //判斷使用者是否輸入網址
                {
                    MessageBox.Show("請輸入網址", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                this.webBrowser2.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser2_DocumentCompleted); //啟動webBrowser2_DocumentCompleted事件
                webBrowser2.Navigate(txtHttp.Text);     //載入網頁
            }
            catch (Exception ex)                        //若網頁載入失敗,顯示相關錯誤訊息
            {
                string strMessage = String.Format("載入網頁失敗 \n{0}", ex.ToString());
                MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //上面程式"網頁載入"成功後即觸發此事件,轉換網頁相對應的DOM_Tree
        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
           
            try
            {
           if (webBrowser2.ReadyState.ToString() == "Complete")//判斷網頁是否讀取完成
              {
                DOMTreeView2.Nodes.Clear();
                DOMTreeView3.Nodes.Clear();
                /*
                HtmlDocument dom = webBrowser2.Document;
                TreeNode root = DOMTreeView2.Nodes.Add("<HTML>");
                TreeNode root2 = DOMTreeView3.Nodes.Add("<HTML>");
                //TreeNode root3 = DOMTreeView4.Nodes.Add("<HTML>");
                HtmlElement htmlElement = dom.GetElementsByTagName("HTML")[0];
                richTextBox4.Text = htmlElement.InnerHtml;
                HtmlElementCollection em = webBrowser2.Document.All;
                root.Tag = htmlElement;
                root2.Tag = htmlElement;
               // root3.Tag = htmlElement;
               // ProcessElement(htmlElement, root3);
                ProcessElement(htmlElement, root);
                //InsertNode(htmlElement, root);
                ProcessElement(htmlElement, root2);
                //InsertDOMTree(htmlElement, root2);
                DOMTreeView2.SelectedNode = DOMTreeView2.Nodes[0];
                DOMTreeView2.Select();
                 */
               //jQuery
                HtmlDocument dom = webBrowser2.Document;
                HtmlElement HE = dom.CreateElement("SCRIPT");
                HE.SetAttribute("src", "https://code.jquery.com/jquery-1.11.1.min.js");
                dom.Body.AppendChild(HE);
                HtmlElement HE2 = dom.CreateElement("SCRIPT");
                HE2.SetAttribute("text",Properties .Resources .attrNodeTopLeft);
                dom.Body.AppendChild(HE2);
                dom.InvokeScript("attrNodeTopLeft");

                HtmlElement htmlElement = dom.GetElementsByTagName("HTML")[0];
                richTextBox4.Text = htmlElement.InnerHtml;
                IHTMLDocument3 doc = webBrowser2.Document.DomDocument as IHTMLDocument3;
                IHTMLDOMNode rootDomNode = (IHTMLDOMNode)doc.documentElement;
                TreeNode root = DOMTreeView2.Nodes.Add("<HTML>");
                root.Tag = rootDomNode;
                TreeNode root2 = DOMTreeView3.Nodes.Add("<HTML>");
                root2.Tag = rootDomNode;
                inttotlevel = 1;
                InsertDOMNodes(rootDomNode, root, inttotlevel, 1);
                InsertDOMNodes(rootDomNode, root2, inttotlevel, 1);
              
                    DOMTreeView2.ExpandAll();
                    this.webBrowser2.DocumentCompleted -= new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser2_DocumentCompleted); //停止webBrowser2_DocumentCompleted事件    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion

        #region btnDeleteNode_Click
        /// <summary>
        /// btnDeleteNode_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteNode_Click(object sender, EventArgs e)
        {
            try
            {
                if (DOMTreeView2.Nodes.Count == 0)
                {
                    MessageBox.Show("尚未讀出檔案", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    Global com = new Global();
                    com.MarkDeleteNode2(DOMTreeView2, DOMTreeView2.Nodes[0]);
                    //com.DeleteNode(DOMTreeView3, DOMTreeView3.Nodes[0]);
                    com.DeleteNode2(DOMTreeView3, DOMTreeView3.Nodes[0]);
                    com.DeleteDataArea1(DOMTreeView3, DOMTreeView3.Nodes[0]); 
                    com = null;
                    toolStripStatusLabel1.Text = "刪除完成";
                    //Thread.Sleep(1000);
                    //MessageBox.Show("去除不相關標記成功", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    HtmlDocument hd = webBrowser1.Document;

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        //上面程式"btnDeleteNode_Click"執行成功後即觸發此事件
        private void DomTreeView3_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //string _hightlightStyle = "border:2px #FF0000 solid", _lastNodeStyle = "";
            //HtmlElement _lastNode = null;
            //去除上次選取Node的Style屬性
            if (_lastNode != null && !string.IsNullOrEmpty(_lastNode.Style))
                _lastNode.Style = _lastNode.Style = _lastNodeStyle;
            //取得這次選取的HtmlElement
            HtmlElement currentNode = e.Node.Tag as HtmlElement;
            if (currentNode != null)
            {
                _lastNodeStyle = currentNode.Style;
                //設定突顯顏色
                currentNode.Style += _hightlightStyle;
                _lastNode = currentNode;
            }
        }
        #endregion

        #region btndataArea_Click
        /// <summary>
        /// btndataArea_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btndataArea_Click(object sender, EventArgs e)
        {
            try
            {
                if (DOMTreeView2.Nodes.Count == 0)
                {
                    MessageBox.Show("尚未讀出檔案", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    Global com = new Global();
                    com.MarkDeleteDataArea(DOMTreeView3, DOMTreeView3.Nodes[0]);
                    //com.DeleteDataArea(DOMTreeView4, DOMTreeView4.Nodes[0]);
                    //com.DeleteDataArea1(DOMTreeView4, DOMTreeView4.Nodes[0]);
                    com = null;
                    toolStripStatusLabel1.Text = "刪除完成";
                    Thread.Sleep(1000);
                    MessageBox.Show("去除不相關資訊區塊", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    HtmlDocument hd = webBrowser1.Document;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region btnSearch2_Click
        /// <summary>
        /// DOM_Tree 內容關鍵字搜尋
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch2_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtKeyword2.Text.ToString().Trim() == "")
                {
                    MessageBox.Show("請輸入關鍵字", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtKeyword2.Focus();
                    return;
                }
                DOMTreeView3.SelectedNode = DOMTreeView3.Nodes[0];
                while (DOMTreeView3.SelectedNode != null)
                {
                    if (!DOMTreeView3.SelectedNode.IsExpanded) DOMTreeView3.SelectedNode.Expand();

                    int search = DOMTreeView3.SelectedNode.Text.IndexOf(txtKeyword2.Text.ToString().Trim());
                    if (search != -1)
                    {
                        //	DOMTreeView.SelectedNode.Checked=true;
                        DOMTreeView3.SelectedNode.BackColor = Color.Yellow;
                        break;
                    }
                    else
                    {
                        //	DOMTreeView.SelectedNode.Checked=false;
                        DOMTreeView3.SelectedNode.BackColor = Color.White;
                    }

                    DOMTreeView3.SelectedNode = DOMTreeView3.SelectedNode.NextVisibleNode;
                    DOMTreeView3.Select();
                    
                }
                   MessageBox.Show("關鍵字搜尋成功", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string strMessage = String.Format("關鍵字搜尋失敗 \n{0}", ex.ToString());
                MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                toolStripStatusLabel1.Text = "搜尋失敗";
            }
        }
        #endregion

        #region 網頁網址之載入
        private void rdoyahoonews_CheckedChanged(object sender, EventArgs e)
        {
            this.txtnewsAddress.Text = this.rdoyahoonews.Tag.ToString();
            cboNextPage.SelectedIndex = 1;
        }
        /*
        private void rdomsnnews_CheckedChanged(object sender, EventArgs e)
        {
            this.txtnewsAddress .Text  = this.rdomsnnews.Tag.ToString();
        }
        */
        private void rdoGoogle_CheckedChanged(object sender, EventArgs e)
        {
            this.txtnewsAddress.Text  = this.rdoGoogle.Tag.ToString();
            cboNextPage.SelectedIndex = 0;
        }
        #endregion

        #region btn_InfoBlock_Click -- 載入DOM_Tree
        /// <summary>
        /// 載入DOM_Tree
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_LoadDOMTree_Click(object sender, EventArgs e)
        {
            try
            {
                if (treeView1.Nodes == null)
                {
                    MessageBox.Show("請載入TreeNode");
                }
                else
                {
                    this.DOMTreeView3.CollapseAll();    //閉合(縮合)DOM_Tree以確保下行指令能抓到根節點
                    TreeNode node = (TreeNode)this.DOMTreeView3.TopNode.Clone();
                    this.treeView1.Nodes.Clear();
                    this.treeView1.Nodes.Add(node);
                    this.treeView1.ExpandAll();
                    sa.Clear();
                    sb.Clear();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region btn_InfoBlock_Click -- 找尋DOM_Tree 內容區塊
        /// <summary>
        /// 找尋DOM_Tree 內容區塊
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        StringBuilder sa = new StringBuilder(); //文字容器
        StringBuilder sb = new StringBuilder();
        private void btn_InfoBlock_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            //richTextBox2.Clear();
            CallRecursive(this.treeView1); //遞迴
        }

        private void CallRecursive(TreeView treeView)           // Call the procedure using the TreeView.
        {
            TreeNodeCollection nodes = treeView.Nodes;  //Nodes 屬性可以用於取得包含樹狀結構中所有根節點的 TreeNodeCollection 物件
            foreach (TreeNode n in nodes)
            {
                PrintRecursive(n);      // Print each node recursively.
            }
        }
        /*
        /// <summary>
        /// 遞迴依序尋找節點並列印出路徑
        /// </summary>
        /// <param name="treeNode"></param>
        private void PrintRecursive(TreeNode treeNode)
        {
            if (treeNode.Tag is mshtml.IHTMLDOMNode)    //如果該節點非葉節點
            {
                if (treeNode.Nodes.Count > 1)   //該節點的子節點(分支數)大於1
                {
                    sa.Append(string.Format("{0}~︵{1}︶,", treeNode.Text, treeNode.Nodes.Count));  //取得該節點的Tag_name及該節點分支數
                    sb.Append(string.Format("{0}~︵{1}︶,", treeNode.Text, treeNode.Nodes.Count));  //取得該節點的Tag_name及該節點分支數
                }
                else
                {
                    sa.Append(string.Format("{0}", treeNode.Text));    //取得該節點的Tag_name
                    sb.Append(string.Format("{0}", treeNode.Text));    //取得該節點的Tag_name
                }
            }
            else          //該節點為葉節點
            {
                sa.Append(string.Format(",{0} \n", treeNode.Text));  //每一分支結尾,取得該葉節點上的文字
                sb.Append(string.Format(",{0} \n" , treeNode.Text));  //每一分支結尾,取得該葉節點上的文字
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
        */
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
                 */ 

            // 遞迴,依序尋找節點
            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive(tn);
            }

            //結果輸出
            richTextBox1.Text = sa.ToString();
            sb.Replace("~", "\n");
            //richTextBox2.Text = sb.ToString();
        }
        /// <summary>
        /// 區塊判斷_Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param> 
        private void btn_GroupInfo_Click(object sender, EventArgs e)
        {
            String[] sections = null;       //儲存每一行字串之分段
            String[] sections_next = null;  //儲存每一行字串之分段
            String[] sections_next1 = null;  //儲存每一行字串之分段
            int words = 0;      //存每行字串後之文字的字數
            int arrayleath =0;
            int dup=0;
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
                    for(int i = 1; i <= dup; i++)
                    {
                        arrayleath = lineNo + i + 1;
                        sections_next = lines[lineNo + i].Split(new char[] { '↖' });
                        
                        if (arrayleath < lines.Length)
                        {
                            sections_next1 = lines[lineNo + i + 1].Split(new char[] { '↖' });
                        }
                         
                        //sections_next1 = lines[lineNo + i + 1].Split(new char[] { '/' });
                       //if (sections_next[0].IndexOf(sections [1]) >0 )
                        if (sections[1].IndexOf(sections_next[0]) !=-1)
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
                else if( Strnew_split .Length == 2)
                {
                    Strnew = Strnew_split[1];
                    if (Strnew.IndexOf("<") != -1)
                    Strnew = Strnew.Replace(Strnew, String.Empty);
                }
                else if (Strnew_split.Length == 1)
                {
                    Strnew = Strnew_split[0];
                    if(Strnew.IndexOf ("<") != -1 )
                    Strnew  = Strnew.Replace(Strnew,String.Empty);
                }
                //Strnew = Strnew.Remove(Strnew.IndexOf('<'), Strnew.IndexOf('>')+2);
                sa.AppendLine(Strnew);
                startleath++;
            }
            //richTextBox3.Text = sa.ToString();

        }
        #endregion

        #region 往上尋找節點並列印出路徑(HtmlElement)
        public string GetElementPath(HtmlElement element, string baseResult = "")
        {
                //往上尋找
                if (element.Parent != null)
                {
                    HtmlElement parentElement = element.Parent;
                    int index = 0;
                    foreach (HtmlElement childElement in parentElement.Children)
                    {
                        //如果是同一個物件，則傳回Index和TagName
                        if (childElement.Equals(element))
                        {
                            return GetElementPath(element.Parent, string.Format("元素({0})<{1}>", index, element.TagName) + (baseResult == "" ? "" : "→") + baseResult);
                        }
                        //反之，繼續找
                        else
                        {
                            index += 1;
                        }
                    }
                    //未預期狀況
                    return GetElementPath(element.Parent, baseResult);
                }
                //直到上層沒有節點為止
                else
                {
                    return baseResult;
                }
            
        }
        #endregion

        #region 播放新聞
        SpeechSynthesizer synh = new SpeechSynthesizer(); //speech物件宣告 Text to Voice
        bool blclick= false ;//布林變數 按鈕邏輯
        private void btn_speech_Click(object sender, EventArgs e)
        {           
                synh.SpeakAsyncCancelAll();
        }
        #endregion

        #region ResetdataGridVeiw
        private void ResetdataGridVeiw()
        {
            dataNews.AllowUserToAddRows = false;      //指出是否已為使用者顯示可加入資料列的選項。
            dataNews.AllowUserToDeleteRows = true;    //指出是否允許使用者從 DataGridView 中刪除資料列。
            dataNews.AllowUserToResizeColumns = true; //指出使用者是否可以調整資料行的大小。
            dataNews.AllowUserToResizeRows = true;    //指出使用者是否可以調整資料列的大小。

            dataNews.TopLeftHeaderCell.Value = "reset"; //取得或設定位於 DataGridView 控制項左上角的標題儲存格。
            dataNews.TopLeftHeaderCell.Style.ForeColor = System.Drawing.Color.Blue;

            dataNews.AutoGenerateColumns = true;      //指出當設定了 DataSource 或 DataMember 屬性時，是否會自動建立資料行。
            dataNews.ReadOnly = true;                 //指出使用者是否可以編輯 DataGridView 控制項的儲存格。
            dataNews.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;

            dataNews.BorderStyle = BorderStyle.Fixed3D;
            // Put the cells in edit mode when user enters them.
            //dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataNews.DefaultCellStyle.WrapMode = DataGridViewTriState.True;  //取得或設定 DataGridView 中的儲存格所要套用的預設儲存格樣式
            //dataGridView1.SortedColumn.SortMode = DataGridViewColumnSortMode.Programmatic;
        }
        #endregion

        #region 擷取新聞
        private void btnGet_Click(object sender, EventArgs e)
        {
            intTotCount = 1; intPageCount = 1; 
            try
            {
                if (DOMTreeView1.Nodes.Count == 0)
                {
                    MessageBox.Show("請先載入網頁", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    dataSet1.Tables.Clear();
                    dataNews.DataSource = null;
                    dataNews.Refresh();
                    dt.Clear();
                    if (rdoyahoonews.Checked)
                    {
                        YahooRecursive(dt);
                    }
                    if (rdoGoogle.Checked)
                    {
                        GoogleRecursive1(dt);
                    }

                }
            }
            catch (Exception ex)
            {
                String strMessage = String.Format("取得資料失敗 \n{0}", ex.ToString());
                MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #region 擷取 Yahoo新聞網部分
        private void YahooRecursive(DataTable dt) //擷取yahoo部分
        {
            int intAddCount = 0;
            DataRow dr = null;
            String newsHref = null;
            String newsTitle = null;
            String newsDate = null;
            String newsSource = null;
            string strNextPage = "";
            string strStart = "";
            //int intLevelCount = 0;
            if (DOMTreeView1.Nodes == null)
            {
                return;
            }
            else
            {
                DOMTreeView1.SelectedNode = DOMTreeView1.Nodes[0].Nodes[1];
                DOMTreeView1.Visible = false;
                DOMTreeView1.ExpandAll();
            }

            while (DOMTreeView1.SelectedNode != null)
            {

                while (intTotCount <= Convert.ToInt16(intCount))
                {
                    if (DOMTreeView1.SelectedNode.Tag is HtmlElement)
                    {
                        HtmlElement el = DOMTreeView1.SelectedNode.Tag as HtmlElement;
                        if (DOMTreeView1.SelectedNode == null) break;
                        intAddCount = Convert.ToInt16(cboNextPage.Text.Substring(cboNextPage.Text.Length - 1, 1));
                        strNextPage = cboNextPage.Text.Substring(0, cboNextPage.Text.Length - 1);
                        strStart = (intPageCount * Convert.ToInt16("10") + intAddCount).ToString();
                        if (el.TagName == "A" && el.OuterHtml.IndexOf("fz-m") != -1)
                        {
                            newsHref = el.GetAttribute("HREF");
                            newsTitle = el.InnerText;

                        }
                        if (el.TagName == "SPAN" && el.OuterHtml.IndexOf("fc-12th") > 1)
                        {
                            newsSource = el.InnerText;

                        }
                        if (el.TagName == "SPAN" && el.OuterHtml.IndexOf("fc-2nd ml-5") > 1)
                        {
                            newsDate = el.InnerText;
                            dr = dt.NewRow();
                            dr[0] = intTotCount;
                            dr[1] = newsHref; //網址
                            dr[2] = newsTitle; //標題
                            dr[3] = newsSource; //來源
                            dr[4] = newsDate;  //日期 
                            dt.Rows.Add(dr);
                            if (intTotCount % Convert.ToInt16("10") == 0)
                            {
                                intTotCount++;
                                bolExecuteFlag = true;
                                webBrowser1.Navigate(txtnewsAddress.Text + System.Uri.EscapeUriString(txtKeyword1.Text) + strNextPage + strStart);
                                intPageCount++;
                                return;
                            }
                            intTotCount++;
                        }
                        DOMTreeView1.SelectedNode = DOMTreeView1.SelectedNode.NextVisibleNode;
                        DOMTreeView1.Select();
                    }
                    else
                    {
                        DOMTreeView1.SelectedNode = DOMTreeView1.SelectedNode.NextVisibleNode;
                        DOMTreeView1.Select();
                    }
                }
                break;
            }
            DOMTreeView1.Visible = true;
            bolExecuteFlag = false;
            dataSet1.Tables.Add(dt);
            dataNews.DataSource = dataSet1.Tables[0];
        }
        #endregion
        /*
        #region 擷取 Google新聞網部分
        private void GoogleRecursive(TreeNode treeNode, DataTable dt) //擷取google部分
        {
            String Tagid = null;
            if (treeNode.Tag is HtmlElement)
            {

                HtmlElement em = treeNode.Tag as HtmlElement;
                if (em.TagName == "DIV")
                {
                    Tagid = em.GetAttribute("id");
                    if (Tagid == "ires")
                    {

                        //CutArea(em, node);
                        GoogleCutArea(em, dt);
                    }
                }
            }
            // 遞迴,依序尋找節點
            foreach (TreeNode tn in treeNode.Nodes)
            {
                GoogleRecursive(tn, dt);
            }
        }
        private void GoogleCutArea(HtmlElement el, DataTable dt)
        {
            try
            {
                int leath = 0;
                int intAddCount = 0;
                DataRow dr = null;
                String newsHref = null;
                String newsTitle = null;
                String newsDate = null;
                String newsSource = null;
                String[] newsSourcespiliter = null;
                foreach (HtmlElement em in el.All)
                {
                    em.SetAttribute("target", "_self");
                    if (em.TagName == "A" && em.InnerText != null)
                    {
                       
                        newsHref = em.GetAttribute("HREF");
                        newsTitle = em.InnerText;
                    }

                    if (em.TagName == "DIV" && em.InnerText !=null)
                    {
                        leath = em.InnerText.Length;
                        if (leath < 18)
                        {
                            if (em.InnerText.IndexOf('-') != -1)
                            {
                                newsSourcespiliter = em.InnerText.Split(new char[] { '-' });
                                newsSource = newsSourcespiliter[0].Trim();
                                newsDate = newsSourcespiliter[1].Trim();
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }
                        dr = dt.NewRow();
                        dr[0] = intAddCount;
                        dr[1] = newsHref; //網址
                        dr[2] = newsTitle; //標題
                        dr[3] = newsSource; //來源
                        dr[4] = newsDate;  //日期
                        dt.Rows.Add(dr);
                        intAddCount++;
                    }
                }
                dataSet1.Tables.Add(dt);
                dataNews.DataSource = dataSet1.Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
        #endregion
        */
        private void GoogleRecursive1(DataTable dt) //擷取Google部分
        {
            int intAddCount = 0;
            DataRow dr = null;
            String newsHref = null;
            String newsTitle = null;
            String newsDate = null;
            String newsSource = null;
            string strNextPage = "";
            string strStart = "";
            //int intLevelCount = 0;
            if (DOMTreeView1.Nodes == null)
            {
                return;
            }
            else
            {
                DOMTreeView1.SelectedNode = DOMTreeView1.Nodes[0].Nodes[1];
                DOMTreeView1.Visible = false;
                DOMTreeView1.ExpandAll();
            }

            while (DOMTreeView1.SelectedNode != null)
            {

                while (intTotCount <= Convert.ToInt16(intCount))
                {
                    if (DOMTreeView1.SelectedNode.Tag is HtmlElement)
                    {
                        HtmlElement el = DOMTreeView1.SelectedNode.Tag as HtmlElement;
                        if (DOMTreeView1.SelectedNode == null) break;
                        intAddCount = Convert.ToInt16(cboNextPage.Text.Substring(cboNextPage.Text.Length - 1, 1));
                        strNextPage = cboNextPage.Text.Substring(0, cboNextPage.Text.Length - 1);
                        strStart = (intPageCount * Convert.ToInt16("10") + intAddCount).ToString();
                        if (el.TagName == "A" && el.OuterHtml.IndexOf("l _HId") > 0)
                        {
                            newsHref = el.GetAttribute("HREF");
                            newsTitle = el.InnerText;
                        }
                        if (el.TagName == "SPAN" && el.OuterHtml.IndexOf("_tQb _IId") != -1)
                        {
                            newsSource = el.InnerText;
                        }
                        if (el.TagName == "SPAN" && el.OuterHtml.IndexOf("f nsa _uQb") != -1)
                        {
                            newsDate = el.InnerText;
                            dr = dt.NewRow();
                            dr[0] = intTotCount;
                            dr[1] = newsHref; //網址
                            dr[2] = newsTitle; //標題
                            dr[3] = newsSource; //來源
                            dr[4] = newsDate;  //日期 
                            dt.Rows.Add(dr);
                            if (intTotCount % Convert.ToInt16("10") == 0)
                            {
                                intTotCount++;
                                bolExecuteFlag = true;
                                webBrowser1.Navigate(txtnewsAddress.Text + System.Uri.EscapeUriString(txtKeyword1.Text) + strNextPage + strStart);
                                intPageCount++;
                                return;
                            }
                            intTotCount++;
                        }

                        DOMTreeView1.SelectedNode = DOMTreeView1.SelectedNode.NextVisibleNode;
                        DOMTreeView1.Select();
                    }
                    else
                    {
                        DOMTreeView1.SelectedNode = DOMTreeView1.SelectedNode.NextVisibleNode;
                        DOMTreeView1.Select();
                    }
                }
                break;
            }
            DOMTreeView1.Visible = true;
            bolExecuteFlag = false;
            dataSet1.Tables.Add(dt);
            dataNews.DataSource = dataSet1.Tables[0];
        }
        #endregion
       
        #region DOM Tree展開
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                DOMTreeView1.ExpandAll();
            }
            else
            {
                DOMTreeView1.CollapseAll();
            }
        }
        #endregion

        #region 儲存格式
        /// <summary>
        /// 儲存格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            DataRow dr = null;
            int i = 0;
            frmOpenWebBrowser frmOpenWebBrowser = null;
            string MainDir= null;
            bool IsSuccess = false;
            if (dataSet1.Tables.Count == 0)
            {
                strMessage = String.Format("無網頁資料，請先處理【取得資料】");
                MessageBox.Show(strMessage, "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                if (txtKeyword1.Text.Equals("") || txtKeyword1.Text.Equals("請輸入檔名"))
                {
                    strMessage = String.Format("請輸入儲存檔名");
                    MessageBox.Show(strMessage, "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtKeyword1.Focus();
                    return;
                }
                if (rdoyahoonews.Checked)
                {
                    MainDir = strCurrentPath + rdoyahoonews.Text + txtKeyword1.Text + intCount +"篇";
                    strSelectDir1 = MainDir;
                }
                if (rdoGoogle.Checked)
                {
                    MainDir = strCurrentPath + rdoGoogle.Text + txtKeyword1.Text + intCount + "篇";
                    strSelectDir1 = MainDir;
                }
                if (!Directory.Exists(MainDir))
                {
                    //建立目錄
                    Directory.CreateDirectory(MainDir);
                }
                else
                {
                    DirectoryInfo Dir = new DirectoryInfo(MainDir);
                    FileInfo[] totFile = Dir.GetFiles();
                    if (totFile.Length != 0)
                    {
                        foreach (FileInfo subFile in totFile)
                        {
                            if (subFile.Name == "data.txt" || subFile.Extension.Equals(".txt"))
                            {
                                strMessage = String.Format("儲存{0}檔案已存在，是否要覆蓋？", txtKeyword1.Text);
                                DialogResult DlogRut = MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (DlogRut == DialogResult.No) return;

                                //刪除目錄
                                Dir.Delete(true);

                                //建立目錄
                                Directory.CreateDirectory(MainDir);

                                break;
                            }
                        }
                    }
                    dt = dataSet1.Tables[0];
                    Encoding Ecd = null;

                    if (cboSaveFormat.SelectedIndex == 0)
                    {
                        Ecd = Encoding.UTF8;
                    }
                    else if (cboSaveFormat.SelectedIndex == 1)
                    {
                        Ecd = Encoding.Default;
                    }
                    else if (cboSaveFormat.SelectedIndex == 2)
                    {
                        Ecd = Encoding.Unicode;
                    }
                    StreamWriter sw = new StreamWriter(MainDir + @"\data.txt", true, Ecd);
                    foreach (DataGridViewRow dr1 in dataNews.Rows)
                    {
                        //讀出新聞起始時間
                        DateTime dts = DateTime.Now;

                        string strSeqNo = dr1.Cells[0].Value.ToString();
                        string strHttp = dr1.Cells[1].Value.ToString();
                        string strCaption = dr1.Cells[2].Value.ToString();
                        string strFrom = dr1.Cells[3].Value.ToString();
                        string strDate = dr1.Cells[4].Value.ToString();
                        string strFileName = DeleteErrName(strSeqNo + "." + strCaption);
                        dr = dt.Rows[i];
                        string strContect = "";
                        try
                        {
                            //宣告表單
                            frmOpenWebBrowser = frmOpenWebBrowser.Instance;
                            frmOpenWebBrowser = new frmOpenWebBrowser(strHttp);
                            frmOpenWebBrowser.Text = frmOpenWebBrowser.Text + "-" + strFileName;
                            frmOpenWebBrowser.Minutes = "10";
                            frmOpenWebBrowser.Seconds = "30";
                            frmOpenWebBrowser.Contect = "";
                            frmOpenWebBrowser.ShowDialog(this);
                        }
                        catch (InvalidOperationException ex)
                        {
                            frmOpenWebBrowser.Error = true;
                        }
                        catch (Exception ex)
                        {
                            frmOpenWebBrowser.Error = true;
                        }
                        if (frmOpenWebBrowser.Error)
                        {
                            System.TimeSpan diff = UsedTime(dts);

                            dr[5] = diff.Minutes.ToString() + "分" + diff.Seconds.ToString() + "秒" + "\n異常";

                            //儲存明細新聞內容-異常
                            StreamWriter sw1 = new StreamWriter(MainDir + @"\" + strFileName + "_異常Err.txt", true, Ecd);
                            sw1.Close();

                            //顯示異常顏色
                            dr1.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
                        }
                        else
                        {
                            while (!frmOpenWebBrowser.Busy)
                            {
                                strContect = frmOpenWebBrowser.Contect;

                                System.TimeSpan diff = UsedTime(dts);

                                if (strContect.Equals(""))
                                {
                                    dr[5] = diff.Minutes.ToString() + "分" + diff.Seconds.ToString() + "秒" + "\n逾時";

                                    //儲存明細新聞內容-逾時
                                    StreamWriter sw1 = new StreamWriter(MainDir + @"\" + strFileName + "_逾時Err.txt", true, Ecd);
                                    sw1.Close();

                                    //顯示逾時顏色
                                    dr1.DefaultCellStyle.BackColor = System.Drawing.Color.Red;
                                }
                                else
                                {
                                    dr[5] = diff.Minutes.ToString() + "分" + diff.Seconds.ToString() + "秒" + "\n完成";// strContect;

                                    string strKey = strFileName + "\r\n" + strHttp + "\r\n" + strFrom + "\r\n" + strDate + "\r\n" + strContect + "\r\n";

                                    sw.WriteLine(strKey);

                                    StreamWriter sw1 = new StreamWriter(MainDir + @"\" + strFileName + ".txt", true, Ecd);
                                    sw1.WriteLine(dr[2]);//寫入個別檔標題
                                    sw1.WriteLine(dr[4]);//寫入個別檔日期
                                    sw1.WriteLine(strContect + "\n\n");//寫入個別檔內容根據句號分段落
                                    sw1.Close();
                                    IsSuccess = true;
                                }
                                frmOpenWebBrowser.Busy = true;
                            }
                        }
                        i++;
                        frmOpenWebBrowser.Close();
                    }
                    sw.Close();

                    if (IsSuccess)
                    {
                        strMessage = String.Format("儲存{0}成功 \n路徑{1}",txtKeyword1.Text ,MainDir + @"\data.txt");
                    }
                    else
                    {
                        strMessage = String.Format("查無符合{0}內容!", txtKeyword1.Text);
                    }

                    MessageBox.Show(strMessage, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                  
                }
            }     
        }
        #endregion

        #region DeleteErrFile
        /// <summary>
        /// 將不合法的命名去除
        /// </summary>
        /// <param name="strFileName"></param>
        /// <returns></returns>
        private string DeleteErrName(string strFileName)
        {
            strFileName = strFileName.Replace("\\", "＼");
            strFileName = strFileName.Replace("/", "／");
            strFileName = strFileName.Replace(":", "：");
            strFileName = strFileName.Replace("*", "﹡");
            strFileName = strFileName.Replace("?", "？");
            strFileName = strFileName.Replace("\"", "〞");
            strFileName = strFileName.Replace("<", "〈");
            strFileName = strFileName.Replace(">", "〉");
            strFileName = strFileName.Replace("|", "│");

            return (strFileName);
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
            Uri myUri = new Uri(strUrl);
            System.Net.HttpWebRequest myWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(myUri);
            myWebRequest.Timeout = 12000;
            System.Net.HttpWebResponse myWebResponse = (System.Net.HttpWebResponse)myWebRequest.GetResponse();

            if (myWebResponse.StatusCode == System.Net.HttpStatusCode.OK)
                return true;
            else
                return false;
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

        #region dataGridView1_DoubleClick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            DataRow dr1 = dt.Rows[dataNews.CurrentRow.Index];

            string strSeqNo = dr1[0].ToString();
            string strHttp = dr1[1].ToString();
            string strCaption = dr1[2].ToString();
            string strFrom = dr1[3].ToString();
            string strDate = dr1[4].ToString();
            string strFileName = DeleteErrName(strSeqNo + "." + strCaption);
            string MainDir = null;
            string line = null;
            if (rdoyahoonews.Checked)
            {
                MainDir = strCurrentPath + rdoyahoonews.Text + txtKeyword1.Text + intCount + "篇";
            }
            if (rdoGoogle .Checked)
            {
                MainDir = strCurrentPath + rdoGoogle.Text + txtKeyword1.Text + intCount + "篇";
            }
            if (Directory.Exists(MainDir))
            {
                DirectoryInfo Dir = new DirectoryInfo(MainDir);
                FileInfo[] totFile = Dir.GetFiles(strFileName + ".txt");

                if (totFile.Length != 0)
                {
                    foreach (FileInfo subFile in totFile)
                    {
                        StringBuilder SB = new StringBuilder();
                        StreamReader sr = new StreamReader(subFile.FullName, Encoding.Default);
                        while (sr.EndOfStream != true)
                        {
                            line = sr.ReadLine().Trim();
                            if (line.StartsWith("↖"))
                            {
                                line = line.Trim("↖".ToCharArray());
                            }
                            SB.AppendLine(line);
                        }
                        
                        //string strContect = sr.ReadToEnd();
                        synh.SpeakAsync(SB.ToString ());
                        webBrowser1.Navigate(strHttp);
                        DialogResult DlogRut = MessageBox.Show(SB.ToString(), "新聞內容-" + strFileName, MessageBoxButtons.OK);
                        if (DlogRut == DialogResult.OK)
                        {
                            synh.SpeakAsyncCancelAll();
                           
                        }
                        sr.Close();
                    }

                }
                else
                {
                    MessageBox.Show("無此新聞內容", "新聞內容-" + strFileName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        #endregion

        #region 播放新聞
        private void btn_speak_Click(object sender, EventArgs e)
        {
            string patternjapen = @"\b(\p{IsHiragana}+(\s)?)+\p{L}"; // 判斷日文
            string pattenEnglish = @"\b[A-Z]\w*\b"; //判斷英文
            string input = richTextBox1.Text;
            if (Regex.IsMatch(input, patternjapen))
            {
                synh.SelectVoice("Microsoft Haruka Desktop");
            }
            else if (Regex.IsMatch(input, pattenEnglish))
            {
                synh.SelectVoice("Microsoft Zira Desktop");
            }
            else
            {
                synh.SelectVoice("Microsoft Hanhan Desktop");
            }
            if (blclick)
            {
                synh.SpeakAsyncCancelAll();
                btn_speak.Text = "播放新聞";
            }
            else
            {
                btn_speak.Text = "停止播放";
                synh.SpeakAsync(richTextBox1.Text);
            }
            blclick = !blclick;
        }
        #endregion

        #region 語音輸入
        String s1;
        private void btnSpeakOutput_Click(object sender, EventArgs e)
        {
            SpeechRecognitionEngine SRE = new SpeechRecognitionEngine();
            SRE.SetInputToDefaultAudioDevice();
            label4.Text = "請說話";
            GrammarBuilder GB = new GrammarBuilder();
            //GB.Append("我要查");
            GB.Append(new Choices(new string[] { "銅葉綠素", "華夏技術學院","馬來西亞航空","服貿" }));
            Grammar G = new Grammar(GB);
            G.Name = "main command grammar";
            SRE.LoadGrammar(G);
            G.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(G_SpeechRecognized);
            SRE.RecognizeAsync(RecognizeMode.Multiple);
        }
        void G_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            s1 = e.Result.Text;
            switch (s1)
            {
                case "銅葉綠素":
                    txtKeyword1.Text = "銅葉綠素";
                    label8.Text = "銅葉綠素";
                    break;
                case "華夏技術學院":
                    txtKeyword1.Text = "華夏技術學院";
                    label8.Text = "華夏技術學院";
                    break;
                case "馬來西亞航空":
                    txtKeyword1.Text = "馬來西亞航空";
                    label8.Text = "馬來西亞航空";
                    break;
                case "服貿":
                    txtKeyword1.Text = "服貿";
                    label8.Text = "服貿";
                    break;
            }
        }
        #endregion

        #region 區塊判斷2 
        private void btn_InfoBlock2_Click(object sender, EventArgs e)
        {
            weightlist.Clear();
            txtlist.Clear();
            Minweighnum = 0;
            txtleath2 = 0;
            leath = 0;
            size = 0;
            PrintRecursive2(treeView1.Nodes[0]);
        }
        int txtleath2 = 0;
        int Minweighnum = 0;
        ArrayList weightlist = new ArrayList();
        //List<string> weightlist = new List<string>();
        private void PrintRecursive2(TreeNode treeNode)
        {
            int weightnum = 0;
            /*
            if (treeNode.NextVisibleNode.Text.IndexOf("高雄市左營區") != -1)
            {
            }
             */
            if (treeNode.NextVisibleNode != null)
            {
                if (treeNode.NextVisibleNode.Nodes.Count == 0)//如果是葉節點
                {
                    weightnum = TreenodeGroup(treeNode, weightnum, Minweighnum);
                    if (weightnum > Minweighnum &&leath >txtleath2  )
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
        public struct TreeNodeData
        {
            public int weight;
            public int leath;
            public int size;
            public int top;
        }
       /// <summary>
       /// 計算權重值
       /// </summary>
      /// <param name="node"></param>
      /// <param name="weight"></param>
       /// <returns></returns>
        ArrayList txtlist = new ArrayList();
        List<TreeNodeData > tnd = new List<TreeNodeData> ();
        int leath = 0;
        int size = 0;
        int size2 = 0;
        public int TreenodeGroup(TreeNode node, int weight = 0, int Minweigh = 0, String txtnode = null)
        {
            //往上尋找
            if (node.Parent != null)
            {
                string width = null;
                string height = null;
                string Y = null;
                mshtml.IHTMLDOMNode domnode = null;
                int nodeCount = node.Parent.Nodes.Count;
                TreeNode parentnode = node.Parent;
                int index = node.Index;
                int txtleath = 0;
                RichTextBox rtb = new RichTextBox ();
                IHTMLElement IHE = parentnode.Tag as IHTMLElement;
                string node2 = IHE.outerHTML;
                string pattern = @"
(?:<)(?<Tag>[^\s/>]+)       # Extract the tag name.
(?![/>])                    # Stop if /> is found
                     # -- Extract Attributes Key Value Pairs  --
 
((?:\s+)             # One to many spaces start the attribute
 (?<Key>[^=]+)       # Name/key of the attribute
 (?:=)               # Equals sign needs to be matched, but not captured.
 
(?([\x22\x27])              # If quotes are found
  (?:[\x22\x27])
  (?<Value>[^\x22\x27]+)    # Place the value into named Capture
  (?:[\x22\x27])
 |                          # Else no quotes
   (?<Value>[^\s/>]*)       # Place the value into named Capture
 )
)+                  # -- One to many attributes found!";
                var attributes = (from Match mt in Regex.Matches(node2, pattern, RegexOptions.IgnorePatternWhitespace)
                                  select new
                                  {
                                      Name = mt.Groups["Tag"],
                                      Attrs = (from cpKey in mt.Groups["Key"].Captures.Cast<Capture>().Select((a, i) => new { a.Value, i })
                                               join cpValue in mt.Groups["Value"].Captures.Cast<Capture>().Select((b, i) => new { b.Value, i }) on cpKey.i equals cpValue.i
                                               select new KeyValuePair<string, string>(cpKey.Value, cpValue.Value)).ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
                                  }).First().Attrs;
                foreach (KeyValuePair<string, string> kvp in attributes)
                {
                    //Console.WriteLine("Key {0,15}    Value: {1}", kvp.Key, kvp.Value);
                    switch (kvp.Key)
                    {
                        case "top":
                            Y = kvp.Value;
                            break;
                        case "height":
                            height = kvp.Value;
                            break;
                        case "width":
                            width =kvp.Value;
                            break;

                    }
                }                              
                foreach ( TreeNode childnode in parentnode.Nodes)
                {
                    if (childnode.Parent.Equals(node.Parent)) //如果 兩個node路徑相同時回傳 權重值 
                    {
                            if (childnode.Tag is mshtml.IHTMLDOMNode) // 如果此節點是mshtml
                            {
                               

                              if(childnode .NextVisibleNode !=null)
                               domnode = childnode.NextVisibleNode.Tag as mshtml.IHTMLDOMNode;
                               else 
                                domnode = childnode.Tag as mshtml.IHTMLDOMNode;
                              


                              
                                if (childnode.Nodes.Count <= 1)
                                {
                                    
                                    if (domnode.nodeName == "#text") //如果是葉節點(text node)
                                    {
                                        if (childnode.NextVisibleNode!=null)
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
                                        if (childnode.NextVisibleNode != null)
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
                //leath = txtleath; 
                /*    
                else
                {
                    weight = txtleath;
                }
                  */   
                     
                if (weight > Minweigh &&  txtleath >leath   )
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
        String max = null;
        public void GetInfoGroup()
        {
            for (int i = 0; i < weightlist .Count; i++)
            {
                for (int j = i + 1; j < weightlist.Count; j++)
                {
                    if (Convert .ToInt32 (weightlist[i]) > Convert .ToInt32 (weightlist[j]))
                    {
                        int temp = Convert .ToInt32 (weightlist[i]);
                        weightlist[i] =weightlist[j];
                        weightlist[j] = temp;
                    }
                }
            }
            max =Convert.ToString(weightlist[weightlist.Count - 1]);
            PrintRecursive4(treeView1.Nodes[0]);
        }
        StringBuilder SB = new StringBuilder();
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
                                if ( tn.NextVisibleNode.Text.IndexOf(">") == -1 && domnode.nodeName == "#text")
                                    SB.AppendLine(tn.NextVisibleNode.Text);
                            }
                        }
                    }
                    else 
                    {
                        mshtml.IHTMLDOMNode domnode1 = tt.NextVisibleNode.Tag as mshtml.IHTMLDOMNode;
                        if(tt.NextVisibleNode.Text .IndexOf (">") ==-1&&domnode1.nodeName=="#text")
                            SB.AppendLine (tt.NextVisibleNode.Text); 
                    }
                }
                richTextBox1.Text = SB.ToString();
            }
            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive4(tn);
            }
        }
        private void btn_GroupInfo2_Click_1(object sender, EventArgs e)
        {
            SB.Clear();
            PrintRecursive3(treeView1.Nodes[0]);
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
        #endregion

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
                            strNode ="<"+ (string)child_node.nodeName.ToString().ToUpper().Trim()+">";

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

        #region 選擇抓取新聞篇數
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                intCount = 50;
            }
            else
            {
                intCount = 0;
            }
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                intCount = 100;
            }
            else
            {
                intCount = 0;
            }
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                intCount = 10;
            }
            else
            {
                intCount = 0;
            }
        }
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                intCount = 20;
            }
            else
            {
                intCount = 0;
            }
        }
        #endregion

        #region 檔案合併
        //------------------------------------------------------------------------------------------------------------------------------------------------
        //----tabPage1----檔案合併----tabPage1----檔案合併----tabPage1----檔案合併----tabPage1----檔案合併----tabPage1----檔案合併----tabPage1----檔案合併
        //------------------------------------------------------------------------------------------------------------------------------------------------
        string dirPath1 = null;
        private void btn_openFiles_Click(object sender, EventArgs e)
        {
            if (txt_p1_openFileName.Text != "")
                openFileDialog1.InitialDirectory = txt_p1_openFileName.Text;
            else
                openFileDialog1.InitialDirectory = strCurrentPath;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)    //開啟Window開啟檔案管理對話視窗
            {
                string[] fileNames = openFileDialog1.FileNames;     //由檔案開啟對話盒選擇多個檔案合併
                string  dirPath = txt_p1_openFileName.Text = Path.GetDirectoryName(fileNames[0]);    //將選取的檔案路徑顯示於於txt_openFileInfo(textBox)視窗
                //--- 設定合併檔儲存路徑與檔名 ---
                dirPath = dirPath.Substring(0, (dirPath.LastIndexOf(@"\")));  //取得合併文件檔所在"父目錄路徑"
                //txt_p1_saveFileName.Text = dirPath + @"\test.txt";            //儲存合併檔檔案名稱為"test.txt"
                dirPath1 = @"C:\Users\user\Desktop\處理文件段落摘要";
                txt_p1_saveFileName.Text = dirPath1  + @"\test.txt";            //儲存合併檔檔案名稱為"test.txt"
                StringBuilder SB_headLines = new StringBuilder();    //StringBuilder物件 用於儲存文件標題檔 
                StringBuilder SB = new StringBuilder();      //StringBuilder適合未知數目之字串動態串接,此處使用string物件則無法完成此大資料檔案的讀取串接
                string line = "";
                int pargraphs = 1;               //起始"段落編號"
                int docNo = int.Parse(txt_Count.Text); //起始"文件編號",由"txt_Count"對話盒設定
                if (ckb_File_1.Checked == true)
                {
                    foreach (string fileName in fileNames)
                    {
                        if (File.Exists(fileName))      //若該文件檔案存在
                        {
                            if (fileName.IndexOf("異常") != -1 || fileName.IndexOf("逾時Err") != -1)
                            {
                                continue;
                            }
                            else
                            {
                                SB.AppendLine("<DocNo>" + docNo + "</DocNo>");  //加入"文件編號"Tag;
                            }
                            using (StreamReader myTextFileReader = new StreamReader(fileName, Encoding.Default))   //開啟文件檔案
                            {
                                // 文件檔案第一行為"新聞標題"
                                line = myTextFileReader.ReadLine().Trim();      //第一行為"新聞標題"
                                SB.AppendLine("<TITLE>" + line + "</TITLE>");   //加入"新聞標題"Tag
                                if (ckb_p1_headLine.Checked == true)            //"輸出文件標題檔"
                                    SB_headLines.AppendLine("<TITLE>" + docNo + "</TITLE>" + line);

                                // ---第四行起為"文件內文" 
                                pargraphs = 1;      //文件段落編號
                                while (myTextFileReader.EndOfStream != true)     //非資料檔案結尾
                                {
                                    if ((line = myTextFileReader.ReadLine().Trim()).Length > 1)  //非空白行再處理並儲存
                                    {
                                        if (ckb_p1_pargraphs.Checked == true)    //文件各段落加入"文件段落編號"標記
                                        {
                                            line = "<P" + pargraphs + ">" + line + "</P>";        //加入"文件段落編號"Tag
                                            pargraphs++;
                                        }

                                        SB.AppendLine(line);
                                    }
                                }
                            }

                            docNo++;
                        }
                        SB.AppendLine();        //每一文件間加一空白行
                    }
                }
                if (ckb_File_2.Checked == true)
                {
                    dirPath1 = @"C:\Users\user\Desktop\處理相似文件段落摘要";
                    txt_p1_saveFileName.Text = dirPath1 + @"\test.txt"; 
                    foreach (string fileName in fileNames)
                    {
                        if (File.Exists(fileName))      //若該文件檔案存在
                        {
                            using (StreamReader myTextFileReader = new StreamReader(fileName, Encoding.Default))   //開啟文件檔案
                            {
                                while (myTextFileReader.EndOfStream != true)     //非資料檔案結尾
                                {
                                    if ((line = myTextFileReader.ReadLine().Trim()).Length > 1)  //非空白行再處理並儲存
                                    {
                                        if (line.StartsWith("↖"))
                                            {
                                                SB.AppendLine("<DocNo>" + docNo + "</DocNo>" + line);
                                                docNo++;
                                            }
                                            /*
                                            else
                                            {
                                                SB.AppendLine(line);
                                            }
                                             */ 
                                       
                                    }
                                }
                            }

                        }
                        SB.AppendLine();        //每一文件間加一空白行
                    }
                }
                rtb_fileMerged.Text = SB.ToString();    //輸出顯示文件合併結果
                //單獨輸出"文件標題"檔案 
                if (ckb_p1_headLine.Checked == true)
                {
                    string headLineFileName = dirPath + @"\HeadLineFile.txt";     //設定"輸出文件標題檔" 檔名
                    File.WriteAllText(@headLineFileName, SB_headLines.ToString(), Encoding.Default);     //"輸出文件標題檔"
                }
            }
        }

        // --- 儲存合併檔 ---以"txt_DirPath2"所示為檔名
        private void btn_save_Click(object sender, EventArgs e)
        {
            //saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            /*
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)        //開啟Window儲存檔案管理對話視窗設定檔名
            {
                txt_p1_saveFileName.Text = saveFileDialog1.FileName; //將選取的檔案名稱顯示於於txt_saveFileInfo(textBox)視窗
                File.WriteAllText(@txt_p1_saveFileName.Text, rtb_fileMerged.Text, Encoding.Default);
                MessageBox.Show("合併文件儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
             */
            if (!Directory.Exists(dirPath1))
            {
                //建立目錄
                Directory.CreateDirectory(dirPath1);
                File.WriteAllText(@txt_p1_saveFileName.Text, rtb_fileMerged.Text, Encoding.Default);
                MessageBox.Show("檔案合併文件儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("檔案已存在,是否要覆寫檔案", "檔案儲存作業", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (dialogResult == DialogResult.OK)
                {
                    File.WriteAllText(@txt_p1_saveFileName.Text, rtb_fileMerged.Text, Encoding.Default);
                    MessageBox.Show("檔案合併文件儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                }
            }
            txt_p2_openFileName.Text = txt_p1_saveFileName.Text;
        }
        #endregion

        #region 中文斷詞
        //------------------------------------------------------------------------------------------------------------------------------------------------
        //----tabPage2----中文斷詞----tabPage2----中文斷詞----tabPage2----中文斷詞----tabPage2----中文斷詞----tabPage2----中文斷詞----tabPage2----中文斷詞
        //------------------------------------------------------------------------------------------------------------------------------------------------
        private void btn_startCKIP_Click(object sender, EventArgs e)
        {
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
                System.Diagnostics.Process.Start(@"C:\Program Files\CKIP\Autotag\WinAT.exe");
        }

        private void btn_tagClearing_Click(object sender, EventArgs e)
        {
            //openFileDialog1.InitialDirectory = dirPath1;
           // if (openFileDialog1.ShowDialog() == DialogResult.OK)   //開啟Window開啟檔案管理對話視窗
            //{
                //txt_p2_openFileName.Text = openFileDialog1.FileName; //將選取的檔案名稱顯示於於txt_openFileInfo(textBox)視窗
                //--- 設定斷詞標籤清除後之儲存路徑與檔名 ---
            txt_p2_openFileName.Text = dirPath1 + @"\test斷詞結果";
                string dirPath = txt_p2_openFileName.Text;
                dirPath = dirPath.Substring(0, (dirPath.LastIndexOf(@"\")));     //取得斷詞文件檔所在"父目錄路徑"
                txt_p2_saveFileName.Text = dirPath + @"\test.cle";               //設定斷詞標籤清除後之儲存檔案名稱為"test.cle"
                
                rtb_p2_segFile.Clear();
                rtb_p2_segFile.Text = File.ReadAllText(txt_p2_openFileName.Text, Encoding.Default);  //讀取斷詞文件檔

                string line = "", subLine = "";
                using (StreamReader myTextFileReader = new StreamReader(txt_p2_openFileName.Text, Encoding.Default))
                {
                    StringBuilder SB = new StringBuilder();            //StringBuilder適合未知數目之字串動態串接,此處使用string物件則無法完成此大資料檔案的讀取串接
                    while (myTextFileReader.EndOfStream != true)       //非資料檔案結尾
                    {
                        line = myTextFileReader.ReadLine().Trim();

                        // --- Delete all the marks that produced by CKIP program
                        if ((line.Length > 2) || (line != Environment.NewLine))  //"符號行"與"空白行"不要處理
                        {
                            line = line.Replace("　", "||");                //中文全形空白轉換成英文半形空白
                            line = line.Replace("(PERIODCATEGORY)", "");
                            line = line.Replace("(COLONCATEGORY)", "");
                            line = line.Replace("(PERIODCATEGORY)", "");
                            line = line.Replace("(PARENTHESISCATEGORY)", "");
                            line = line.Replace("(PAUSECATEGORY)", "");
                            line = line.Replace("(PAUSECATEGORY)", "");
                            line = line.Replace("(QUESTIONCATEGORY)", "");
                            line = line.Replace("(COMMACATEGORY)", "");
                            line = line.Replace("(SEMICOLONCATEGORY)", "");
                            line = line.Replace("(ETCCATEGORY)", "");
                            line = line.Replace("(EXCLANATIONCATEGORY)", "");
                            line = line.Replace("(DASHCATEGORY)", "");

                            //if (line.Contains("<DocNO>") && line.IndexOf("</DocNo>")!=-1)             //找到文件編號就分行處理
                            if (line.IndexOf("<DocNo>") != -1)
                            {
                                subLine = line.Substring(line.IndexOf("<DocNo>"), ((line.IndexOf("</DocNo>") + 8) - line.IndexOf("<DocNo>"))); //Substring(起點, 長度) 
                                line = line.Remove(line.IndexOf("<DocNo>"), ((line.IndexOf("</DocNo>") + 8) - line.IndexOf("<DocNo>")));
                                SB.AppendLine();
                                SB.AppendLine(subLine);
                            }
                            if (line.IndexOf("<lineNo>") != -1)
                            {
                                subLine = line.Substring(line.IndexOf("<lineNo>"), ((line.IndexOf("</lineNo>") + 9) - line.IndexOf("<lineNo>"))); //Substring(起點, 長度) 
                                line = line.Remove(line.IndexOf("<lineNo>"), ((line.IndexOf("</lineNo>") + 9) - line.IndexOf("<lineNo>")));
                                SB.AppendLine();
                                SB.AppendLine(subLine);
                            }
                            /*/if (line.IndexOf("<TITLE>") != -1 && line.IndexOf("</TITLE>") != -1)             //找到文件編號就分行處理
                            if (line.IndexOf("<TITLE>") != -1 )
                            {
                                string[] lines = line.Split(new string[] {"</TITLE>"}, StringSplitOptions.RemoveEmptyEntries); 
                               //subLine = line.Substring(line.IndexOf("<TITLE>"), ((line.IndexOf("</TITLE>") + 8) - line.IndexOf("<TITLE>")) ); //Substring(起點, 長度) 
                                //line = line.Remove(line.IndexOf("<TITLE>"), ((line.IndexOf("</TITLE>") + 8) - line.IndexOf("<TITLE>")) );
                                SB.AppendLine(lines[0]);
                                line = lines[1];
                            }//*/

                        }
                        if (line.StartsWith("。"))          //清除起始句點"。"
                            line = line.Remove(0, 1);

                        if (!line.EndsWith("。"))         //清除CKIP所造成之非句點"。"之分段
                        {
                          if(line.Length !=0)
                            line = line.Remove(line.Length - 1);      //去除行尾(即分行符號)
                            SB.Append(line);
                        }
                        else
                            SB.AppendLine(line);

                    }
                    rtb_p2_tagClean.Text = SB.ToString();
                }
           // }
            btn_saveFile.Enabled = true;
        }

        private void btn_p2_saveFile_Click(object sender, EventArgs e)
        {
            if (txt_p2_saveFileName.Text == "")
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)    //開啟Window儲存檔案管理對話視窗設定檔名
                {
                    txt_p2_saveFileName.Text = saveFileDialog1.FileName; //將選取的檔案名稱顯示於於txt_saveFileInfo(textBox)視窗
                    File.WriteAllText(txt_p2_saveFileName.Text, rtb_p2_tagClean.Text, Encoding.Default);
                    MessageBox.Show("文件斷詞標記儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                }
            }
            else if (File.Exists(@txt_p2_saveFileName.Text))
            {
                DialogResult dialogResult = MessageBox.Show("檔案已存在,是否要覆寫檔案", "檔案儲存作業", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (dialogResult == DialogResult.OK)
                {
                    File.WriteAllText(@txt_p2_saveFileName.Text, rtb_p2_tagClean.Text, Encoding.Default);
                     txt_p5_fileName.Text = txt_p2_saveFileName .Text;
                    MessageBox.Show("文件斷詞標記儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                }
                else
                    MessageBox.Show("文件儲存已取消", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            else
            {
                txt_p5_fileName.Text = txt_p2_saveFileName.Text;
                File.WriteAllText(txt_p2_saveFileName.Text, rtb_p2_tagClean.Text, Encoding.Default);
                MessageBox.Show("文件斷詞標記儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
        }
        #endregion


        #region 詞彙挑選
        //------------------------------------------------------------------------------------------------------------------------------------------------
        //----tabPage5----詞彙挑選----tabPage5----詞彙挑選----tabPage5----詞彙挑選----tabPage5----詞彙挑選----tabPage5----詞彙挑選----tabPage5----詞彙挑選
        //------------------------------------------------------------------------------------------------------------------------------------------------
        // -------- 每篇文件所含之"詞彙"與相對應之"tfidf"串列list ------------
        List<string> Term_list = new List<string>();         //記錄經詞彙出現次數挑選後出現在每一篇文件中的所有詞彙
        List<KeyValuePair<string, double>> DocTermtfidf_list = new List<KeyValuePair<string, double>>();    //-記錄每篇文件出現之詞彙集詞頻 -

        private void btn_featureSelec_Click(object sender, EventArgs e)
        {
            string line = "", Term = "", POS = "";      //資料列, 詞彙, 詞性標記

           // if ((openFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p5_fileName.Text = openFileDialog1.FileName) != null))   //開啟Window開啟檔案管理對話視窗
           // {
                rtb_p5_souce.Text = File.ReadAllText(@txt_p5_fileName.Text, Encoding.Default);    //顯示原始檔(未經詞彙挑選前)

                using (StreamReader myTextFileReader = new StreamReader(@txt_p5_fileName.Text, Encoding.Default))
                {
                    rtb_p5_selected.Clear();        //經詞彙挑選後
                    StringBuilder SB = new StringBuilder();            //StringBuilder適合未知數目之字串動態串接,此處使用string物件則無法完成此大資料檔案的讀取串接
                    while (myTextFileReader.EndOfStream != true)       //非資料檔案結尾
                    {
                        line = myTextFileReader.ReadLine().Trim();
                        if (line.Contains("<DocNo>"))
                        {
                            SB.Append("\n" + line + "\n");     //輸出文件編號抬頭
                            continue;
                        }
                        if (line.Contains("<lineNo>"))
                        {
                            SB.Append("\n" + line + "\n");     //輸出文件編號抬頭
                            continue;
                        }
                        if (line.Length >= 2)
                        {
                            string[] terms = line.Split(new char[] { '|', '|' }, StringSplitOptions.RemoveEmptyEntries);

                            foreach (string str in terms)
                            {
                                if (str.Contains("TITLE"))
                                {
                                    SB.Append(str.Substring(0, str.IndexOf(">") + 1) + "|");
                                    continue;
                                }
                                if (str == "，" || str == "。")
                                {
                                    SB.Append(str + "|");
                                    continue;
                                }
                                if (str.IndexOf("(") == -1 || str.IndexOf(")") == -1)       //找不到'()'，就不處理/
                                    continue;

                                Term = str.Substring(0, str.IndexOf("("));
                                POS = str.Substring(str.IndexOf("("));
                                if (POS.Length != 0)        //若有"詞性標記"存在
                                {
                                    if (((POS == "(Na)") && (ckb_Na.Checked == true))
                                          || ((POS == "(Nb)") && (ckb_Nb.Checked == true))
                                          || ((POS == "(Nc)") && (ckb_Nc.Checked == true))
                                          || ((POS == "(Nd)") && (ckb_Nd.Checked == true))
                                          || ((POS == "(VA)") && (ckb_VA.Checked == true))
                                          || ((POS == "(VB)") && (ckb_VB.Checked == true))
                                          || ((POS == "(VC)") && (ckb_VC.Checked == true))
                                          || ((POS == "(VD)") && (ckb_VD.Checked == true))
                                          || ((POS == "(N)") && (ckb_longTerm.Checked == true))
                                          || ((POS == "(Name)") && (ckb_Name.Checked == true))
                                          || ((POS == "(FW)") && (ckb_FW.Checked == true))
                                          || ((POS == "(A)") && (ckb_Adj.Checked == true)))
                                    {
                                        //--刪除一字詞
                                        if (ckb_delSingle.Checked == false)
                                            SB.Append(str + "|");
                                        else if (Term.Length >= 2)        //詞彙長度兩字以上 
                                            SB.Append(str + "|");
                                        else
                                            continue;
                                    }
                                }
                            } //文件中的每一行文字
                        }
                    }//每一文件
                    rtb_p5_selected.Text = SB.ToString();      //顯示經詞彙挑選後資料
                }
            //}
        }

        // ========================================================================================================================
        // --------------------------------------------------- 文件詞彙列表處理 ---------------------------------------------------
        // ========================================================================================================================
        int MaxTfValue = 0;
        private void btn_p5_termList_Click(object sender, EventArgs e)
        {
            string line = "", Term = "", POS = "", DocNO_tag = "";      //資料列, 詞彙, 詞性標記
            int index = 0, value = 0, pt = 0;
            StringBuilder SB5 = new StringBuilder();
            List<string> temp_list = new List<string>();
            SortedList Term_DF_slist = new SortedList();            //整個文件集之"詞彙列表" 及詞會出現次數 DF
            List<KeyValuePair<string, int>> Doctermtf_list = new List<KeyValuePair<string, int>>(); //-記錄每篇文件出現之詞彙集詞頻 -
           // if ((openFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p5_fileName.Text = openFileDialog1.FileName) != null))   //開啟Window開啟檔案管理對話視窗
            //{
                using (StreamReader myTextFileReader = new StreamReader(@txt_p5_fileName.Text, Encoding.Default))
                {
                    SortedList DocTermTF_slist = new SortedList();            //暫存每篇文件之"詞彙列表" 及詞會出現次數 TF

                    // ---- 處理文件集中的每一篇文件與文件中的所有詞彙 (詞彙列表 ,計算 TF , 計算 DF)----
                    while (myTextFileReader.EndOfStream != true)       //非資料檔案結尾
                    {
                        line = (myTextFileReader.ReadLine().Trim()) + " ";

                        // ----- 若為一篇新文件的開始
                        if (line.Contains("<DocNo>"))
                        {
                            TotalDocNum++;
                            DocNO_tag = line.Substring(line.IndexOf("<DocNo>"), ((line.IndexOf("</DocNo>") + 8) - line.IndexOf("<DocNo>")));
                            //line = line.Remove(line.IndexOf("<DocNo>"), ((line.IndexOf("</DocNo>") + 8) - line.IndexOf("<DocNo>")));

                            // -------------------------------------- 計算詞彙之 DF ---------------------------------------------------
                            if (temp_list.Count != 0)
                            {
                                foreach (string str in temp_list)               //複製與計算出現在每一文件之詞彙的"DF"
                                {
                                    //-- 計算詞彙之 DF --
                                    if ((index = Term_DF_slist.IndexOfKey(str)) != -1)    //在termList中找尋特定的Term,若Term不存在怎則傳回-1
                                    {
                                        value = (int)Term_DF_slist.GetByIndex(index);
                                        value++;                                        //該term存在,則term的出現次數加1
                                        Term_DF_slist.SetByIndex(index, value);
                                    }
                                    else
                                        Term_DF_slist.Add(str, 1);               //該詞彙(term)不存在於termList中,則加入該新詞彙(term)
                                }
                                temp_list.Clear();     //清除紀錄每一篇文件內之詞彙的SortedList
                            }
                            //--------------記錄每篇文件出現之詞彙集詞頻 -Doctermtf_list -----------
                            if (DocTermTF_slist.Count != 0)
                            {
                                foreach (DictionaryEntry obj in DocTermTF_slist)
                                {
                                    KeyValuePair<string, int> x = new KeyValuePair<string, int>((string)obj.Key, (int)obj.Value);
                                    Doctermtf_list.Add(x);
                                }
                                DocTermTF_slist.Clear();
                            }
                            KeyValuePair<string, int> x1 = new KeyValuePair<string, int>(DocNO_tag, 0); //加入文件編號標籤
                            Doctermtf_list.Add(x1);
                            
                            continue;
                        }

                        // ---------------------------------- 記錄並產生整個文件集之"詞彙列表" 及詞會出現次數 TF------------------------------------------
                        string[] terms = line.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string str in terms)
                        {
                            if (str.IndexOf("(") == -1 || str.IndexOf(")") == -1)       //找不到'()'，就不處理/
                                continue;

                            Term = str.Substring(0, str.IndexOf("("));
                            POS = str.Substring(str.IndexOf("("));
                            if (Term.Length != 0)        //若有"詞彙"存在
                            {
                                //--- 記錄每篇文件之詞彙及其出現次數 ---
                                if ((pt = DocTermTF_slist.IndexOfKey((Term + POS))) != -1)
                                {
                                    value = (int)DocTermTF_slist.GetByIndex(pt);
                                    value++;                                        //該term存在,則term的出現次數加1
                                    DocTermTF_slist.SetByIndex(pt, value);         //紀錄整個文件集之"詞彙列表" 及詞會出現次數 TF
                                }
                                else
                                    DocTermTF_slist.Add((Term + POS), 1);             //該詞彙(term)不存在於termList中,則加入該新詞彙(term)

                                // --- 紀錄出現在每一文件之詞彙有哪些,用於計算詞彙之DF 
                                if ((index = temp_list.IndexOf((Term + POS))) == -1)
                                    temp_list.Add((Term + POS));                //紀錄出現在每一文件之詞彙有哪些,用於計算詞彙之DF 
                            }
                        }
                        //TF詞頻比較
                        if (Doctermtf_list.Count != 0)
                        {
                            foreach (KeyValuePair<string, int> kvp in Doctermtf_list)
                            {
                                //KeyValuePair<string, int> x = new KeyValuePair<string, int>((string)obj.Key, (int)obj.Value);
                                
                            }
                        }
                    }//End of while每一文件
                    // ----記錄並計算最後一篇文件之詞彙 DF --------
                    if (temp_list.Count != 0)
                    {
                        foreach (string str in temp_list)               //複製與計算出現在每一文件之詞彙的"DF"
                        {
                            //-- 計算詞彙之 DF --
                            if ((index = Term_DF_slist.IndexOfKey(str)) != -1)    //在termList中找尋特定的Term,若Term不存在怎則傳回-1
                            {
                                value = (int)Term_DF_slist.GetByIndex(index);
                                value++;                                        //該term存在,則term的出現次數加1
                                Term_DF_slist.SetByIndex(index, value);
                            }
                            else
                                Term_DF_slist.Add(str, 1);               //該詞彙(term)不存在於termList中,則加入該新詞彙(term)
                        }
                        temp_list.Clear();     //清除紀錄每一篇文件內之詞彙的SortedList
                    }
                    if (DocTermTF_slist.Count != 0)
                    {
                        foreach (DictionaryEntry obj in DocTermTF_slist)
                        {
                            KeyValuePair<string, int> x = new KeyValuePair<string, int>((string)obj.Key, (int)obj.Value);
                            Doctermtf_list.Add(x);
                        }
                        DocTermTF_slist.Clear();
                    }

                }//End of using

                // --- 顯示詞彙列表 "term-tf" (未過濾前) ---
                StringBuilder SB = new StringBuilder();            //StringBuilder適合未知數目之字串動態串接,此處使用string物件則無法完成此大資料檔案的讀取串接
                foreach (KeyValuePair<string, int> kvp in Doctermtf_list)
                    SB.AppendLine(kvp.Key + "  " + kvp.Value);
                rtb_p5_souce.Text = SB.ToString();      //顯示經詞彙挑選後資料
            //}//End of if


            //  ============================================== 依詞彙出現次數挑選過濾詞彙  ==============================================
            int tf = int.Parse(txt_p5_tf.Text);         //過濾掉詞彙次數tf(txt_p5_tf.Text)以下之詞彙
            List<KeyValuePair<string, int>> DoctermtfSelected_list = new List<KeyValuePair<string, int>>();
            MaxTfValue = 0;
            foreach (KeyValuePair<string, int> kvp in Doctermtf_list)
            {
                if (kvp.Value >= tf || kvp.Value == 0)
                {
                    DoctermtfSelected_list.Add(kvp);        //建立"挑選過濾後之詞彙-詞彙頻率列表"
                    if (kvp.Value > MaxTfValue)             //找尋最大"詞彙頻率TF"值
                        MaxTfValue = kvp.Value;
                }
                if (kvp.Value > tf && Term_list.Contains(kvp.Key) == false)
                    Term_list.Add(kvp.Key);                 //建立"挑選過濾後之全部詞彙列表"
                Term_list.Sort();
            }
            
            // --- 顯示詞彙列表 "term-tf" (過濾後) ---
            Dictionary<string, int> termSelected_DF_dic = new Dictionary<string, int>();    //挑選過濾後之"詞彙總數"列表及詞彙之"文件頻df"
            rtb_p5_selected.Clear();
            StringBuilder SB1 = new StringBuilder();            //StringBuilder適合未知數目之字串動態串接,此處使用string物件則無法完成此大資料檔案的讀取串接
            StringBuilder SB2 = new StringBuilder();
            foreach (KeyValuePair<string, int> kvp in DoctermtfSelected_list)
            {
                SB1.AppendLine(kvp.Key + "  " + kvp.Value.ToString());

                if (((index = Term_DF_slist.IndexOfKey(kvp.Key)) != -1) && (termSelected_DF_dic.ContainsKey(kvp.Key) == false)) //詞彙挑選後之"term-TF" 與 "term-DF"相互配對對應
                {
                    termSelected_DF_dic.Add((string)Term_DF_slist.GetKey(index), (int)Term_DF_slist.GetByIndex(index));
                    SB2.AppendLine((string)Term_DF_slist.GetKey(index) + " " + (int)Term_DF_slist.GetByIndex(index));
                }
            }
            rtb_p5_selected.Text = SB1.ToString();      //顯示經詞彙挑選後資料
            rtb_p5_termDF.Text = SB2.ToString();        //顯示經詞彙挑選後資料 


            /*/ ============================================= 計算詞彙的權重 "tf_idf" =============================================
            StringBuilder SB3 = new StringBuilder();
            Dictionary<string, double> termSelected_TDIDF_dic = new Dictionary<string, double>();
            int TF = 0, DF = 0;
            double idf = 0.0, tf_idf = 0.0;
            foreach (KeyValuePair<string, int> kvp in DoctermtfSelected_list)
            {
                if (kvp.Key.Contains("<DocNo>") && (kvp.Value == 0))
                {
                    KeyValuePair<string, double> tag = new KeyValuePair<string, double>(kvp.Key, 0.0);
                    DocTermtfidf_list.Add(tag);
                    tf_idf = 0.0;
                }

                TF = kvp.Value;
                if (termSelected_DF_dic.TryGetValue(kvp.Key, out DF))
                {
                    idf = 1 / (double)DF;
                    tf_idf = (TF / (double)MaxTfValue) * (Math.Log(1 / idf));       //基底=e
                }
                KeyValuePair<string, double> x = new KeyValuePair<string, double>(kvp.Key, tf_idf);
                DocTermtfidf_list.Add(x);

                SB3.AppendLine(kvp.Key + " " + tf_idf);
            }
            rtb_p5_tfidf.Text = SB3.ToString();      //顯示經詞彙挑選後之"詞彙與tfidf" 

            // ---- 顯示經詞彙出現次數挑選後出現在每一篇文件中的所有詞彙
            StringBuilder SB4 = new StringBuilder();
            foreach (string str in Term_list)
                SB4.AppendLine(str);
            rtb_p5_TermList.Text = SB4.ToString();

        }//End 0f Event Fuction  */

        // ============================================= 詞彙的權重都設為1 =============================================
            StringBuilder SB3 = new StringBuilder();
            Dictionary<string, double> termSelected_TDIDF_dic = new Dictionary<string, double>();
            int TF = 0, DF = 0;
            double idf = 0.0, tf_idf = 0.0;
            foreach (KeyValuePair<string, int> kvp in DoctermtfSelected_list)
            {
                if (kvp.Key.Contains("<DocNo>") && (kvp.Value == 0))
                {
                    KeyValuePair<string, double> tag = new KeyValuePair<string, double>(kvp.Key, 0.0);
                    DocTermtfidf_list.Add(tag);
                    tf_idf = 0.0;
                }

               // TF = kvp.Value;
                if (termSelected_DF_dic.TryGetValue(kvp.Key, out DF))
                {
                    //idf = 1 / (double)DF;
                    tf_idf = 1;       //基底=e
                }
                KeyValuePair<string, double> x = new KeyValuePair<string, double>(kvp.Key, tf_idf);
                DocTermtfidf_list.Add(x);

                SB3.AppendLine(kvp.Key + " " + tf_idf);
            }
            rtb_p5_tfidf.Text = SB3.ToString();      //顯示經詞彙挑選後之"詞彙與tfidf" 

            // ---- 顯示經詞彙出現次數挑選後出現在每一篇文件中的所有詞彙
            StringBuilder SB4 = new StringBuilder();
            foreach (string str in Term_list)
                SB4.AppendLine(str);
            rtb_p5_TermList.Text = SB4.ToString();

        }//End 0f Event Fuction


        // ======================================================存檔===========================================================
        private void btn_p5_saveFile_Click(object sender, EventArgs e)
        {
           // if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p5_fileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
                txt_p5_fileName.Text = dirPath1 + @"\test.sel";
                File.WriteAllText(txt_p5_fileName.Text, rtb_p5_selected.Text, Encoding.Default);
                MessageBox.Show("test.sel儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void btn_p5_saveFileDF_Click(object sender, EventArgs e)
        {
            //if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p5_fileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
            txt_p5_fileName.Text = dirPath1 + @"\df";
                File.WriteAllText(txt_p5_fileName.Text, rtb_p5_termDF.Text, Encoding.Default);
                MessageBox.Show("tf儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void btn_p5_saveFileTerm_TFIDF_Click(object sender, EventArgs e)
        {
            //if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p5_fileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
            txt_p5_fileName.Text = dirPath1 + @"\tfidf";
            txt_p6_OpenDocTermList.Text = txt_p5_fileName.Text;
             File.WriteAllText(txt_p5_fileName.Text, rtb_p5_tfidf.Text, Encoding.Default);
             MessageBox.Show("tfidf儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }
        private void btn_p5_saveFileTF_Click(object sender, EventArgs e)
        {
           // if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p5_fileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
            txt_p5_fileName.Text = dirPath1 + @"\tf";
                File.WriteAllText(txt_p5_fileName.Text, richTextBox5.Text, Encoding.Default);
                MessageBox.Show("tf儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }
        #endregion

        #region 相似矩陣
        //------------------------------------------------------------------------------------------------------------------------------------------------
        //----tabPage6----相似矩陣----tabPage6----相似矩陣----tabPage6----相似矩陣----tabPage6----相似矩陣----tabPage6----相似矩陣----tabPage6----相似矩陣
        //------------------------------------------------------------------------------------------------------------------------------------------------
        // ---------------------------- 建立詞彙權重矩陣 ------------------------------------------
        private void btn_p6_OpenDocTermList_Click(object sender, EventArgs e)
        {
            //if ((openFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p6_OpenDocTermList.Text = openFileDialog1.FileName) != null))   //開啟Window開啟檔案管理對話視窗
          //  {
                int DocNo = 0;
                string line = "", Term = "";
                double tfidf = 0.0;
                TotalDocNum = 0;
                using (StreamReader myTextFileReader0 = new StreamReader(@txt_p6_OpenDocTermList.Text, Encoding.Default))
                {
                    Term_list.Clear();
                    // -------------建立總詞彙列表---------------
                    while (myTextFileReader0.EndOfStream != true)       //非資料檔案結尾
                    {
                        line = myTextFileReader0.ReadLine().Trim();

                        if (line.Contains("<DocNo>"))
                            TotalDocNum++;
                        else
                        {
                            Term = line.Substring(0, line.IndexOf(" "));     //詞彙
                            if (Term_list.IndexOf(Term) == -1)
                                Term_list.Add(Term);
                        }
                    }
                }
                Term_list.Sort();
                TotalTermNum = Term_list.Count;

                //顯示全部詞彙
                StringBuilder SB = new StringBuilder();
                SB.AppendLine("共" + Term_list.Count.ToString() + "個詞彙");
                foreach (string st in Term_list)
                    SB.AppendLine(st);
                rtb_p6_TermList.Text = SB.ToString();


                // --- 建立詞彙權重矩陣 ---
                int TermNo = Term_list.Count;
                Double[,] DocTerm_mtx = new Double[TotalDocNum, TotalTermNum];
                using (StreamReader myTextFileReader = new StreamReader(@txt_p6_OpenDocTermList.Text, Encoding.Default))
                {
                    StringBuilder SB1 = new StringBuilder();

                    while (myTextFileReader.EndOfStream != true)       //非資料檔案結尾
                    {
                        line = myTextFileReader.ReadLine().Trim();
                        SB1.AppendLine(line);

                        if (line.Contains("<DocNo>"))               //文件起始
                        {
                            DocNo = Int32.Parse(line.Substring((line.IndexOf("<DocNo>") + 7), (line.IndexOf("</DocNo>") - (line.IndexOf("<DocNo>") + 7)))) - 1; //讀取文件編號
                            if (DocNo > TotalDocNum)           //讀得之文件編號不得大於總文件數
                            {
                                MessageBox.Show("文件數目計數錯誤", "建立文件矩陣錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                break;
                            }
                        }
                        else if (line.Length < 2)       //非文字行則刪除
                        {
                            //Delete this line
                        }
                        else
                        {
                            Term = line.Substring(0, line.IndexOf(" "));     //詞彙
                            tfidf = double.Parse(line.Substring(line.IndexOf(" ") + 1));
                            if ((TermNo = Term_list.IndexOf(Term)) != -1)
                                DocTerm_mtx[DocNo, TermNo] = tfidf;
                            else
                            {
                                MessageBox.Show("詞彙列表找不到該詞彙", "建立文件矩陣錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                break;
                            }
                        }
                    }
                    rtb_p6_termTFIDF.Text = SB1.ToString();
                }
                // ---顯示詞彙權重矩陣 ---
                StringBuilder SB2 = new StringBuilder();
                for (int i = 0; i < TotalDocNum; i++)
                {
                    for (int j = 0; j < TotalTermNum; j++)
                        SB2.Append(DocTerm_mtx[i, j].ToString() + "/");
                    SB2.AppendLine();
                }
                rtb_p6_DocTermMtx.Text = SB2.ToString();

                //"文件-詞彙權重"矩陣轉置
                Double[,] TermDoc_mtx = new Double[TotalTermNum, TotalDocNum];
                for (int i = 0; i < TotalDocNum; i++)      //文件數
                    for (int j = 0; j < TotalTermNum; j++)   //詞彙數
                        TermDoc_mtx[j, i] = DocTerm_mtx[i, j];

                //文件相似矩陣
                StringBuilder SB3 = new StringBuilder();
                Double[,] DocSimilarity_mtx = new Double[TotalDocNum, TotalDocNum];
                double sum = 0.0, sum4 = 0.0, sum5 = 0.0, sumAns = 0.0;
                Int32 DocCount1 = 0, DocCount2 = 0, DocCount3 = 0;
                for (int i = 0; i < TotalDocNum; i++)       //文件數
                {
                    DocCount1 = 0; DocCount2 = 0;
                    sum = 0.0;
                    for (int m = 0; m < i; m++)     //矩陣下三角顯示"0"
                        SB3.Append("0" + "/");
                    SB3.Append("0" + "/");          //矩陣對角線元素顯示"0"

                    for (int j = (i + 1); j < TotalDocNum; j++)   //文件數 
                    {
                        DocCount1 = 0; DocCount2 = 0;
                        sum = 0.0;
                        for (int k = 0; k < TotalTermNum; k++) //詞彙數
                        {
                            if (DocTerm_mtx[i, k] != 0.0)   //計算文件i的詞彙數
                                DocCount1++;
                            if (TermDoc_mtx[k, j] != 0.0)   //計算文件j的詞彙數
                                DocCount2++;
                            if (DocTerm_mtx[i, k] != 0.0 && TermDoc_mtx[k, j] != 0.0)   //計算文件的相似值
                                 sum +=1;
                         }
                        if (DocCount1 > DocCount2)

                            sum = sum/DocCount2;
                        else
                            sum = sum/DocCount1;
                        sum4 += sum;
                        DocCount3++;

                        DocSimilarity_mtx[i, j] = sum;
                        SB3.Append(Convert.ToString (sum) + "/");
                    }
                    SB3.AppendLine();
                }
                //算出 最大值與最小值之間的數值
                sum5 = sum4 / DocCount3;
                rtb_p6_DocSimilarityMtx.Text = SB3.ToString();
                double maxValue2 = 0.0;
                int indexX = 0, indexY = 0;
                for (int i = 0; i < TotalDocNum; i++)      //文件數
                    for (int j = (i + 1); j < TotalDocNum; j++)   //詞彙數
                        if (DocSimilarity_mtx[i, j] > maxValue2)
                        {
                            maxValue2 = DocSimilarity_mtx[i, j];
                            indexX = i;
                            indexY = j;
                        }
                sumAns = Math.Round ( maxValue2 - sum5,1);
                txb_Threshold.Text =Convert .ToString (sumAns * Threshold);

                //詞彙相似矩陣
                StringBuilder SB4 = new StringBuilder();
                Double[,] TermSimilarity_mtx = new Double[TotalTermNum, TotalTermNum];
                
                double sum1 = 0.0;
                for (int i = 0; i < TotalTermNum; i++)      //詞彙數
                {
                    for (int m = 0; m < i; m++)
                        SB4.Append("0" + "/");      //矩陣下三角顯示"0"
                    SB4.Append("0" + "/");          //矩陣對角線元素顯示"0"

                    for (int j = (i + 1); j < TotalTermNum; j++)  //詞彙數 
                    {
                        sum1 = 0.0;
                        for (int k = 0; k < TotalDocNum; k++)     //文件數
                            sum1 += TermDoc_mtx[i, k] * DocTerm_mtx[k, j];
                        TermSimilarity_mtx[i, j] = sum1;
                        SB4.Append(sum1.ToString() + "/");
                    }
                    SB4.AppendLine();
                }
                rtb_p6_TermSimilarityMtx.Text = SB4.ToString();
                
                // --找大於中界值txb_Threshold的文件
                ArrayList doclist = new ArrayList();
                //int indexX = 0, indexY = 0;
                lbl_p6.Text = null;
                lbl_p10_value.Text = null;
                lbl_p8_Value.Text = null;
                for (int i = 0; i < TotalDocNum; i++)      //文件數
                    for (int j = (i + 1); j < TotalDocNum; j++)   //詞彙數
                        if (DocSimilarity_mtx[i, j] > double.Parse(txb_Threshold.Text))
                        {
                            if (doclist.IndexOf(i) == -1)
                                doclist.Add(i);
                            if (doclist.IndexOf(j) == -1)
                                doclist.Add(j);
                        }
                doclist.Sort();
                for (int x = 0; x < doclist.Count; x++)
                {
                    int docNun = Int32 .Parse ( doclist[x].ToString () )+1;
                    lbl_p6.Text += docNun.ToString ()+",";
                    if (txt_p6_OpenDocTermList.Text.IndexOf("相似文件") != -1)
                    {
                        lbl_p10_value.Text += docNun.ToString() + ",";
                        txt_p10_openFileName.Text = dirPath1 + @"\test.txt";
                    }
                    else
                        lbl_p8_Value.Text = lbl_p6.Text;

                }
                  
                /*/ --- 找最大者 ---
                double maxValue = 0.0;
                int indexX = 0, indexY = 0;
                for (int i = 0; i < TotalDocNum; i++)      //文件數
                    for (int j = (i + 1); j < TotalDocNum; j++)   //詞彙數
                        if (DocSimilarity_mtx[i, j] > maxValue)
                        {
                            maxValue = DocSimilarity_mtx[i, j];
                            indexX = i;
                            indexY = j;
                        }
                lbl_p6.Text = indexX.ToString() + "  " + indexY.ToString() + "  " + maxValue.ToString();
                 */
                 
                btn_p6_saveDocTermWeight_mtx.Enabled = true;
                btn_p6_saveDocSimilarity_mtx.Enabled = true;
                btn_p6_saveTermSimilarity_mtx.Enabled = true;
            //}
        }

        // ======================================================存檔===========================================================
        private void btn_p6_saveDocTermWeight_mtx_Click(object sender, EventArgs e)
        {
            if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p6_SaveFileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
                File.WriteAllText(txt_p6_SaveFileName.Text, rtb_p6_DocTermMtx.Text, Encoding.Default);
        }

        private void btn_p6_saveDocSimilarity_mtx_Click(object sender, EventArgs e)
        {
            if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p6_SaveFileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
                File.WriteAllText(txt_p6_SaveFileName.Text, rtb_p6_DocSimilarityMtx.Text, Encoding.Default);
           // btn_p7_openDocSimilarMtx.Enabled = true;
        }

        private void btn_p6_saveTermSimilarity_mtx_Click(object sender, EventArgs e)
        {
            if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p6_SaveFileName.Text = saveFileDialog1.FileName) != null))    //開啟Window儲存檔案管理對話視窗設定檔名
                File.WriteAllText(txt_p6_SaveFileName.Text, rtb_p6_TermSimilarityMtx.Text, Encoding.Default);
           // btn_p7_openTermSimilarMtx.Enabled = true;
        }
        //====操作臨界值====
        double Threshold = 0.0;
        private void trk_Threshold_Scroll(object sender, EventArgs e)
        {
            lbl_p6_Value.Text  = trk_Threshold.Value.ToString()+"%";
            Threshold = percent(trk_Threshold.Value);
            btn_p6_OpenDocTermList.Enabled = true;
        }

        double percent(int Value)
        {
            double x = 0.0;
             x = (double)Value / 100;
            return x;
        }
        #endregion

        #region 語音輸出
        private void button1_Click(object sender, EventArgs e)
        {
            SpeechRecognitionEngine SRE = new SpeechRecognitionEngine();
            //SRE.SetInputToDefaultAudioDevice();
            GrammarBuilder GB = new GrammarBuilder();
            GB.Append(new Choices(new string[] { "銅葉綠素", "華夏技術學院","服貿","馬來西亞航空","王鴻遠","北捷殺人","阿帕契","課綱"}));
            Grammar G = new Grammar(GB);
            G.Name = "main command grammar";
            SRE.LoadGrammar(G);
            DictationGrammar DG = new DictationGrammar(); //自然發音
            DG.Name = "dictation";
            SRE.LoadGrammar(DG);
            try
            {

                label4.Text = "Speak Now";
                SRE.SetInputToDefaultAudioDevice();
                
                //SRE.RecognizeAsync(RecognizeMode.Multiple);
                RecognitionResult result = SRE.Recognize();
                if (result == null) return; 
                label8.Text = result.Text;
                txtKeyword1.Text = result.Text;
            }
            catch (InvalidOperationException exception)
            {
                //button1.Text = String.Format("Could not recognize input from default aduio device. Is a microphone or sound card available?\r\n{0} - {1}.", exception.Source, exception.Message);
            }
            finally
            {
                SRE.UnloadAllGrammars();
            }                          
           
        }
        #endregion

        #region 新聞文件
        //------------------------------------------------------------------------------------------------------------------------------------------------
        //----tabPage7----新聞文件------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// 建置新聞目錄
        /// </summary>
        /// <param name="strPath"></param>
        private void PopulateTreeView()
        {
           // treeView1.Nodes.Clear();
            string MainDir = null;
            if (rdoyahoo.Checked == true)
            {
                MainDir = strCurrentPath + rdoyahoo.Text + comboBox1.Text + comboBox2.Text + "篇";
                strSelectDir = MainDir;
            }

            if (rdo_google.Checked == true)
            {
                MainDir = strCurrentPath + rdo_google.Text + comboBox1.Text + comboBox2.Text + "篇";
                strSelectDir = MainDir;
            }
                
            if (Directory.Exists(MainDir))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(MainDir);
                FileInfo[] totFile = DirInfo.GetFiles();
                if (totFile.Length != 0)
                {
                    foreach (FileInfo subFile in totFile)
                    {
                        if (subFile.Name != "data.txt")
                        {
                            listBox1.Items.Add(subFile.Name);
                        }
                    }
                }
                txtShowData.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            PopulateTreeView();
            PopulateNews();
            PopulateNews2();
            ckb_File_1.Checked = false;
            ckb_File_2.Checked = true;
            txt_p1_openFileName.Text  = strSelectDir1;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;

            string curItem = listBox1.SelectedItem.ToString();
            string MainDir = strSelectDir;
            DirectoryInfo Dir = new DirectoryInfo(MainDir);
            FileInfo[] totFile = Dir.GetFiles(curItem);
            if (totFile.Length != 0)
            {
                foreach (FileInfo subFile in totFile)
                {
                    StreamReader sr = new StreamReader(subFile.FullName, Encoding.Default);
                    txtShowData.Text = sr.ReadToEnd();
                    sr.Close();
                }
            }
        }
        private void PopulateNews()
        {
            // treeView1.Nodes.Clear();
            String[] NewsNun = lbl_p8_Value.Text.Split(new char[] { ',' });
            string MainDir = null;
            string MainDir1 = null;
            if (rdoyahoo.Checked == true)
            {
                MainDir = strCurrentPath + rdoyahoo.Text + comboBox1.Text + comboBox2.Text + "篇";
                strSelectDir = MainDir;
            }

            if (rdo_google.Checked == true)
            {
                MainDir = strCurrentPath + rdo_google.Text + comboBox1.Text + comboBox2.Text + "篇";
                strSelectDir = MainDir;
            }
            MainDir1 = strCurrentPath + comboBox1.Text + "相似文件";
            strSelectDir1 = MainDir1;
            if (!Directory.Exists(MainDir1))
            {
                //建立目錄
                Directory.CreateDirectory(MainDir1);
            }
            else
            {
                if (Directory.Exists(MainDir))
                {
                    DirectoryInfo DirInfo = new DirectoryInfo(MainDir);
                    FileInfo[] totFile = DirInfo.GetFiles();
                    if (totFile.Length != 0)
                    {
                        foreach (FileInfo subFile in totFile)
                        {
                            String[] FileName = subFile.Name.Split(new char[] { '.' });
                            for (int i = 0; i < NewsNun.Length; i++)
                            {
                                if (FileName[0] == NewsNun[i])
                                {
                                    string fullpath1 = MainDir +"\\"+ subFile.Name;
                                    string fullpath2 = MainDir1 + "\\" + subFile.Name;
                                    //listBox2.Items.Add(subFile.Name);
                                    if (!File.Exists(fullpath2))
                                    {
                                        File.Copy(fullpath1, fullpath2);
                                    }
                                }
                            }
                        }
                    }
            }
           
                txtShowData.Text = "";
            }
        }
        private void PopulateNews2()
        {
            // treeView1.Nodes.Clear();
            String[] NewsNun = lbl_p8_Value.Text.Split(new char[] { ',' });
            string MainDir = null;
            string MainDir1 = null;
            if (rdoyahoo.Checked == true)
            {
                MainDir = strCurrentPath + rdoyahoo.Text + comboBox1.Text + comboBox2.Text + "篇";
                strSelectDir = MainDir;
            }

            if (rdo_google.Checked == true)
            {
                MainDir = strCurrentPath + rdo_google.Text + comboBox1.Text + comboBox2.Text + "篇";
                strSelectDir = MainDir;
            }
            MainDir1 = strCurrentPath + comboBox1.Text + "相似文件";
            strSelectDir1 = MainDir1;
                if (Directory.Exists(MainDir1))
                {
                    DirectoryInfo DirInfo = new DirectoryInfo(MainDir1);
                    FileInfo[] totFile = DirInfo.GetFiles();
                    if (totFile.Length != 0)
                    {
                        foreach (FileInfo subFile in totFile)
                        {
                            
                            listBox2.Items.Add(subFile.Name);
                            
                        }
                    }
                }

                txtShowData.Text = "";
            
        }
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;

            string curItem = listBox2.SelectedItem.ToString();
            string MainDir = strSelectDir1;
            DirectoryInfo Dir = new DirectoryInfo(MainDir);
            FileInfo[] totFile = Dir.GetFiles(curItem);
            if (totFile.Length != 0)
            {
                foreach (FileInfo subFile in totFile)
                {
                    StreamReader sr = new StreamReader(subFile.FullName, Encoding.Default);
                    txtShowData.Text = sr.ReadToEnd();
                    sr.Close();
                }
            }
        }
        #endregion
     
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //textBox1.Text = e.Node.Tag.ToString ();
        }


        #region 段落濃縮
        private void btn_p10_openFiles_Click(object sender, EventArgs e)
        {
            //if ((openFileDialog1.ShowDialog() == DialogResult.OK) && ((txt_p10_openFileName.Text  = openFileDialog1.FileName) != null))   //開啟Window開啟檔案管理對話視窗  
            //{
                ArrayList sectionNumber = new ArrayList();
                String[] Number = lbl_p10_value.Text.Split(new char[] { ',' });
                String line = null;
                StringBuilder SB1 = new StringBuilder();
                StringBuilder SB2 = new StringBuilder();
                StringBuilder SB3 = new StringBuilder();
                StringBuilder SB4 = new StringBuilder();
                int DocNo = 0;
                for (int i = Number.Length - 1; i >= 1; i--)
                {
                    sectionNumber.Add(Number[i]);
                }
                    using (StreamReader myTextFileReader0 = new StreamReader(@txt_p10_openFileName.Text, Encoding.Default))
                    {
                        while (myTextFileReader0.EndOfStream != true)
                        {
                            line = myTextFileReader0.ReadLine().Trim();
                            SB1.AppendLine(line);
                        }
                        richTextBox2.Text = SB1.ToString();
                    }
                using (StreamReader myTextFileReader = new StreamReader(@txt_p10_openFileName.Text, Encoding.Default))
                {
                    while (myTextFileReader.EndOfStream != true)
                    {
                        line = myTextFileReader.ReadLine().Trim();
                        if (line.Contains("<DocNo>"))
                        {
                            DocNo = Int32.Parse(line.Substring((line.IndexOf("<DocNo>") + 7), (line.IndexOf("</DocNo>") - (line.IndexOf("<DocNo>") + 7)))); //讀取文件編號
                        }
                        if (sectionNumber.Contains(DocNo.ToString())) 
                        {
                            SB2.AppendLine(line);
                        }
                        else
                        {
                            SB3.AppendLine(line);
                        }
                       
                    }
                    richTextBox3.Text = SB3.ToString();
                    
                }
                using (StreamReader myTextFileReader = new StreamReader(@txt_p10_openFileName.Text, Encoding.Default))
                {
                    while (myTextFileReader.EndOfStream != true)
                    {
                        line = myTextFileReader.ReadLine().Trim();
                        if (line.Contains("<DocNo>"))
                        {
                            DocNo = Int32.Parse(line.Substring((line.IndexOf("<DocNo>") + 7), (line.IndexOf("</DocNo>") - (line.IndexOf("<DocNo>") + 7)))); //讀取文件編號
                        }
                        if (sectionNumber.Contains(DocNo.ToString()))
                        {
                            SB2.AppendLine(line);
                        }
                        else
                        {
                            if (line.Contains("<DocNo>"))
                            {
                                line  = line.Remove(line.IndexOf("<DocNo>"),line.IndexOf ("</DocNo>")+8);
                                line = line.Trim("↖".ToCharArray());
                            }
                            SB4.AppendLine(line);
                        }

                    }
                    synh.SpeakAsync(SB4.ToString());
                    DialogResult DlogRut = MessageBox.Show(SB4.ToString());
                    if (DlogRut == DialogResult.OK)
                    {
                        synh.SpeakAsyncCancelAll();

                    }
                }

            //}
        }
        private void btn_p10_save_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)        //開啟Window儲存檔案管理對話視窗設定檔名
            {
                txt_p10_saveFileName.Text = saveFileDialog1.FileName; //將選取的檔案名稱顯示於於txt_saveFileInfo(textBox)視窗
                File.WriteAllText(@txt_p10_saveFileName.Text, richTextBox3.Text, Encoding.Default);
                 MessageBox.Show("合併文件儲存完成", "檔案儲存作業", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
              
            }
        }
        #endregion

        private void btn_savenews_Click(object sender, EventArgs e)
        {
            string MainDir = null;
            Encoding Ecd = null;
            MainDir = strCurrentPath + "單幅新聞文件";
            if (!Directory.Exists(MainDir))
            {
                //建立目錄
                Directory.CreateDirectory(MainDir);

            }
            Ecd = Encoding.Unicode;
            string strFileName = null;
            
            strFileName  =webBrowser2.Document.Title;
            string dirPath2 = MainDir + @"\"+ strFileName +".txt" ;
            if (!File.Exists(dirPath2))
            {
                File.WriteAllText(dirPath2, richTextBox1.Text, Encoding.Default);
                MessageBox.Show("儲存成功");
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("檔案已存在,是否要覆寫檔案", "檔案儲存作業", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (dialogResult == DialogResult.OK)
                {
                    File.WriteAllText(dirPath2, richTextBox1.Text, Encoding.Default);
                    MessageBox.Show("儲存成功");
                }
            }

        }

        private void 檔案資料開啟ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form frmForm = frmOpenFile.Instance;
            //frmForm.MdiParent = this;
            //frmForm.MaximizeBox = true;
            frmForm.Show();
        }

    }
}
