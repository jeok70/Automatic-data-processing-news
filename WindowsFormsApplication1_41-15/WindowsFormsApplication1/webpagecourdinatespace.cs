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
using WindowsFormsApplication1;

namespace WindowsFormsApplication4
{
    class webpagecourdinatespace
    {
        #region 變數宣告
        private String Text1 { get; set; }
        private int number { get; set; }
        int inttotlevel = 0;
        int rcount = 0,testans=0;
        int championxy = 0, Challengerxy = 0;
        string championpath = null, Challengerpath = null,Ans=null;
        #endregion
        #region  Get wabsource
        /// <summary>
        /// 將網頁原始碼擷取出來
        /// </summary>
        /// <param name="HD">轉換網頁的物件</param>
        /// <returns>回傳出網頁的原始碼</returns>
        public string Getwebpagesource(HtmlDocument HD)
        {
            JQ(HD);
            HtmlElement htmlElement = HD.GetElementsByTagName("HTML")[0];
            return htmlElement.InnerHtml;
        }
        public void JQ(HtmlDocument HD)
        {
             HtmlElement HE = HD.CreateElement("SCRIPT");
             HE.SetAttribute("src", "https://code.jquery.com/jquery-1.11.1.min.js");
             HD.Body.AppendChild(HE);
             HtmlElement HE2 = HD.CreateElement("SCRIPT");
             HE2.SetAttribute("text", Properties.Resources.attrNodeTopLeft);
             HD.Body.AppendChild(HE2);
             HD.InvokeScript("attrNodeTopLeft");
        }
        #endregion
        #region  Create Domtree
        public void create(TreeNode rt,IHTMLDocument3 doc)
        {
            IHTMLDOMNode rootDomNode = (IHTMLDOMNode)doc.documentElement;
            rt.Tag = rootDomNode;
            inttotlevel = 1;
            InsertDOMNodes(rootDomNode, rt, inttotlevel, 1);
        }
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
        #endregion
        #region  Judge 
        /*之前寫的專題原型會在Button上顯示所有的Table標籤數量，改用的方式來寫時發現為了傳出參數會使整個程式重複運算＝無結果
        ，在尚未找到解決方法的情況下就將顯示Table標籤數量的部份給去除掉，不過本來這功能只是拿來方便對造程式有無演算錯誤而設的，故去除此功能對專題並無任何影響。
         */ 
        #region  tablecount
        public void Tablecount(TreeNode treeNode)
        {
            if (treeNode.Text == "<TABLE>")
            {
                testans += 1;
              
            }
          

        foreach(TreeNode tn in treeNode.Nodes)
        {
            Tablecount(tn);
        }
        //return testans.ToString();
           
        }
        #endregion

        #region  tablecount2
        ArrayList txtlist = new ArrayList();
        bool startcheck = false;
        public void Tablecount2(TreeNode treenode, int weight = 0, int Minweigh = 0, String txtnode = null)
        {
            int x = 0, y = 0, xy = 0;
            String top = null;
            String left = null;
            string width = null;
            string height = null;
            TreeNode parentnode = treenode.Parent;
            // int nodeCount = treenode.Parent.Nodes.Count;

            try
            {
                if (treenode.Tag is mshtml.IHTMLDOMNode)
                {
                    mshtml.IHTMLElement HE = treenode.Tag as mshtml.IHTMLElement;

                    if (HE != null)
                    {

                        string node2 = HE.outerHTML;
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
                        if (treenode.Text != "<TH>")
                        /*關於 if (treenode.Text != "<TH>") :月報表的部分總是出錯並顯示:"已經加入含有相同索引鍵的項目。"
                         * 經查證問題為標籤<"TH">部分所致，然分析後並無結果只以判斷式來擋掉<"TH">，使程式得以順暢運行。
                         * 使用後仍能得到月報表的目標資料為不幸中的大幸，但此舉為下下之策，警惕之。
                         */
                        {
                            var attributes = (from Match mt in Regex.Matches(node2, pattern, RegexOptions.IgnorePatternWhitespace)
                                              select new
                                              {
                                                  Name = mt.Groups["Tag"],
                                                  Attrs = (from cpKey in mt.Groups["Key"].Captures.Cast<Capture>().Select((a, i) => new { a.Value, i })
                                                           join cpValue in mt.Groups["Value"].Captures.Cast<Capture>().Select((b, i) => new { b.Value, i }) on cpKey.i equals cpValue.i
                                                           select new KeyValuePair<string, string>(cpKey.Value, cpValue.Value)).ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
                                              }).First().Attrs;

                            //過濾後
                            if (treenode.Text.Equals("<TABLE>"))
                            {
                                foreach (KeyValuePair<string, string> kvp in attributes)
                                {
                                    switch (kvp.Key)
                                    {
                                        case "top":
                                            top = kvp.Value;
                                            break;
                                        case "left":
                                            left = kvp.Value;
                                            break;
                                        case "height":
                                            height = kvp.Value;
                                            y = Int32.Parse(height);
                                            break;
                                        case "width":
                                            width = kvp.Value;
                                            x = Int32.Parse(width);
                                            break;
                                    }
                                }
                                xy = x * y;
                                rcount += 1;
                                Chanllage(xy, treenode.FullPath, rcount, startcheck);
                                if (rcount == testans)
                                {
                                    startcheck = true;
                                    Chanllage(xy, treenode.FullPath, rcount, startcheck);
                                }
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {

            }
            foreach (TreeNode tn in treenode.Nodes)
            {
                Tablecount2(tn);
            }
            //return testans.ToString();
        }
        #endregion
        #region VS
        public void Chanllage(int targetxy, string targetpath, int roundcount, bool start)
        {
            if (roundcount == 1)
            {
                championxy = targetxy;
                championpath = targetpath;
            }
            if (championxy != 0 && championpath != null && roundcount > 1)
            {
                Challengerxy = targetxy;
                Challengerpath = targetpath;
                if (roundcount == 2)
                {
                    championxy = Challengerxy;
                    championpath = Challengerpath;
                }
                else if (Challengerxy > championxy)
                {
                    championxy = Challengerxy;
                    championpath = Challengerpath;
                }
            }

            if (start == true)
            {
                Ans = championpath;
            }
           
        }
        #endregion



        #region Answer
        public void Answer(TreeNode tn)
        {
            if (tn.FullPath == Ans)
            {
                tn.ForeColor = Color.Red;
                pathcolor(tn);
            }

            foreach (TreeNode node in tn.Nodes)
            {
                Answer(node);
            }
        }
   
    private void pathcolor(TreeNode tn)
    {
        tn.ForeColor = Color.Red;
        foreach (TreeNode node in tn.Nodes)
        {
            pathcolor(node);
        }
    }
        public string Getpath(bool get)
    {
        return Ans;
    }
        #endregion

        #endregion
    }
}
