using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;

namespace WindowsFormsApplication1
{
    class Global
    {
        #region MarkDeleteNode
        /// <summary>
        /// 將可去除之不相關標記注記為紅色,不實際刪除之
        /// <param name="DOMTreeView2">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        public void MarkDeleteNode(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null)
                {
                    return;
                }
                else
                {
                    DOMTreeView.SelectedNode = tree_node;
                    DOMTreeView.Visible = false;    //關閉Tree_View
                    DOMTreeView.ExpandAll();
                }

                while (DOMTreeView.SelectedNode != null)
                {

                    if (DOMTreeView.SelectedNode.Tag is HtmlElement)
                    {
                        HtmlElement em = DOMTreeView.SelectedNode.Tag as HtmlElement ;
                        if (em.TagName == "!" || em.TagName == "SCRIPT" || em.TagName == "STYLE" || em.TagName == "NOSCRIPT" || em.InnerText == null)
                        {
                            DOMTreeView.SelectedNode.BackColor = Color.Red;
                        }
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region DeletNode
        /// <summary>
        ///  將可去除之不相關標記實際刪除之
        /// <param name="DOMTreeView">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        public void DeleteNode(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null)
                {
                    return;
                }
                else
                {
                    DOMTreeView.SelectedNode = tree_node;
                    DOMTreeView.Visible = false;
                    DOMTreeView.ExpandAll();
                }

                while (DOMTreeView.SelectedNode != null)
                {
                    //if (!DOMTreeView.SelectedNode.IsEditing) DOMTreeView.SelectedNode.Expand();

                    if (DOMTreeView.SelectedNode.Tag is HtmlElement)
                    {
                        HtmlElement em = DOMTreeView.SelectedNode.Tag as HtmlElement;
                        if (em.TagName == "!" || em.TagName == "SCRIPT" || em.TagName == "STYLE" || em.TagName == "NOSCRIPT" || em.InnerText == null||em.InnerText == " ")
                        {
                            DOMTreeView.Nodes.Remove(DOMTreeView.SelectedNode);
                            DeleteNode(DOMTreeView, DOMTreeView.SelectedNode);
                        }
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;

                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region MarkDeleteDataArea
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DOMTreeView3"></param>
        /// <param name="tree_node"></param>
        public void MarkDeleteDataArea(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null) return;
                DOMTreeView.SelectedNode = tree_node;
                DOMTreeView.Visible = false;
                //DOMTreeView.Select();
                while (DOMTreeView.SelectedNode != null)
                {
                    if (!DOMTreeView.SelectedNode.IsEditing) DOMTreeView.SelectedNode.Expand();
                    if (DOMTreeView.SelectedNode.Tag == null)
                    {
                        HtmlElement em = DOMTreeView.SelectedNode.Parent.Tag as HtmlElement;
                        string textnode = em.InnerText.ToString();
                        int cnt = (int)DOMTreeView.SelectedNode.GetNodeCount(false);    //傳回子樹狀節點的數目
                        if (textnode.Length < 15)
                        {
                            DOMTreeView.SelectedNode.Parent.BackColor = Color.Green;
                        }
                    }
                    if (DOMTreeView.SelectedNode == null) return;
                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    //DOMTreeView.Select();
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region DeleteDataArea
        /// <summary>
        /// 依據節點內文字的長度來刪除節點
        /// </summary>
        /// <param name="DOMTreeView"></param>
        /// <param name="tree_node"></param>
        public void DeleteDataArea(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null) return;      //如果沒有選取任何 TreeNode

                DOMTreeView.SelectedNode = tree_node;   //設定目前在樹狀檢視控制項中選取的樹狀節點。
                DOMTreeView.Visible = false;
                //DOMTreeView.Select();     // Select() 方法會啟動此控制項。
                while (DOMTreeView.SelectedNode != null)    //如果目前沒有選取任何 TreeNode， SelectedNode 屬性為 null。 
                {
                    if (!DOMTreeView.SelectedNode.IsEditing) DOMTreeView.SelectedNode.Expand();

                    if (DOMTreeView.SelectedNode.Tag == null)   //SelectedNode內沒儲存資料。(Tag 屬性的常見用途是儲存與控制項密切關聯的資料。)
                    {
                        HtmlElement em = DOMTreeView.SelectedNode.Parent.Tag as HtmlElement;    //設定控制項的父容器內儲存的資料為--表示 Web 網頁內的 HTML 項目。
                        string textnode = em.InnerText.ToString();      //InnerText 將會傳回該 HTML 中的所有文字，而且會移除其中的標記。
                        if (textnode.Length < 15)   //節點內文字的長度<15字
                        {
                            DOMTreeView.SelectedNode.Nodes.Remove(DOMTreeView.SelectedNode);
                            DeleteDataArea(DOMTreeView, DOMTreeView.SelectedNode);
                        }
                    }

                    if (DOMTreeView.SelectedNode == null) return;   //如果目前沒有選取任何 TreeNode， SelectedNode 屬性為 null。

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    //DOMTreeView.Select();
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion

        #region DeleteDataArea1
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DOMTreeView"></param>
        /// <param name="tree_node"></param>
        public void DeleteDataArea1(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null) return;      //如果沒有選取任何 TreeNode

                DOMTreeView.SelectedNode = tree_node;
                DOMTreeView.Visible = false;
                //DOMTreeView.Select();
                while (DOMTreeView.SelectedNode != null)    //如果目前沒有選取任何 TreeNode， SelectedNode 屬性為 null。
                {
                    if (!DOMTreeView.SelectedNode.IsEditing) DOMTreeView.SelectedNode.Expand();

                    int count = DOMTreeView.SelectedNode.Nodes.Count;   //取得指派給樹狀檢視控制項的樹狀節點集合的節點個數。
                    string tagname = DOMTreeView.SelectedNode.Text;
                    if (tagname.IndexOf("<") >= 0)      //"<" = Html 標記(Tag)的起始記號
                    {
                        if (count == 0)
                        {
                            DOMTreeView.SelectedNode.Nodes.Remove(DOMTreeView.SelectedNode);
                            DeleteDataArea1(DOMTreeView, DOMTreeView.SelectedNode);
                        }
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    //DOMTreeView.Select();
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region MarkFullPathColor
        /// <summary>
        /// 標記路徑顏色
        /// </summary>
        /// <param name="DOMTreeView1">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        private void MarkFullPathColor(TreeView DOMTreeView, TreeNode tree_node)
        {
            if (tree_node.Parent == null) return;

            tree_node.BackColor = Color.Red;
            MarkFullPathColor(DOMTreeView, DOMTreeView.SelectedNode);
        }
        #endregion

        #region DeletePrevDataArea
        /// <summary>
        /// 去除資料位置之前的區塊
        /// <param name="DOMTreeView">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        private void DeletePrevDataArea(TreeView DOMTreeView, TreeNode tree_node)
        {
            string strValue1 = "";
            string strValue2 = "";
            bool bolStart = false;
            try
            {
                string[] strFullPath1 = tree_node.FullPath.Split('\\');

                DOMTreeView.SelectedNode = DOMTreeView.Nodes[0].Nodes[0];
                //DOMTreeView.Select();
                DOMTreeView.Visible = false;

                while (DOMTreeView.SelectedNode != null)
                {
                    if (DOMTreeView.SelectedNode.FullPath == tree_node.FullPath)    //取得從根樹狀節點通往目前樹狀節點的路徑
                    {
                        goto ReadNext;
                    }

                    string[] strFullPath2 = DOMTreeView.SelectedNode.FullPath.Split('\\');  //取得這個執行個體中,由('\\')所分隔之子字串

                    int cnt = (int)DOMTreeView.SelectedNode.GetNodeCount(false);    //傳回子樹狀節點的數目

                    //取得第一次分支的節點
                    if (DOMTreeView.SelectedNode.Level < 4 && cnt > 1 && strValue1.Equals(""))
                    {
                        TreeNode temp_node1 = tree_node;
                        for (int i = 0; i < strFullPath1.Length - strFullPath2.Length; i++)
                        {
                            strValue1 = temp_node1.ImageKey;
                            if (temp_node1.Parent == null) break;
                            temp_node1 = temp_node1.Parent;
                        }
                        bolStart = true;
                    }

                    //取得第一次分支的節點之下一個子節點
                    if (!strValue1.Equals("") && !bolStart && strValue2.Equals(""))
                    {
                        strValue2 = DOMTreeView.SelectedNode.ImageKey;
                    }

                    //判斷是否為同一區塊
                    //由Body之後區分
                    if (strFullPath1.Length >= 2 && strFullPath2.Length >= 2)
                    {
                        if (strFullPath1[2].ToString() != strFullPath1[2].ToString())
                        {
                            //    DOMTreeView1.Nodes.Remove(DOMTreeView1.SelectedNode);
                        }
                        else
                        {
                            if (!strValue1.Equals("") && !strValue2.Equals("") && strValue1 != strValue2)
                            {
                                DOMTreeView.Nodes.Remove(DOMTreeView.SelectedNode);
                                break;
                            }
                        }
                    }

                    if (bolStart)   bolStart = !bolStart;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    DOMTreeView.Select();
                }

                ReadNext:
                    DOMTreeView.SelectedNode = tree_node;
                    //DOMTreeView.Select();
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region MarkDeleteNode
        /// <summary>
        /// 將可去除之不相關標記注記為紅色,不實際刪除之
        /// <param name="DOMTreeView2">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        public void MarkspiliterArea(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null)
                {
                    return;
                }
                else
                {
                    DOMTreeView.SelectedNode = tree_node;
                    DOMTreeView.Visible = false;    //關閉Tree_View
                    DOMTreeView.ExpandAll();
                }

                while (DOMTreeView.SelectedNode != null)
                {
                    string trfullpath1 = DOMTreeView.SelectedNode.FullPath; //紀錄TreeNode的路徑
                    int cnt = (int)DOMTreeView.SelectedNode.GetNodeCount(false);    //傳回子樹狀節點的數目
                    if (cnt == 1 && DOMTreeView .SelectedNode .Level >=  4 && DOMTreeView .SelectedNode .Text == "<DIV>")
                    {
                        DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region AreaTextprint
        /// <summary>
        /// 依據節點內文字的長度來刪除節點
        /// </summary>
        /// <param name="DOMTreeView"></param>
        /// <param name="tree_node"></param>
        Stack tagname = new Stack();
        public void AreaTextprint(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null) return;      //如果沒有選取任何 TreeNode

                DOMTreeView.SelectedNode = tree_node;   //設定目前在樹狀檢視控制項中選取的樹狀節點。
                DOMTreeView.Visible = false;
                DOMTreeView.Select();     // Select() 方法會啟動此控制項。
                while (DOMTreeView .SelectedNode .Nodes .Count >1)    //如果目前沒有選取任何 TreeNode， SelectedNode 屬性為 null。 
                {
                    tagname.Push(tree_node);
                    
                    
                    if (DOMTreeView.SelectedNode == null) return;   //如果目前沒有選取任何 TreeNode， SelectedNode 屬性為 null。

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    DOMTreeView.Select();
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion


        #region DeleteDataArea2
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DOMTreeView"></param>
        /// <param name="tree_node"></param>
        public void DeleteDataArea2(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null) return;      //如果沒有選取任何 TreeNode

                DOMTreeView.SelectedNode = tree_node;
                //DOMTreeView.Visible = false;
                //DOMTreeView.Select();
                while (DOMTreeView.SelectedNode != null)    //如果目前沒有選取任何 TreeNode， SelectedNode 屬性為 null。
                {
                    if (!DOMTreeView.SelectedNode.IsEditing) DOMTreeView.SelectedNode.Expand();

                    int count = DOMTreeView.SelectedNode.Nodes.Count;   //取得指派給樹狀檢視控制項的樹狀節點集合的節點個數。
                    string tagname = DOMTreeView.SelectedNode.Text;
                    if (tagname.IndexOf("<") >= 0)      //"<" = Html 標記(Tag)的起始記號
                    {
                        if (count == 0)
                        {
                            DOMTreeView.SelectedNode.Nodes.Remove(DOMTreeView.SelectedNode);
                            DeleteDataArea2(DOMTreeView, DOMTreeView.SelectedNode);
                        }
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                    //DOMTreeView.Select();
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region DeletNode2
        /// <summary>
        ///  將可去除之不相關標記實際刪除之
        /// <param name="DOMTreeView">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        public void DeleteNode2(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null)
                {
                    return;
                }
                else
                {
                    DOMTreeView.SelectedNode = tree_node;
                    DOMTreeView.Visible = false;
                    DOMTreeView.ExpandAll();
                }

                while (DOMTreeView.SelectedNode != null)
                {
                    //if (!DOMTreeView.SelectedNode.IsEditing) DOMTreeView.SelectedNode.Expand();

                    if (DOMTreeView.SelectedNode.Tag is mshtml .IHTMLDOMNode)
                    {
                        mshtml.IHTMLDOMNode em = DOMTreeView.SelectedNode.Tag as mshtml.IHTMLDOMNode;
                       
                        
                        //HtmlElement em = DOMTreeView.SelectedNode.Tag as HtmlElement;
                        if (em.nodeName  == "!" || em.nodeName == "SCRIPT" || em.nodeName == "STYLE" || em.nodeName == "NOSCRIPT" || em.nodeValue == null)
                        {
                            DOMTreeView.Nodes.Remove(DOMTreeView.SelectedNode);
                            DeleteNode2(DOMTreeView, DOMTreeView.SelectedNode);
                        }
                          
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;

                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region MarkDeleteNode2
        /// <summary>
        /// 將可去除之不相關標記注記為紅色,不實際刪除之
        /// <param name="DOMTreeView2">畫面物件</param>
        /// <param name="tree_node">dom com樹狀結構</param>
        /// <returns>無</returns>
        public void MarkDeleteNode2(TreeView DOMTreeView, TreeNode tree_node)
        {
            try
            {
                if (tree_node == null)
                {
                    return;
                }
                else
                {
                    DOMTreeView.SelectedNode = tree_node;
                    DOMTreeView.Visible = false;    //關閉Tree_View
                    DOMTreeView.ExpandAll();
                }

                while (DOMTreeView.SelectedNode != null)
                {

                    if (DOMTreeView.SelectedNode.Tag is mshtml.IHTMLDOMNode)
                    {
                        mshtml.IHTMLDOMNode em = DOMTreeView.SelectedNode.Tag as mshtml.IHTMLDOMNode;


                        //HtmlElement em = DOMTreeView.SelectedNode.Tag as HtmlElement;
                        if (em.nodeName == "!" || em.nodeName == "SCRIPT" || em.nodeName == "STYLE" || em.nodeName == "NOSCRIPT" || em.nodeValue == null|| em.nodeName =="BR")
                        {
                            DOMTreeView.SelectedNode.BackColor = Color.Red;
                        }
                        if (DOMTreeView.SelectedNode.Text.IndexOf("<") != -1)
                        {
                            if (DOMTreeView.SelectedNode.Nodes.Count == 0)
                            {
                                DOMTreeView.SelectedNode.BackColor = Color.Red;
                            }
                        }
                    }

                    if (DOMTreeView.SelectedNode == null) return;

                    DOMTreeView.SelectedNode = DOMTreeView.SelectedNode.NextVisibleNode;
                }
                DOMTreeView.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

    }
}
