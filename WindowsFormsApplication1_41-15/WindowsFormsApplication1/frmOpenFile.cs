using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Speech.Synthesis;

namespace WindowsFormsApplication1
{
    public partial class frmOpenFile : Form
    {
        #region 變數
        private String strMessage = "";
        private static frmOpenFile instance;
        private static readonly string strCurrentPath = AppDomain.CurrentDomain.BaseDirectory;
        private String strSelectDir = "";
        #endregion

        #region Instance
        public static frmOpenFile Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new frmOpenFile();
                }
                instance.Activate();
                return instance;
            }
        }
        #endregion

        #region frmDataImport
        /// <summary>
        /// 畫面元件建置子
        /// </summary>
        public frmOpenFile()
        {
            InitializeComponent();
        }
        #endregion

        #region frmOpenFile_Load
        /// <summary>
        /// 畫面載入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmOpenFile_Load(object sender, EventArgs e)
        {
            listView1.Clear();
            listView1.View = View.Details;

            //表頭欄位設定
            ColumnHeader columnheader;

            columnheader = new ColumnHeader();
            columnheader.Text = "序號";
            columnheader.Width = 50;
            columnheader.TextAlign = HorizontalAlignment.Center;
            listView1.Columns.Add(columnheader);

            columnheader = new ColumnHeader();
            columnheader.Text = "檔案名稱";
            columnheader.Width = 500;
            columnheader.TextAlign = HorizontalAlignment.Left;
            listView1.Columns.Add(columnheader);

            columnheader = new ColumnHeader();
            columnheader.Text = "大小";
            columnheader.Width = 150;
            columnheader.TextAlign = HorizontalAlignment.Right;
            listView1.Columns.Add(columnheader);

            PopulateTreeView();
        }
        #endregion

        #region frmOpenFile_FormClosed
        /// <summary>
        /// 畫面關閉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmOpenFile_FormClosed(object sender, FormClosedEventArgs e)
        {
            instance = null;
        }
        #endregion

        #region PopulateTreeView-local dir
        /// <summary>
        /// 建置根目錄
        /// </summary>
        /// <param name="strPath"></param>
        private void PopulateTreeView()
        {
            treeView1.Nodes.Clear();

            DirectoryInfo DirInfo = new DirectoryInfo(strCurrentPath);

            if (DirInfo.Exists)
            {
                TreeNode rootNode = new TreeNode("資料");
                GetDirectories(DirInfo.GetDirectories(), rootNode);

                treeView1.Nodes.Add(rootNode);
                treeView1.ExpandAll();

                txtShowData.Text = "";
            }
        }
        #endregion

        #region GetDirectories-local dir
        /// <summary>
        /// 建置子目錄
        /// </summary>
        /// <param name="subDirs"></param>
        /// <param name="nodeToAddTo"></param>
        private void GetDirectories(DirectoryInfo[] subDirs, TreeNode subNode)
        {
            TreeNode tempNode;

            DirectoryInfo[] subSubDirs;

            foreach (DirectoryInfo subDir in subDirs)
            {
                tempNode = new TreeNode(subDir.Name, 0, 0);
                tempNode.Tag = subDir.FullName;
                subSubDirs = subDir.GetDirectories();

                if (subSubDirs.Length != 0)
                {
                    GetDirectories(subSubDirs, tempNode);
                }
                subNode.Nodes.Add(tempNode);
            }
        }
        #endregion

        #region treeView1_NodeMouseClick
        /// <summary>
        /// 選擇目錄時帶出檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Text != "資料")
            {
                //取得的文字
                strSelectDir = e.Node.Text;

                //取得目前選取目錄
                string MainDir = strCurrentPath + strSelectDir;

                if (Directory.Exists(MainDir))
                {
                    DirectoryInfo Dir = new DirectoryInfo(MainDir);

                    FileInfo[] totFile = Dir.GetFiles();

                    if (totFile.Length != 0)
                    {
                        int i = 1;
                        listView1.Items.Clear();

                        foreach (FileInfo subFile in totFile)
                        {
                            if (subFile.Name != "data.txt")
                            {
                                ListViewItem listviewitem;

                                listviewitem = new ListViewItem(i.ToString());
                                listviewitem.SubItems.Add(subFile.Name);
                                listviewitem.SubItems.Add(subFile.Length.ToString());
                                listView1.Items.Add(listviewitem);

                                i++;
                            }
                        }
                    }
                }
            }
        }
        #endregion

        #region treeView1_KeyDown
        /// <summary>
        /// 是否刪除此筆目錄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (treeView1.SelectedNode == null) return;

            //取得目前選取記錄
            TreeNode tvi = treeView1.SelectedNode;

            //取得的文字
            strSelectDir = tvi.Text;

            //取得目前選取目錄
            string MainDir = strCurrentPath + strSelectDir;

            if (Directory.Exists(MainDir))
            {
                if (e.KeyCode == Keys.Delete)
                {
                    strMessage = String.Format("是否刪除此筆目錄？");
                    DialogResult DlogRut = MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (DlogRut == DialogResult.No) return;

                    //刪除目錄
                    Directory.Delete(MainDir, true);

                    //移除節點
                    treeView1.Nodes.Remove(tvi);

                    //清除listView1明細
                    for (int i = listView1.Items.Count - 1; i >= 0; i--)
                    {
                        listView1.Items[i].Remove();
                    }

                    txtShowData.Text = "";
                }
            }
            else
            {
                if (e.KeyCode == Keys.F5)
                {
                    //重新整理
                    PopulateTreeView();
                }
            }
        }
        #endregion

        #region listView1_SelectedIndexChanged
        /// <summary>
        /// 選擇檔案時帶出內容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            //取得目前選取記錄
            ListViewItem lvi = listView1.SelectedItems[0];

            //取得目前選取目錄
            string MainDir = strCurrentPath + strSelectDir;
            StringBuilder SB = new StringBuilder();
            DirectoryInfo Dir = new DirectoryInfo(MainDir);
            FileInfo[] totFile = Dir.GetFiles(lvi.SubItems[1].Text);
            string line = null;
            if (totFile.Length != 0)
            {
                foreach (FileInfo subFile in totFile)
                {
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
                        txtShowData.Text = SB.ToString();
                        sr.Close();
                }

            }
        }
        #endregion

        #region listView1_KeyDown
        /// <summary>
        /// 是否刪除此筆檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            //取得目前選取記錄
            ListViewItem lvi = listView1.SelectedItems[0];

            //取得目前選取目錄
            string MainDir = strCurrentPath + strSelectDir;

            DirectoryInfo Dir = new DirectoryInfo(MainDir);
            FileInfo[] totFile = Dir.GetFiles(lvi.SubItems[1].Text);

            if (totFile.Length != 0)
            {
                foreach (FileInfo subFile in totFile)
                {
                    if (e.KeyCode == Keys.Delete)
                    {
                        strMessage = String.Format("是否刪除此筆檔案？");
                        DialogResult DlogRut = MessageBox.Show(strMessage, "警告訊息", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (DlogRut == DialogResult.No) return;

                        //刪除檔案
                        subFile.Delete();

                        //移除節點
                        lvi.Remove();

                        txtShowData.Text = "";
                    }
                }
            }
        }
        #endregion

        bool blclick =false;//布林變數 按鈕邏輯
        SpeechSynthesizer synh = new SpeechSynthesizer(); //speech物件宣告 Text to Voice
        private void btn_speak_Click(object sender, EventArgs e)
        {
            if (blclick)
            {
                synh.SpeakAsyncCancelAll();
                btn_speak.Text = "播放新聞";
            }
            else
            {
                btn_speak.Text = "停止播放";
                synh.SpeakAsync(txtShowData.Text);
            }
            blclick = !blclick;
        }
    }
}