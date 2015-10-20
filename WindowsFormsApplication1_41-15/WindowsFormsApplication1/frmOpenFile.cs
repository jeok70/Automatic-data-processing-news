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
        #region �ܼ�
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
        /// �e������ظm�l
        /// </summary>
        public frmOpenFile()
        {
            InitializeComponent();
        }
        #endregion

        #region frmOpenFile_Load
        /// <summary>
        /// �e�����J
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmOpenFile_Load(object sender, EventArgs e)
        {
            listView1.Clear();
            listView1.View = View.Details;

            //���Y���]�w
            ColumnHeader columnheader;

            columnheader = new ColumnHeader();
            columnheader.Text = "�Ǹ�";
            columnheader.Width = 50;
            columnheader.TextAlign = HorizontalAlignment.Center;
            listView1.Columns.Add(columnheader);

            columnheader = new ColumnHeader();
            columnheader.Text = "�ɮצW��";
            columnheader.Width = 500;
            columnheader.TextAlign = HorizontalAlignment.Left;
            listView1.Columns.Add(columnheader);

            columnheader = new ColumnHeader();
            columnheader.Text = "�j�p";
            columnheader.Width = 150;
            columnheader.TextAlign = HorizontalAlignment.Right;
            listView1.Columns.Add(columnheader);

            PopulateTreeView();
        }
        #endregion

        #region frmOpenFile_FormClosed
        /// <summary>
        /// �e������
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
        /// �ظm�ڥؿ�
        /// </summary>
        /// <param name="strPath"></param>
        private void PopulateTreeView()
        {
            treeView1.Nodes.Clear();

            DirectoryInfo DirInfo = new DirectoryInfo(strCurrentPath);

            if (DirInfo.Exists)
            {
                TreeNode rootNode = new TreeNode("���");
                GetDirectories(DirInfo.GetDirectories(), rootNode);

                treeView1.Nodes.Add(rootNode);
                treeView1.ExpandAll();

                txtShowData.Text = "";
            }
        }
        #endregion

        #region GetDirectories-local dir
        /// <summary>
        /// �ظm�l�ؿ�
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
        /// ��ܥؿ��ɱa�X�ɮ�
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Text != "���")
            {
                //���o����r
                strSelectDir = e.Node.Text;

                //���o�ثe����ؿ�
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
        /// �O�_�R�������ؿ�
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (treeView1.SelectedNode == null) return;

            //���o�ثe����O��
            TreeNode tvi = treeView1.SelectedNode;

            //���o����r
            strSelectDir = tvi.Text;

            //���o�ثe����ؿ�
            string MainDir = strCurrentPath + strSelectDir;

            if (Directory.Exists(MainDir))
            {
                if (e.KeyCode == Keys.Delete)
                {
                    strMessage = String.Format("�O�_�R�������ؿ��H");
                    DialogResult DlogRut = MessageBox.Show(strMessage, "ĵ�i�T��", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (DlogRut == DialogResult.No) return;

                    //�R���ؿ�
                    Directory.Delete(MainDir, true);

                    //�����`�I
                    treeView1.Nodes.Remove(tvi);

                    //�M��listView1����
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
                    //���s��z
                    PopulateTreeView();
                }
            }
        }
        #endregion

        #region listView1_SelectedIndexChanged
        /// <summary>
        /// ����ɮ׮ɱa�X���e
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            //���o�ثe����O��
            ListViewItem lvi = listView1.SelectedItems[0];

            //���o�ثe����ؿ�
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
                            if (line.StartsWith("��"))
                            {
                                line = line.Trim("��".ToCharArray());
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
        /// �O�_�R�������ɮ�
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            //���o�ثe����O��
            ListViewItem lvi = listView1.SelectedItems[0];

            //���o�ثe����ؿ�
            string MainDir = strCurrentPath + strSelectDir;

            DirectoryInfo Dir = new DirectoryInfo(MainDir);
            FileInfo[] totFile = Dir.GetFiles(lvi.SubItems[1].Text);

            if (totFile.Length != 0)
            {
                foreach (FileInfo subFile in totFile)
                {
                    if (e.KeyCode == Keys.Delete)
                    {
                        strMessage = String.Format("�O�_�R�������ɮסH");
                        DialogResult DlogRut = MessageBox.Show(strMessage, "ĵ�i�T��", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (DlogRut == DialogResult.No) return;

                        //�R���ɮ�
                        subFile.Delete();

                        //�����`�I
                        lvi.Remove();

                        txtShowData.Text = "";
                    }
                }
            }
        }
        #endregion

        bool blclick =false;//���L�ܼ� ���s�޿�
        SpeechSynthesizer synh = new SpeechSynthesizer(); //speech����ŧi Text to Voice
        private void btn_speak_Click(object sender, EventArgs e)
        {
            if (blclick)
            {
                synh.SpeakAsyncCancelAll();
                btn_speak.Text = "����s�D";
            }
            else
            {
                btn_speak.Text = "�����";
                synh.SpeakAsync(txtShowData.Text);
            }
            blclick = !blclick;
        }
    }
}