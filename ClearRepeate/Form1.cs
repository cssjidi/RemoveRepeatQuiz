using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using NPOI.XWPF.UserModel;

namespace ClearRepeate
{
    public partial class Form1 : Form
    {
        private static List<string> list;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = this.openFileDialog1;
            dlg.Filter = "word文件|*.docx;";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                button1.Enabled = false;
                label1.Visible = true;
                button2.Enabled = false;
                button4.Enabled = false;
                string fileName = dlg.FileName;
                FileInfo info = new FileInfo(fileName);
                if (info.Length / 1024 > 2048)
                {
                    MessageBox.Show("上传文件不能超过2M");
                    return;
                }
                try
                {
                    FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                    if (fileName.Contains("docx"))
                    {
                        XWPFDocument myDocx = new XWPFDocument(fs);
                        list = new List<string>();
                        string compareStr = "";
                        bool similar = false;
                        int index = 0;
                        foreach (var para in myDocx.Paragraphs)
                        {

                            string strSence = para.ParagraphText;
                            if (strSence.Contains("我的答案") || strSence.Contains("得分"))
                            {
                                continue;
                            }
                            else
                            {
                                if (Regex.IsMatch(strSence, @"^(\d)+、?"))
                                {
                                    compareStr = Regex.Replace(strSence, @"^(\d)+、?", "");
                                    similar = false;
                                    if (list.Exists(n => n.Contains(compareStr)))
                                    {
                                        similar = true;
                                        compareStr = "";
                                    }
                                    else
                                    {
                                        index++;
                                        list.Add(index.ToString() + "、" + compareStr);
                                    }
                                }
                                else if (similar)
                                {
                                    continue;
                                }
                                else
                                {
                                    list.Add(strSence);
                                }
                            }
                        }
                    }
                    
                    fs.Close();
                    fs.Dispose();
                    button3.Enabled = true;
                    label1.Visible = false;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button4.Enabled = true;
                    textBox1.Enabled = true;
                    viewInTextBox();
                }
                catch (Exception ex)
                {
                    button3.Enabled = true;
                    label1.Visible = false;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button4.Enabled = true;
                    textBox1.Enabled = true;
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void viewInTextBox() {
            string str = "";
            list.ForEach(n => {
                str += n + "\r\n";
            });
            textBox1.ScrollToCaret();
            textBox1.Text = str;
        }

        private void SaveFileAsWrod() {
            var doc = new XWPFDocument();
            list.ForEach(n => {
                var p1 = doc.CreateParagraph();
                p1.Alignment = ParagraphAlignment.LEFT;
                var runTitle = p1.CreateRun();
                if (Regex.IsMatch(n, @"^(\d)+、?"))
                {
                    runTitle.FontSize = 14;
                    runTitle.SetFontFamily("微软雅黑", FontCharRange.None);
                    runTitle.SetText(n.Trim() + "\r\n");
                } else if (n.Contains("参考答案"))
                {
                    runTitle.SetColor("#f00");
                    runTitle.FontSize = 14;
                    runTitle.SetText(n + "\r\n\r\n");
                }
                else if (!n.Equals("\r\n")) {
                     runTitle.SetText(n + "\r\n");
                }
            });
            //runTitle.FontSize = 12;
            //runTitle.SetFontFamily("微软雅黑", FontCharRange.None);
            //var ms = new MemoryStream();
            //doc.Write(ms);
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = DateTimeKind.Local.ToString();
            sfd.Filter = "Word Document(*.docx)|*.docx";
            sfd.DefaultExt = "Word Document(*.docx)|*.docx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                button1.Enabled = false;
                button3.Enabled = false;
                FileStream fs = (FileStream)sfd.OpenFile();
                doc.Write(fs);
                fs.Close();
                doc.Close();
                doc = null;
                fs = null;
            }
            MessageBox.Show("保存成功");
            button1.Enabled = true;
            button3.Enabled = true;
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileAsWrod();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            printDialog1.ShowDialog();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Visible = true;
            this.BackColor = Color.Transparent;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("点击上传需要处理的word文档，点击保存保存去重后的文档");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本软件只授权给华师网络学院学生使用，如商用或除华师网络学院学生外的使用的，请联系cjd2015@qq.com进行授权，特此声明");
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
