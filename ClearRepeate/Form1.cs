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
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using RemoveRepeat;

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
                        /*
                        foreach(var picture in myDocx.AllPackagePictures) {
                            Console.WriteLine(picture);
                        }
                        */
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
                                    if (list.Exists(n => n.Contains(compareStr.Trim())))
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
                    list.ForEach(n => Console.WriteLine(n));
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
        #region  比较字符串相似度
        private static int min(int one, int two, int three)
        {
            int min = one;
            if (two < min)
            {
                min = two;
            }
            if (three < min)
            {
                min = three;
            }
            return min;
        }
        public static int LD(String str1, String str2)
        {
            int[,] d;     // 矩阵 
            int n = str1.Length;
            int m = str2.Length;
            int i;     // 遍历str1的 
            int j;     // 遍历str2的 
            char ch1;     // str1的 
            char ch2;     // str2的 
            int temp;     // 记录相同字符,在某个矩阵位置值的增量,不是0就是1 
            if (n == 0)
            {
                return m;
            }
            if (m == 0)
            {
                return n;
            }
            d = new int[n + 1, m + 1];
            for (i = 0; i <= n; i++)
            {     // 初始化第一列 
                d[i, 0] = i;
            }
            for (j = 0; j <= m; j++)
            {     // 初始化第一行 
                d[0, j] = j;
            }
            for (i = 1; i <= n; i++)
            {     // 遍历str1 
                ch1 = str1[i - 1];
                // 去匹配str2 
                for (j = 1; j <= m; j++)
                {
                    ch2 = str2[j - 1];
                    if (ch1 == ch2)
                    {
                        temp = 0;
                    }
                    else
                    {
                        temp = 1;
                    }
                    // 左边+1,上边+1, 左上角+temp取最小 
                    d[i, j] = min(d[i - 1, j] + 1, d[i, j - 1] + 1, d[i - 1, j - 1] + temp);
                }
            }
            return d[n, m];
        }

        //返回两个字符串的相似度，返回一个0到100之间的整数，值越大，表示相似度越高
        public static int similar(String newStr, String targetStr)
        {
            int ld = LD(newStr, targetStr);
            double i = 1 - (double)ld / (double)Math.Max(newStr.Length, targetStr.Length);
            int similar = Convert.ToInt32(Math.Round((Convert.ToDecimal(i)), 2, MidpointRounding.AwayFromZero) * 100);
            return similar;
        }
        #endregion

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            //IDataObject iData = Clipboard.GetDataObject();
            Console.WriteLine(Clipboard.ContainsImage());
            Console.WriteLine(Clipboard.ContainsText());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                IDataObject iData = Clipboard.GetDataObject();
                var allData = iData.GetData(DataFormats.Html);
                allData = Regex.Replace(allData.ToString(), @"\<(?!img.*?).*?\>", "");
                //Console.WriteLine(allData.ToString());
                var splitStr = Regex.Split(allData.ToString(), @"(?=\d?、.*?)");
                List<Quiz> Quiz = new List<Quiz>();
                foreach (string str in splitStr)
                {
                    //Console.WriteLine(str);
                    if (Regex.IsMatch(str, @"\d?、?"))
                    {
                        //添加到列表
                        Quiz quiz = new Quiz
                        {
                            Title = str
                        };
                        Regex regImg = new Regex(@"src=""(?<imgUrl>.*?)""", RegexOptions.IgnoreCase);
                        // 搜索匹配的字符串 
                        MatchCollection matches = regImg.Matches(str);
                        int i = 0;
                        string[] sUrlList = new string[matches.Count];
                        // 取得匹配项列表 
                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches)
                            {
                                if (!string.IsNullOrEmpty(match.Groups["imgUrl"].Value))
                                {
                                    sUrlList[i++] = match.Groups["imgUrl"].Value;
                                }
                            }
                            quiz.Picture = sUrlList;
                            Quiz.Add(quiz);
                        }
                    }
                }
                Quiz.ForEach(n=>Console.WriteLine(n.Picture[0]));
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
            
        }
        
    }
}
