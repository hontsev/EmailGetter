using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Text.RegularExpressions;

using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

using OpenPop.Pop3;
using OpenPop.Mime;

namespace EmailGetter2
{
    public enum SystemState { exit,login,working,endwork};
    public partial class Form1 : Form
    {

        public TcpClient Server;
        public NetworkStream NetStrm;
        public StreamReader RdStrm;
        public string Data;
        public byte[] szData;
        public string CRLF = "\r\n";

        public int emailNum = 0;
        public List<EmailInfo> emails;
        SetDate setDateWindow;
        public DateTime selectDate1,selectDate2;

        private SystemState state = SystemState.exit;

        int lEmailNo = 0;
        int rEmailNo = 0;

        string excelFileName = "";

        delegate void sendStringDelegate(string str);
        delegate void sendVoidDelegate();
        delegate void sendStateDelegate(SystemState state);
        sendStringDelegate printstr;
        sendVoidDelegate showlist;
        sendStateDelegate updateState;

        private string userName;
        private string userPassword;
        private string emailServerName;

        Pop3Client emailClient;


        public Form1()
        {
            InitializeComponent();
            
            button2.Enabled = false;
            groupBox2.Enabled = false;
            groupBox3.Enabled = false;
            emails = new List<EmailInfo>();
            setDateWindow = new SetDate();
            setDateWindow.Owner = this;

            printstr = new sendStringDelegate(printl);
            showlist = new sendVoidDelegate(updateEmailList);
            updateState = new sendStateDelegate(updateUIState);
            emailClient = new Pop3Client();
        }

        /// <summary>
        /// 初始化邮箱地址、密码等数据
        /// </summary>
        private void initEmailAddressInfo()
        {
            if (textBox3.Text.Length <= 0 || textBox4.Text.Length <= 0 || textBox3.Text.Split('@').Length < 2)
            {
                printl("邮箱信息不合法，无法访问");
                return;
            }

            emailServerName = "pop3." + textBox3.Text.Split('@')[1];
            userName = textBox3.Text.Split('@')[0];
            userPassword = textBox4.Text;

            textBox5.Text = "";
        }

        /// <summary>
        /// 更新收取邮件的日期上下限
        /// </summary>
        public void updateDates()
        {
            try
            {
                if(setDateWindow.dateNo==1)
                {
                    textBox2.Text = setDateWindow.dt.ToString("yyyy年MM月dd日");
                }
                else if(setDateWindow.dateNo==2)
                {
                    textBox6.Text = setDateWindow.dt.ToString("yyyy年MM月dd日");
                }
            }
            catch (Exception e)
            {
                Invoke(printstr, (object)(e.Message));
            }

        }

        /// <summary>
        /// 根据是否登录，更新界面按钮等的状态
        /// </summary>
        private void updateUIState(SystemState state)
        {
            this.state = state;
            switch (state)
            {
                case SystemState.exit:
                    button1.Enabled = true;
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button6.Enabled = false;

                    groupBox1.Enabled = true;
                    groupBox2.Enabled = false;
                    groupBox3.Enabled = false;
                    
                    emails.Clear();
                    textBox2.Clear();
                    textBox6.Clear();
                    listBox1.Items.Clear();
                    textBox1.Clear();
                    break;
                case SystemState.login:
                    button1.Enabled = false;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button6.Enabled = false;

                    groupBox1.Enabled = true;
                    groupBox2.Enabled = true;
                    groupBox3.Enabled = false;

                    emails.Clear();
                    textBox2.Clear();
                    textBox6.Clear();
                    listBox1.Items.Clear();
                    textBox1.Clear();
                    break;
                case SystemState.working:
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button6.Enabled = false;

                    groupBox1.Enabled = false;
                    groupBox2.Enabled = false;
                    groupBox3.Enabled = false;
                    break;
                case SystemState.endwork:
                    button1.Enabled = false;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button6.Enabled = true;

                    groupBox1.Enabled = true;
                    groupBox2.Enabled = true;
                    groupBox3.Enabled = true;
                    break;
                default: break;
            }
        }

        /// <summary>
        /// 打印邮件正文
        /// </summary>
        /// <param name="email"></param>
        private void printEmail(EmailInfo email)
        {
            textBox1.Text = "";
            textBox1.Text += ("主题:" + email.Subject + CRLF);
            textBox1.Text += ("发信人:" + email.From + CRLF);
            textBox1.Text += ("收信人:" + email.To + CRLF);
            textBox1.Text += ("抄送:" + email.Cc + CRLF);
            textBox1.Text += ("日期:" + email.Date + CRLF);
            textBox1.Text += ("正文:" + email.Content + CRLF);
        }

        /// <summary>
        /// 输出debug信息
        /// </summary>
        /// <param name="str"></param>
        private void printl(string str)
        {
            textBox5.AppendText(str + CRLF);
        }

        /// <summary>
        /// 更新邮件列表内容
        /// </summary>
        private void updateEmailList()
        {
            listBox1.Items.Clear();
            foreach (EmailInfo message in emails)
            {
                listBox1.Items.Add(message.Subject);
            }
        }


        /// <summary>
        /// 登陆邮箱
        /// </summary>
        private void workLogin()
        {
            //连接邮箱服务器
            Invoke(printstr, (object)("正在登陆邮箱…"));
            try
            {
                emailClient = new Pop3Client();
                if (emailClient.Connected)
                {
                    emailClient.Disconnect();
                }
                emailClient.Connect(emailServerName, 110, false);
                emailClient.Authenticate(userName, userPassword);
                emailNum = emailClient.GetMessageCount();
                Invoke(printstr, (object)("登陆成功，邮箱里有 " + emailNum + " 封邮件。"));
                Invoke(updateState,(object)SystemState.login);
            }
            catch (Exception e)
            {
                Invoke(printstr, (object)("登陆失败，错误码：" + e.Message));
                Invoke(updateState, (object)SystemState.exit);
            }
        }

        /// <summary>
        /// 获取限定范围内的邮件内容
        /// </summary>
        private void workGetEmails()
        {
            Invoke(updateState, (object)SystemState.working);
            Invoke(printstr, (object)("开始计算指定日期内邮件编号范围，从" + selectDate1.ToShortDateString() + "到" + selectDate2.ToShortDateString()));
            try
            {
                //二分法定好邮件编号上下限
                getEmailNoRange();

                emails.Clear();
                Invoke(printstr, (object)("开始获取邮件，从" + lEmailNo + "到" + rEmailNo));
                for (int i = lEmailNo; i <= rEmailNo; i++)
                {
                    try
                    {
                        emails.Add(getEmailInfo(i));
                        Invoke(printstr, (object)("第" + (i + 1 - lEmailNo) + "封，共" + (rEmailNo - lEmailNo + 1) + "封"));
                    }
                    catch
                    {
                        Invoke(printstr, (object)("第" + (i + 1 - lEmailNo) + "封出错！，共" + (rEmailNo - lEmailNo + 1) + "封"));
                    }
                }
            }
            catch (Exception e)
            {
                Invoke(printstr, (object)("出错：" + e.Message));
            }
            Invoke(printstr, (object)("邮件获取完毕[" + DateTime.Now.ToShortTimeString() + "]"));
            Invoke(showlist);
            Invoke(updateState, (object)SystemState.endwork);
        }

        /// <summary>
        /// 汉字转换为Unicode编码
        /// </summary>
        /// <param name="str">要编码的汉字字符串</param>
        /// <returns>Unicode编码的的字符串</returns>
        public string ToUnicode(string str)
        {
            byte[] bts = Encoding.Default.GetBytes(str);
            string r = Encoding.GetEncoding("UTF-8").GetString(bts);
            return r;
        }

        /// <summary>
        /// 将邮件列表内容存入excel
        /// </summary>
        private void workOutputExcel()
        {
            Invoke(updateState, (object)SystemState.working);
            Invoke(printstr, (object)("开始输出邮件数据到excel：" + excelFileName));
            try
            {
                string excelPath = excelFileName.Substring(0, excelFileName.LastIndexOf("\\")) + "\\";
                string attPath = excelPath + @"邮件附件\";
                if (!Directory.Exists(attPath)) Directory.CreateDirectory(attPath);
                Invoke(printstr, (object)("附件内容输出到：" + attPath));
                IWorkbook wb = new HSSFWorkbook();
                ICellStyle style1 = wb.CreateCellStyle();
                IFont font1 = wb.CreateFont();
                font1.Underline = FontUnderlineType.None;
                font1.Color = HSSFColor.Black.Index;
                style1.SetFont(font1);
                ICellStyle style2 = wb.CreateCellStyle();
                IFont font2 = wb.CreateFont();
                font2.Underline = FontUnderlineType.Single;
                font2.Color = HSSFColor.Blue.Index;
                style2.SetFont(font2);
                string[] strHead = new string[] { "收信日期", "标题", "发信人", "收信人", "抄送", "正文", "附件" };
                int[] columnWidth = new int[] { 21, 40, 15, 15, 15, 30, 30 };
                ISheet sheet1 = wb.CreateSheet();
                var headRow = sheet1.CreateRow(0);
                for (int i = 0; i < columnWidth.Length; i++)
                {
                    var tcell=headRow.CreateCell(i);
                    tcell.SetCellValue(strHead[i]);
                    tcell.CellStyle = style1;
                    sheet1.SetColumnWidth(i, columnWidth[i] * 256);
                }
                OpenPop.Mime.Message message;
                for (int i = 0; i < emails.Count; i++)
                {
                    //向excel文件中添加项
                    var contentRow = sheet1.CreateRow(i + 1);
                    ICell tcell;
                    tcell = contentRow.CreateCell(0);
                    tcell.SetCellValue(emails[i].Date);
                    tcell.CellStyle = style1;
                    tcell = contentRow.CreateCell(1);
                    tcell.SetCellValue(emails[i].Subject);
                    tcell.CellStyle = style1;
                    tcell = contentRow.CreateCell(2);
                    tcell.SetCellValue(emails[i].From);
                    tcell.CellStyle = style1;
                    tcell = contentRow.CreateCell(3);
                    tcell.SetCellValue(emails[i].To);
                    tcell.CellStyle = style1;
                    tcell = contentRow.CreateCell(4);
                    tcell.SetCellValue(emails[i].Cc);
                    tcell.CellStyle = style1;
                    string tmp = emails[i].Content;
                    if (tmp.Length > 1000)
                    {
                        tmp = tmp.Substring(0, 1000) + "****内容未完****";
                    }
                    tcell = contentRow.CreateCell(5);
                    tcell.SetCellValue(tmp);
                    tcell.CellStyle = style1;

                    //存储附件
                    if (emailClient.GetMessageSize(emails[i].no) > 10 * 1024 * 1024)
                    {
                        //邮件大于10m不接收附件
                        tcell = contentRow.CreateCell(6);
                        tcell.SetCellValue("****附件过大，请去邮箱自行查看****");
                        tcell.CellStyle = style1;
                        Invoke(printstr, (object)(String.Format("保存第 {0} 封邮件 ,附件过大无法下载。", i + 1)));
                    }
                    else
                    {
                        message = emailClient.GetMessage(emails[i].no);
                        List<MessagePart> attachments = message.FindAllAttachments();
                        int attno = 0;
                        foreach (MessagePart attachment in attachments)
                        {
                            attno++;
                            string filename = attPath + "[" + emails[i].no + attachments.Count + "] - " + attachment.FileName;
                            FileInfo fi = new FileInfo(filename);
                            attachment.Save(fi);
                            HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.Url);
                            link.Address = filename;
                            //link.Address = ToUnicode(filename);
                            ICell cell = contentRow.CreateCell(5 + attno, CellType.Blank);
                            cell.SetCellValue(attachment.FileName);
                            cell.Hyperlink = link;
                            cell.CellStyle = style2;
                        }
                        if (attno == 0)
                        {
                            //无附件
                            tcell = contentRow.CreateCell(6);
                            tcell.SetCellValue("无");
                            tcell.CellStyle = style1;
                        }
                        Invoke(printstr, (object)(String.Format("保存第 {0} 封邮件 , {1} 个附件", i + 1, attachments.Count)));
                    }

                        

                }
                FileStream file = new FileStream(excelFileName, FileMode.Create, FileAccess.ReadWrite);
                wb.Write(file);
                file.Close();
                file.Dispose();
                Invoke(printstr, (object)("excel文件输出完毕。请查看：" + excelFileName));
                Invoke(updateState, (object)SystemState.endwork);
            }
            catch (Exception e)
            {
                Invoke(printstr, (object)("出错，错误码：" + e.Message));
                Invoke(updateState, (object)SystemState.endwork);
            }
        }


        #region 邮件读取

        /// <summary>
        /// 将传入日期与选定的日期上下限比较。返回整数值表示是否合适，-1表示小于下限，0表示在限定区间内，1表示高于上限。
        /// </summary>
        /// <param name="timestr"></param>
        /// <returns>整数值表示是否合适，-1表示小于下限，0表示在限定区间内，1表示高于上限。</returns>
        private int compareDate(DateTime dt)
        {
            int ok = 0;
            /*
            DateTime dt;
            if (!DateTime.TryParse(timestr, out dt))
            {
                System.Globalization.DateTimeFormatInfo g = new System.Globalization.DateTimeFormatInfo();
                timestr = timestr.Replace(",", "");
                string[] cx = timestr.Split(' ');
                g.LongDatePattern = "d MMMM yyyy";
                dt = DateTime.Parse(string.Format("{0} {1} {2} {3}", cx[1], cx[2], cx[3], cx[4]), g);
            }
            //Invoke(printstr, (object)(timestr + "=>" + dt.ToString()));
            //DateTime dt = DateTime.ParseExact(timestr, "yyyy年MM月dd日HH:mm:ss", new System.Globalization.CultureInfo("zh-CN", true));
             */
            if (dt < selectDate1) ok = -1;
            else if (dt > selectDate2) ok = 1;
            return ok;
        }

        /// <summary>
        /// 二分法获取邮件序号的上下限
        /// </summary>
        private void getEmailNoRange()
        {
            lEmailNo = 1;
            rEmailNo = emailNum;
            int left = lEmailNo;
            int right = rEmailNo;
            int mid = (left + right) / 2;
            while (true)
            {
                if (compareDate(getEmailDate(mid)) == -1)
                {
                    left = mid;
                    mid = (mid + right) / 2;
                }
                else
                {
                    right = mid;
                    mid = (left + mid) / 2;
                }
                if (left == mid)
                {
                    lEmailNo = right;
                    break;
                }
            }
            left = lEmailNo;
            right = rEmailNo;
            mid = (left + right) / 2;
            while (true)
            {
                if (compareDate(getEmailDate(mid)) == 0)
                {
                    left = mid;
                    mid = (mid + right) / 2;
                }
                else
                {
                    right = mid;
                    mid = (left + mid) / 2;
                }
                if (left == mid)
                {
                    rEmailNo = left;
                    break;
                }
            }
            if (lEmailNo <= 0) lEmailNo = 1;
            if (rEmailNo <= 0) rEmailNo = 1;
            //Invoke(printstr,(object)("开始获取从 " + getEmail(lEmailNo).Headers.Date + " 至 " + getEmail(rEmailNo).Headers.Date + " 之间的信件。共 " + (rEmailNo - lEmailNo) + "封。请稍等"));
        }

        /// <summary>
        /// 获取某个编号的邮件的日期
        /// </summary>
        /// <param name="no"></param>
        /// <returns></returns>
        private DateTime getEmailDate(int no)
        {
            OpenPop.Mime.Header.MessageHeader head = emailClient.GetMessageHeaders(no);
            return head.DateSent;
        }

        /// <summary>
        /// 获取某个编号的那封邮件，返回EmailInfo对象
        /// </summary>
        /// <param name="no"></param>
        /// <returns></returns>
        private EmailInfo getEmailInfo(int no)
        {
            OpenPop.Mime.Header.MessageHeader head = emailClient.GetMessageHeaders(no);
            //OpenPop.Mime.Message message = emailClient.GetMessage(no);
            EmailInfo emailinfo = new EmailInfo();
            emailinfo.From = head.From.Raw;
            foreach (var toAdd in head.To)
            {
                emailinfo.To += toAdd.Raw + ",";
            }
            foreach (var ccAdd in head.Cc)
            {
                emailinfo.Cc += ccAdd.Raw + ",";
            }
            emailinfo.Subject = head.Subject;
            emailinfo.Date = head.DateSent.ToString();
            emailinfo.no = no;

            string content = "";
            int size=emailClient.GetMessageSize(no);
            if (size > 10 * 1024 * 1024)
            {
                emailinfo.Content = "****邮件过大，无法获取。请到邮箱中查看此邮件！****";
            }
            else
            {
                OpenPop.Mime.Message message = emailClient.GetMessage(no);
                if (message.MessagePart.IsText)
                {
                    content = message.MessagePart.GetBodyAsText();
                }
                else if (message.MessagePart.IsMultiPart)
                {
                    MessagePart plainTextPart = message.FindFirstPlainTextVersion();
                    if (plainTextPart != null)
                    {
                        // The message had a text/plain version - show that one
                        content = plainTextPart.GetBodyAsText();
                    }
                    else
                    {
                        // Try to find a body to show in some of the other text versions
                        //List<MessagePart> textVersions = message.FindAllTextVersions();
                        //if (textVersions.Count >= 1)
                        //    content = textVersions[0].GetBodyAsText();
                        //else
                        //    content = "";
                    }
                }
                emailinfo.Content = content;
            }
            return emailinfo;
        }

        #endregion



        private void button1_Click(object sender, EventArgs e)
        {
            initEmailAddressInfo();
            new Thread(workLogin).Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (emailClient.Connected)
            {
                emailClient.Disconnect();
            }
            textBox5.Text = "";
            printl("连接断开。");
            Invoke(updateState, (object)SystemState.exit);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //收信

            textBox1.Clear();
            if (selectDate1 == null || selectDate2 == null || selectDate1 > selectDate2 || selectDate1.Year < 1950 || selectDate2.Year < 1950)
            {
                printl("日期范围错误，无法获取邮件");
                return;
            }

            //开始根据编号上下限获取邮件内容
            new Thread(workGetEmails).Start();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            setDateWindow.dateNo = 1;
            setDateWindow.Show();
            this.Enabled = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            setDateWindow.dateNo = 2;
            setDateWindow.Show();
            this.Enabled = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //save
            //saveFileDialog1.InitialDirectory = Application.StartupPath;
            saveFileDialog1.FileName = "邮件输出";
            saveFileDialog1.ShowDialog();
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedIndex<0 || listBox1.SelectedIndex >= emails.Count) return;
            textBox1.Clear();
            EmailInfo emailinfo = emails[listBox1.SelectedIndex];
            printEmail(emailinfo);
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            excelFileName = saveFileDialog1.FileName;
            new Thread(workOutputExcel).Start();
        }
    }
}
