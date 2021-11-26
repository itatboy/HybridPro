using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.IO;
using System.IO.Ports;
using System.Drawing.Drawing2D;

namespace HybridPro
{
    public struct SerialOption
    {
        public static string[] baud = { "4800", "9600", "14400", "19200", "38400", "115200", "128000", "230400", "256000", "460800", "500000" };
        public static string[] dataBits = { "5", "6", "7", "8" };
        public static string[] stopBits = { "0", "1", "2" };
        public static string[] parityBits = { "None", "Odd", "Even", "Mark", "Space" };

        public static int baudIndex;
        public static int dataBitsIndex;
        public static int stopBitsIndex;
        public static int parityBitsIndex;
    };

    public partial class HybridPro : Form
    {
        SerialPort com = new SerialPort();
        StringBuilder sb = new StringBuilder();
        private DateTime current_time = new DateTime();
        private bool bNeedTime = true;
        bool bComFlag = false;
        bool bTopFlag = false;
        bool bSendFlag = false;
        private long ulRecvCount = 0; // 接收字节计数
        private long ulSendCount = 0; // 发送字节计数

        /* 实时保存 */
        bool bRealTimeSaveFlag = false; // 实时保存功能标志位
        string strFoldPath; // 实时保存文件路径
        String strFileName; // 实时保存文件名称

        public HybridPro()
        {
            InitializeComponent();
        }

        /*
         * 加载
         * */
        private void Form1_Load(object sender, EventArgs e)
        {
            SerialOption.baudIndex = 9;
            SerialOption.dataBitsIndex = 3;
            SerialOption.stopBitsIndex = 1;
            SerialOption.parityBitsIndex = 0;

            /* 显示串口列表 */
            comboBox1.Items.AddRange(ComGetInfo());
            comboBox1.SelectedIndex = 0;

            /* 1s刷新一下串口列表 */
            timer2.Interval = 1000;
            timer2.Start();
        }

        /*
         * 窗口快捷键
         * */
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            /* 串口打开 */
            if (bComFlag)
            {
                if ((e.KeyCode == Keys.S) && e.Control)
                {
                    usart_save();
                }
            }
        }

        /*
         * 菜单栏
         * 置顶
         * */
        private void topToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bTopFlag = !bTopFlag;
            if (bTopFlag)
                TopMost = true;
            else
                TopMost = false;
        }

        /*
        * 菜单栏
        * 发送
        * */
        private void sendToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bSendFlag = !bSendFlag;
            if (bSendFlag)
                this.Width = 1136;
            else
                this.Width = 713;
        }

        /*
        * 菜单栏
        * 计算器
        * */
        private void calculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("calc.exe");
            }
            catch { }
        }

        /*
         * 串口
         * 获取串口号和设备描述
         * */
        public static string[] ComGetInfo()
        {
            List<string> Comstrs = new List<string>();
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("select * from Win32_PnPEntity where Name like '%(COM%'"))
                {
                    var hardInfos = searcher.Get();
                    foreach (var comInfo in hardInfos)
                    {
                        if (comInfo.Properties["Name"].Value.ToString().Contains("COM")) // 在得到的串口列表里，搜索包含COM关键字的串口
                        {
                            string coms = comInfo.Properties["Name"].Value.ToString();
                            string[] strcom = coms.Split(new char[2] { '(', ')' });
                            Comstrs.Add(strcom[1] + "  " + strcom[0]);
                        }
                    }
                    searcher.Dispose();
                }
                return Comstrs.ToArray();
            }
            catch
            {
                return null;
            }
            finally
            { Comstrs = null; }
        }

        /*
         * 串口
         * 配置串口
         * */
        private void button1_Click(object sender, EventArgs e)
        {
            serial OptionDialog = new serial();
            OptionDialog.ShowDialog(this);
        }

        /*
         * 串口
         * 打开或关闭串口
         * */
        private void button2_Click(object sender, EventArgs e)
        {
            if (!bComFlag)
            {
                string str = comboBox1.Text;
                string[] strArray = str.Split(' ');
                com.PortName = strArray[0];
                com.BaudRate = int.Parse(SerialOption.baud[SerialOption.baudIndex]);
                com.DataBits = Convert.ToInt32(SerialOption.dataBits[SerialOption.dataBitsIndex]);
                com.StopBits = (StopBits)Convert.ToInt32(SerialOption.stopBits[SerialOption.stopBitsIndex]);
                com.Parity = (Parity)Convert.ToInt32(SerialOption.parityBitsIndex.ToString());
                com.ReceivedBytesThreshold = 1;
                com.DataReceived += new SerialDataReceivedEventHandler(SerialPort1_DataReceived);
                com.ReadTimeout = 500;
                com.WriteTimeout = 500;
                try
                {
                    com.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                bComFlag = com.IsOpen ? true : false;
                if (bComFlag)
                {
                    button2.ForeColor = Color.Red;
                    button2.Text = "关闭";
                    comboBox1.Enabled = false;
                    button1.Enabled = false;
                }
            }
            else
            {
                try
                {
                    com.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                bComFlag = com.IsOpen ? true : false;
                if (!bComFlag)
                {
                    button2.ForeColor = Color.Black;
                    button2.Text = "打开";
                    comboBox1.Enabled = true;
                    button1.Enabled = true;
                }
            }
        }

        /*
        * 清空接收
        * */
        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            ulRecvCount = 0;
            label8.Text = "接收：" + ulRecvCount.ToString();
        }

        /*
        * 保存
        * */
        private void button4_Click(object sender, EventArgs e)
        {
            usart_save();
        }

        /*
        * 保存
        * */
        private void usart_save()
        {
            String fileName;
            string foldPath;

            /* 获取当前接收区内容 */
            String recv_data = textBox1.Text;

            if (recv_data.Equals(""))
            {
                MessageBox.Show("接收数据为空");
                return;
            }

            /* 弹出文件夹选择框供用户选择 */
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择日志文件存储路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                foldPath = dialog.SelectedPath;
            }
            else
            {
                return;
            }

            fileName = foldPath + "\\log_" + System.DateTime.Now.ToString("yyMMdd_HHmmss") + ".txt";

            try
            {
                FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                byte[] bytes = Encoding.UTF8.GetBytes(recv_data); // 将字符串转换为字节数组
                fs.Write(bytes, 0, bytes.Length); // 向文件中写入字节数组
                fs.Flush(); // 刷新缓冲区
                fs.Close(); // 关闭流
                MessageBox.Show("日志已保存!(" + fileName + ")");
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生异常!(" + ex.ToString() + ")");
            }
        }

        /*
         * 串口数据接收
         * */
        private void SerialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            /* 刷新定时器 */
            if (checkBox2.Checked)
            {
                timer1.Interval = (int)domainUpDown1.Value;
            }

            int num = com.BytesToRead;
            byte[] received_buf = new byte[num];
            ulRecvCount += num;

            com.Read(received_buf, 0, num);
            if (num < 1)
                return;

            sb.Clear();

            if (checkBox1.Checked)
            {
                foreach (byte b in received_buf)
                {
                    sb.Append(b.ToString("X2") + ' ');
                }
            }
            else
            {
                sb.Append(Encoding.ASCII.GetString(received_buf));
            }

            try
            {
                //因为要访问UI资源，所以需要使用invoke方式同步ui
                Invoke((EventHandler)(delegate
                {
                    if (bNeedTime && checkBox2.Checked)
                    {
                        /* 需要加时间戳 */
                        bNeedTime = false;   //清空标志位
                        current_time = System.DateTime.Now;     //获取当前时间
                        textBox1.AppendText("[" + current_time.ToString("HH:mm:ss") + "]" + " " + " " + sb.ToString());
                    }
                    else
                    {
                        /* 不需要时间戳 */
                        textBox1.AppendText(sb.ToString());
                    }

                    if (checkBox6.Checked && bRealTimeSaveFlag)
                    {
                        StreamWriter fs = File.AppendText(strFoldPath + strFileName);
                        fs.Write(sb.ToString());
                        fs.Close();
                    }

                    label8.Text = "接收：" + ulRecvCount.ToString();
                }));
            }
            catch { }
        }

        /*
         * 显示时间
         * */
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                /* 启动定时器 */
                domainUpDown1.Enabled = false;
                timer1.Interval = (int)domainUpDown1.Value;
                timer1.Start();
            }
            else
            {
                /* 取消时间戳，停止定时器 */
                domainUpDown1.Enabled = true;
                timer1.Stop();
            }
        }

        /*
         * 实时保存
         * */
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                /* 弹出文件夹选择框供用户选择 */
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.Description = "请选择日志文件存储路径";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    strFoldPath = dialog.SelectedPath;
                }
                else
                {
                    return;
                }

                strFileName = "\\log_" + System.DateTime.Now.ToString("yyMMdd_HHmmss") + ".txt";

                FileStream fs = File.Create(strFoldPath + strFileName); // 创建文件
                fs.Close();
                bRealTimeSaveFlag = true;
            }
            else
            {
                bRealTimeSaveFlag = false;
            }
        }

        /*
         * 定时器1事件
         * */
        private void timer1_Tick(object sender, EventArgs e)
        {
            bNeedTime = true; // 设置时间内未收到数据，分包
        }

        private bool search_port_is_exist(String item, String[] port_list)
        {
            for (int i = 0; i < port_list.Length; i++)
            {
                if (port_list[i].Equals(item))
                {
                    return true;
                }
            }

            return false;
        }

        /*
         * 定时器2事件
         * 自动刷新串口列表
         * */
        private void timer2_Tick(object sender, EventArgs e)
        {
            if (!bComFlag)
            {
                String[] strPortList = ComGetInfo();
                int count = comboBox1.Items.Count;

                /* combox中无内容，将当前串口列表全部加入 */
                if (count == 0)
                {
                    comboBox1.Items.AddRange(strPortList);
                    return;
                }
                else
                {
                    //判断有无新插入的串口
                    for (int i = 0; i < strPortList.Length; i++)
                    {
                        if (!comboBox1.Items.Contains(strPortList[i]))
                        {
                            //找到新插入串口，添加到combox中
                            comboBox1.Items.Add(strPortList[i]);
                        }
                    }

                    if (count != strPortList.Length)
                    {
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(strPortList);
                        comboBox1.SelectedIndex = 0;
                    }
                }
            }
        }
    }
}
