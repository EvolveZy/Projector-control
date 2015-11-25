using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Net.Sockets;
using System.Threading;
using System.Diagnostics;



namespace CloverControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]

        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                RegistryKey location = Registry.LocalMachine;
                RegistryKey soft = location.OpenSubKey("SOFTWARE", false);
                RegistryKey myPass = soft.OpenSubKey("FTLiang", false);
                if (myPass.GetValue("com").ToString().Trim() == null)
                {
                    Module.com = "COM1";
                }
                else
                {
                    Module.com = "COM" + myPass.GetValue("com").ToString().Trim();
                }
                if (myPass.GetValue("ServerIP").ToString().Trim() == null)
                {
                    Module.serverip = "100.120.0.26";
                }
                else
                {
                    Module.serverip = myPass.GetValue("ServerIP").ToString().Trim();
                }

                Module.xinhao = myPass.GetValue("xinhao").ToString().Trim();
            }
            catch (Exception)
            {
                return;
            }
            CheckForIllegalCrossThreadCalls = false;
            string pathUrl = System.Windows.Forms.Application.ExecutablePath;//程序完整路径
        }

   

        public static bool TestConnection(string host, int port, int millisecondsTimeout)
        {
            TcpClient client = new TcpClient();
            try
            {
                var ar = client.BeginConnect(host, port, null, null);
                ar.AsyncWaitHandle.WaitOne(millisecondsTimeout);
                return client.Connected;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                client.Close();
            }
        } 

        public bool CheckFtp()//验证FTP端口是否可连接
        {
            bool ResultValue = true;
            try
            {
                string strIP = Module.serverip;

                //发送数据，判断是否连接到指定ip 
                if (!TestConnection(strIP, 21, 500))
                {
                    //MessageBox.Show("与服务器通信失败！", "连接服务器");
                    //this.Close();
                   ResultValue=false;
                }
               
            }
            catch
            {
                ResultValue = false;
                
            }
            return ResultValue;
           
        }

        //监测关机关投影
        internal void SystemEvents_SessionEnding(object sender, SessionEndingEventArgs e)
        {
                int i = Convert.ToInt32(Module.xinhao);
                SerialPort serialPort = new SerialPort();
                serialPort.PortName = Module.com;
                if (0 == i)//SONY
                {
                    serialPort.BaudRate = 38400;
                    serialPort.Parity = Parity.Even;//偶
                    serialPort.DataBits = 8;
                    serialPort.Open();
                    serialPort.Write(Module.bufferclose0, 0, Module.bufferclose0.Length);
                    serialPort.Write(Module.bufferclose0, 0, Module.bufferclose0.Length);
                }
                if (1 == i)//NEC-H
                {
                    serialPort.BaudRate = 19200;
                    serialPort.Parity = Parity.None;//无
                    serialPort.DataBits = 8;
                    serialPort.Open();
                    serialPort.Write(Module.bufferclose1, 0, Module.bufferclose1.Length);
                    serialPort.Write(Module.bufferclose1, 0, Module.bufferclose1.Length);
                }
                if (2 == i)//NEC-VGA
                {
                    serialPort.BaudRate = 19200;
                    serialPort.Parity = Parity.None;//无
                    serialPort.DataBits = 8;
                    serialPort.Open();
                    serialPort.Write(Module.bufferclose2, 0, Module.bufferclose2.Length);
                    serialPort.Write(Module.bufferclose2, 0, Module.bufferclose2.Length);
                }
                if (3 == i)//NP-M361X
                {
                    serialPort.BaudRate = 19200;
                    serialPort.Parity = Parity.None;//无
                    serialPort.DataBits = 8;
                    serialPort.Open();
                    serialPort.Write(Module.bufferclose3, 0, Module.bufferclose3.Length);
                    serialPort.Write(Module.bufferclose3, 0, Module.bufferclose3.Length);
                }
                if (4 == i)//激光
                {
                    serialPort.BaudRate = 115200;
                    serialPort.Parity = Parity.None;//无
                    serialPort.DataBits = 8;
                    serialPort.Open();
                    serialPort.Write(Module.bufferclose4, 0, Module.bufferclose4.Length);
                    serialPort.Write(Module.bufferclose4, 0, Module.bufferclose4.Length);
                }
                serialPort.Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            SystemEvents.SessionEnding -= new SessionEndingEventHandler(this.SystemEvents_SessionEnding);
            base.OnClosed(e);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //开投影机
        private void button1_Click(object sender, EventArgs e)
        {
            

            if (Module.xinhao == null)
            {
                MessageBox.Show("请设定投影机型号！");
                return;
            }
            else
            {
                try
                {
                    int i = Convert.ToInt32(Module.xinhao);
                    SerialPort serialPort = new SerialPort();
                    serialPort.PortName = Module.com;
                    if (0 == i)//SONY-VGA
                    {
                        serialPort.BaudRate = 38400;
                        serialPort.Parity = Parity.Even;//偶
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferopen0, 0, Module.bufferopen0.Length);
                        System.Threading.Thread.Sleep(60 * 1000);
                        serialPort.Write(Module.bufferv0, 0, Module.bufferv0.Length);
                    }
                    if (1 == i)//NEC-H
                    {
                        serialPort.BaudRate = 19200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferopen1, 0, Module.bufferopen1.Length);
                        System.Threading.Thread.Sleep(60 * 1000);
                        serialPort.Write(Module.bufferv1, 0, Module.bufferv1.Length);
                    }
                    if (2 == i)//NEC-VGA
                    {
                        serialPort.BaudRate = 19200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferopen2, 0, Module.bufferopen2.Length);
                        System.Threading.Thread.Sleep(60 * 1000);
                        serialPort.Write(Module.bufferv2, 0, Module.bufferv2.Length);
                    }
                    if (3 == i)//NP361-H
                    {
                        serialPort.BaudRate = 19200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferopen3, 0, Module.bufferopen3.Length);
                        System.Threading.Thread.Sleep(60 * 1000);
                        serialPort.Write(Module.bufferv3, 0, Module.bufferv3.Length);
                    }
                    if (4 == i)//激光-H2
                    {
                        serialPort.BaudRate = 115200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferopen4, 0, Module.bufferopen4.Length);
                        System.Threading.Thread.Sleep(60 * 1000);
                        serialPort.Write(Module.bufferv4, 0, Module.bufferv4.Length);
                    }
                    serialPort.Close();
                }
                catch
                {
                    MessageBox.Show("计算机" + Module.com + "不存在！");
                }
            }
        }


        //关投影机
        private void button2_Click(object sender, EventArgs e)
        {

            if (Module.xinhao == null)
            {
                MessageBox.Show("请设定投影机型号！");
                return;
            }
            else
            {
                try
                {
                    int i = Convert.ToInt32(Module.xinhao);
                    SerialPort serialPort = new SerialPort();
                    serialPort.PortName = Module.com;
                    if (0 == i)//SONY-VGA
                    {
                        serialPort.BaudRate = 38400;
                        serialPort.Parity = Parity.Even;//偶
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferclose0, 0, Module.bufferclose0.Length);
                        serialPort.Write(Module.bufferclose0, 0, Module.bufferclose0.Length);
                    }
                    if (1 == i)//NEC-H
                    {
                        serialPort.BaudRate = 19200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferclose1, 0, Module.bufferclose1.Length);
                        serialPort.Write(Module.bufferclose1, 0, Module.bufferclose1.Length);
                    }
                    if (2 == i)//NEC-VGA
                    {
                        serialPort.BaudRate = 19200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferclose2, 0, Module.bufferclose2.Length);
                        serialPort.Write(Module.bufferclose2, 0, Module.bufferclose2.Length);
                    }
                    if (3 == i)//NP-M361X-H1
                    {
                        serialPort.BaudRate = 19200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferclose3, 0, Module.bufferclose3.Length);
                        serialPort.Write(Module.bufferclose3, 0, Module.bufferclose3.Length);
                    }
                    if (4 == i)//激光-H2
                    {
                        serialPort.BaudRate = 115200;
                        serialPort.Parity = Parity.None;//无
                        serialPort.DataBits = 8;
                        serialPort.Open();
                        serialPort.Write(Module.bufferclose4, 0, Module.bufferclose4.Length);
                        serialPort.Write(Module.bufferclose4, 0, Module.bufferclose4.Length);
                    }
                    serialPort.Close();
                }
                catch
                {
                    MessageBox.Show("计算机" + Module.com + "不存在！");
                }
            }
        }

        //定时获取服务器壁纸
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (CheckFtp() == true)
            {
                huanbizi();
            }
            else
            {
                return;
            }
        }

        //换壁纸
        public void ChangeRegedit()
        {
            try
            {
                Image image = Image.FromFile("C:\\Windows\\clover.jpg");
                image.Save("C:\\Windows\\TranscodedWallpaper.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
                SystemParametersInfo(20, 0, "C:\\Windows\\TranscodedWallpaper.bmp", 0x2);
            }
            catch (Exception)
            {
                return;
            }

        }

        //下载壁纸
        public void huanbizi()
        {
            FtpWebRequest reqFTP;
            try
            {
                FileStream outputStream = new FileStream(@"C:\Windows\clover.jpg", FileMode.Create);

                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://" + Module.serverip + "/clover.jpg"));

                reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;

                reqFTP.UseBinary = true;

                reqFTP.Credentials = new NetworkCredential("clover", "cloverccniit123+-");

                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();

                Stream ftpStream = response.GetResponseStream();

                long cl = response.ContentLength;

                int bufferSize = 2048;

                int readCount;

                byte[] buffer = new byte[bufferSize];

                readCount = ftpStream.Read(buffer, 0, bufferSize);

                while (readCount > 0)
                {
                    outputStream.Write(buffer, 0, readCount);

                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                }

                ftpStream.Close();

                outputStream.Close();

                response.Close();
                ChangeRegedit();
            }
            catch (Exception)
            {
                // timer1.Enabled = false;
                return;
            }
        }
           
        //禁止关闭
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
          if (e.CloseReason == CloseReason.UserClosing)
          {
              e.Cancel = true;    //取消"关闭窗口"事件
              this.WindowState = FormWindowState.Minimized;    //使关闭时窗口向右下角缩小的效果
              notifyIcon1.Visible = true;
              this.Hide();
              return;
          }

        }

        //
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }

        //双击显示
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            notifyIcon1.Visible = false;
            this.Show();
            WindowState = FormWindowState.Normal;
            this.Focus();

        }


        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F && e.Control)
            {
                Form yz = new yz();
                yz.ShowDialog();
            }
          //  if (e.KeyCode == Keys.B && e.Control)
           // {
           //     Form msg = new msg();
           //     msg.ShowDialog();
          //  }

        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            label2.Focus();
        }

         
    }
}
