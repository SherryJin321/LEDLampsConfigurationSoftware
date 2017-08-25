using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;

namespace TestLampsProperty
{
    public partial class Form1 : Form
    {
        //定义全局变量
        SerialPort lampsSerialPort = new SerialPort();

        /// <summary>
        /// 灯具串口
        /// </summary>
        Byte[] queryStatusCommand = new Byte[28] { 0x02, 0x89, 0x11, 0x58, 0x12, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00 };
        Byte[] queryVersionCommand = new Byte[28] { 0x02, 0x89, 0x22, 0x85, 0x12, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00 };


        byte[] bt_sensor_readDate = new byte[5] { 0X02, 0X55, 0X33, 0X00, 0X00 };                       //传感器读取数据指令

        public Form1()
        {
            InitializeComponent();
        }

        //启动定时器
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";

            if (lampsSerialPort.IsOpen == false)
            {
                lampsSerialPort.Open();
            }
            //timer1.Start();
            if (lampsSerialPort.IsOpen)
            {
                //lampsSerialPort.Write(queryStatusCommand, 0, 28);

                //textBox1.Text += "TX:";
                //for (int i = 0; i < queryStatusCommand.Length; i++)
                //{

                //    if (i < queryStatusCommand.Length - 1)
                //    {
                //        textBox1.Text += Convert.ToString(queryStatusCommand[i], 16).PadLeft(2, '0').ToUpper() + " ";
                //    }
                //    else
                //    {
                //        textBox1.Text += Convert.ToString(queryStatusCommand[i], 16).PadLeft(2, '0').ToUpper() + "\r\n";
                //    }
                //}

                lampsSerialPort.Write(queryVersionCommand, 0, 28);

                textBox1.Text += "TX:";
                for (int i = 0; i < queryVersionCommand.Length; i++)
                {

                    if (i < queryVersionCommand.Length - 1)
                    {
                        textBox1.Text += Convert.ToString(queryVersionCommand[i], 16).PadLeft(2, '0').ToUpper() + "-";
                    }
                    else
                    {
                        textBox1.Text += Convert.ToString(queryVersionCommand[i], 16).PadLeft(2, '0').ToUpper() + "\r\n";
                    }
                }
            }

        }

        //设置串口参数
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lampsSerialPort.PortName = comboBox1.SelectedItem.ToString();
            lampsSerialPort.StopBits = StopBits.One;
            lampsSerialPort.DataBits = 8;
            lampsSerialPort.Parity = Parity.None;
            lampsSerialPort.BaudRate = 921600;

            //lampsSerialPort.PortName = comboBox1.SelectedItem.ToString();
            //lampsSerialPort.StopBits = StopBits.One;
            //lampsSerialPort.DataBits = 8;
            //lampsSerialPort.Parity = Parity.None;
            //lampsSerialPort.BaudRate = 38400;

            lampsSerialPort.DataReceived += new SerialDataReceivedEventHandler(lampsSerialPortDataReceived);
        }

        //开机启动
        private void Form1_Load(object sender, EventArgs e)
        {
            string[] portsName = SerialPort.GetPortNames();
            comboBox1.DataSource = portsName;
            queryStatusCommand[27] = CalculateCheckOutValue(queryStatusCommand);
            queryVersionCommand[27] = CalculateCheckOutValue(queryVersionCommand);
            textBox1.Text = "";
            CheckForIllegalCrossThreadCalls = false;                                              //解决线程操作无效
            //bt_sensor_readDate[bt_sensor_readDate.Length - 1] = CalculateCheckOutValue(bt_sensor_readDate);     //传感器读取数据指令增加校验字节          
        }

        //生成校验字节
        public byte CalculateCheckOutValue(byte[] CommandNeedAddCheckOutValue)
        {
            byte CheckOutValue = 0X00;                                     //定义校验字节变量

            for (int i = 0; i < (CommandNeedAddCheckOutValue.Length - 1); i++)
            {
                CheckOutValue += CommandNeedAddCheckOutValue[i];                        //检验字节=字节数组所有字节求和，取低8位
            }

            return CheckOutValue;                                          //返回校验字节
        }

        //发送指令
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (lampsSerialPort.IsOpen)
            {
                lampsSerialPort.Write(bt_sensor_readDate, 0, 5);

                textBox1.Text += "TX:";
                for (int i = 0; i < bt_sensor_readDate.Length; i++)
                {

                    if (i < bt_sensor_readDate.Length - 1)
                    {
                        textBox1.Text += Convert.ToString(bt_sensor_readDate[i], 16).PadLeft(2, '0').ToUpper() + " ";
                    }
                    else
                    {
                        textBox1.Text += Convert.ToString(bt_sensor_readDate[i], 16).PadLeft(2, '0').ToUpper() + "\r\n";
                    }
                }
            }
            else
            {
                MessageBox.Show("未打开串口！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void lampsSerialPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (lampsSerialPort.IsOpen)                                                       //判断传感器串口是否打开
            {
                byte[] receivedData = new byte[lampsSerialPort.BytesToRead];           //创建接收字节数组
                lampsSerialPort.Read(receivedData, 0, receivedData.Length);     //读取数据      

                textBox1.Text += "RX:";
                if(receivedData.Length!=0)
                {
                    for (int i = 0; i < receivedData.Length; i++)
                    {
                        if (i < receivedData.Length - 1)
                        {
                            textBox1.Text += Convert.ToString(receivedData[i], 16).PadLeft(2, '0').ToUpper() + " ";
                        }
                        else
                        {
                            textBox1.Text += Convert.ToString(receivedData[i], 16).PadLeft(2, '0').ToUpper() + "\r\n";
                        }
                    }
                }
                else
                {
                    textBox1.Text += "\r\n";
                }
                
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(lampsSerialPort.IsOpen)
            {
                lampsSerialPort.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //timer1.Stop();
            if(lampsSerialPort.IsOpen)
            {
                lampsSerialPort.Close();
            }
        }
    }   
}
