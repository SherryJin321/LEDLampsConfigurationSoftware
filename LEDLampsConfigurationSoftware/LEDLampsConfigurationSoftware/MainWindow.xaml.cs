using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.Collections;
using Microsoft.Office.Interop.Excel;

namespace LEDLampsConfigurationSoftware
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        #region 设置全局变量
        SerialPort lampsPort = new SerialPort();  //定义串口
        bool isQueryStatus = false;  //定义状态查询标识符
        ArrayList ReceivedStatusFeedbackCommand = new ArrayList();  //定义接收到的状态反馈指令
        byte[] receivedStatusFeedbackCommand;  //将集合转换成字节数组做处理
        DateTime StartQueryStatus;
        TimeSpan QueryStatusTimeSpan;
        string userName;
        string password;


        //Excel文件保存
        string str_fileName;                                                  //定义变量Excel文件名
        Microsoft.Office.Interop.Excel.Application ExcelApp;                  //声明Excel应用程序
        Workbook ExcelDoc;                                                    //声明工作簿
        Worksheet ExcelSheet;                                                 //声明工作表
        #endregion

        #region 版本反馈指令参数
        int year = 0;
        int month = 0;
        int date = 0;
        int hardwareVersion = 0;
        int versionBigNumber = 0;
        int versionSmallNumber = 0;
        #endregion

        #region 设置参数指令参数
        byte[] settingIA = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] settingIB = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] settingIIA = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] settingIIB = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte settingReadRFlag = 0x00;  //读取电阻
        byte settingMosFlag = 0x00;  //不开MOSFET
        byte settingBreakFlag = 0x00;  //不带开路
        byte settingLampsNumber = 0x00;  //灯具编号
        #endregion

        #region 发送指令集（尚未计算校验值）
        Byte[] queryStatusCommand = new Byte[28] { 0x02,0x89,0x11,0x58,0x12,0x00,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00 };
        Byte[] queryVersionCommand = new Byte[28] { 0x02, 0x89, 0x22, 0x85, 0x12, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00 };
        Byte[] settingParameterCommand = new Byte[28] { 0x02, 0x55, 0x11, 0x58, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

        #endregion

        #region 8寸灯具各项参数存储集合
        ArrayList RMS1Eightinches = new ArrayList();
        ArrayList Val2Eightinches = new ArrayList();
        ArrayList Val3Eightinches = new ArrayList();
        ArrayList RMSEightinches = new ArrayList();
        ArrayList CurrentRatio1Eightinches = new ArrayList();
        ArrayList CurrentRatio3Eightinches = new ArrayList();
        ArrayList RESIAEightinches = new ArrayList();
        ArrayList RESIIAEightinches = new ArrayList();
        ArrayList SNSIAEightinches = new ArrayList();
        ArrayList SNSIIAEightinches = new ArrayList();
        ArrayList LEDF1Eightinches = new ArrayList();
        ArrayList TEightinches = new ArrayList();
        ArrayList SecondEightinches = new ArrayList();
        #endregion

        #region 12寸灯具各项参数存储集合
        ArrayList RMS1Twelveinches = new ArrayList();
        ArrayList RMS2Twelveinches = new ArrayList();
        ArrayList Val2Twelveinches = new ArrayList();
        ArrayList Val3Twelveinches = new ArrayList();
        ArrayList RMSMID1Twelveinches = new ArrayList();
        ArrayList RMSMID2Twelveinches = new ArrayList();
        ArrayList RMSTwelveinches = new ArrayList();
        ArrayList CurrentRatio1Twelveinches = new ArrayList();
        ArrayList CurrentRatio2Twelveinches = new ArrayList();
        ArrayList CurrentRatio3Twelveinches = new ArrayList();
        ArrayList CurrentRatio4Twelveinches = new ArrayList();
        ArrayList RESIATwelveinches = new ArrayList();
        ArrayList RESIBTwelveinches = new ArrayList();
        ArrayList RESIIATwelveinches = new ArrayList();
        ArrayList RESIIBTwelveinches = new ArrayList();
        ArrayList SNSIATwelveinches = new ArrayList();
        ArrayList SNSIBTwelveinches = new ArrayList();
        ArrayList SNSIIATwelveinches = new ArrayList();
        ArrayList SNSIIBTwelveinches = new ArrayList();
        ArrayList LEDF1Twelveinches = new ArrayList();
        ArrayList TTwelveinches = new ArrayList();
        ArrayList SecondTwelveinches = new ArrayList();
        #endregion

        public MainWindow()
        {                      
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string[] portNames = System.IO.Ports.SerialPort.GetPortNames();      //获取当前电脑所有串口号
            PortNameSelect.ItemsSource = portNames;    //将串口号显示在ComboBox
            PortNameSelect.SelectedIndex = portNames.Length - 1;
            queryVersionCommand[27] = CalculateCheckOutValue(queryVersionCommand);  //计算版本查询指令的校验值
            queryStatusCommand[27] = CalculateCheckOutValue(queryStatusCommand);  //计算状态查询指令的校验值                        
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(lampsPort.IsOpen==true)
            {
                lampsPort.Close();
            }
        }

        #region 串口操作
        //设置串口参数
        private void PortNameSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(lampsPort.IsOpen==false)
            {
                lampsPort.PortName = PortNameSelect.SelectedItem.ToString();               
            }
            if(lampsPort.IsOpen==true)
            {
                MessageBox.Show("串口已打开！请关闭串口操作", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            lampsPort.BaudRate = 921600;
            lampsPort.Parity = Parity.None;
            lampsPort.DataBits = 8;
            lampsPort.StopBits = StopBits.One;
            lampsPort.DataReceived += new SerialDataReceivedEventHandler(lampsPortDataReceived);
        }
        //打开串口
        private void OpenSerialPort_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lampsPort.IsOpen == false)
                {
                    lampsPort.Open();
                   
                }
                if(lampsPort.IsOpen==true)
                {
                    MessageBox.Show("串口已打开！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
           catch
            {
                MessageBox.Show("串口未打开！请选择正确的串口号", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        //关闭串口
        private void CloseSerialPort_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lampsPort.IsOpen == true)
                {
                    lampsPort.Close();                  
                }
                if(lampsPort.IsOpen==false)
                {
                    MessageBox.Show("串口已关闭！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch
            {
                MessageBox.Show("关闭串口失败！请重启软件", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region 生成校验字节
        public byte CalculateCheckOutValue(byte[] CommandNeedAddCheckOutValue)
        {
            byte CheckOutValue = 0X00;                                     //定义校验字节变量

            for (int i = 0; i < (CommandNeedAddCheckOutValue.Length - 1); i++)
            {
                CheckOutValue += CommandNeedAddCheckOutValue[i];                        //检验字节=字节数组所有字节求和，取低8位
            }

            return CheckOutValue;                                          //返回校验字节
        }
        #endregion

        #region 版本查询
        private void QueryVersion_Click(object sender, RoutedEventArgs e)
        {
            if(lampsPort.IsOpen)
            {
                isQueryStatus = false;
                AnswerVersion.Text = "";
                lampsPort.Write(queryVersionCommand, 0, 28);
            }
            else
            {
                MessageBox.Show("未打开串口！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }                       
        }
        #endregion

        #region 状态查询       
        private void QueryStatus_Click(object sender, RoutedEventArgs e)
        {
            if (lampsPort.IsOpen)
            {                
                isQueryStatus = true;
                lampsPort.Write(queryStatusCommand, 0, 28);
                StartQueryStatus = DateTime.Now;    
            }
            else
            {
                MessageBox.Show("未打开串口！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region 导出数据
        private void CreatExcel_Click(object sender, RoutedEventArgs e)
        {
            if(ReceivedStatusFeedbackCommand.Count!=0)
            {
                receivedStatusFeedbackCommand = new byte[ReceivedStatusFeedbackCommand.Count];

                for (int i = 0; i < receivedStatusFeedbackCommand.Length; i++)
                {
                    receivedStatusFeedbackCommand[i] = (byte)ReceivedStatusFeedbackCommand[i];
                }

                if (receivedStatusFeedbackCommand[0] == 0x02 && receivedStatusFeedbackCommand[1] == 0xAA && receivedStatusFeedbackCommand[2] == 0x01 && receivedStatusFeedbackCommand[3] == 0x0C && receivedStatusFeedbackCommand[4] == 0x0C)
                {
                    TwelveInchesDataAnalysis(receivedStatusFeedbackCommand);
                    TwelveInchesLampParametersCreatExcel();
                }
                else if (receivedStatusFeedbackCommand[0] == 0x02 && receivedStatusFeedbackCommand[1] == 0xAA && receivedStatusFeedbackCommand[2] == 0x08 && receivedStatusFeedbackCommand[3] == 0x08)
                {
                    EightInchesDataAnalysis(receivedStatusFeedbackCommand);
                    EightInchesLampParametersCreatExcel();
                }
                else
                {
                    MessageBox.Show("接收指令错误！请重新查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    ReceivedStatusFeedbackCommand.Clear();
                }
            }
            else
            {
                MessageBox.Show("接收指令不能为空！请先进行状态查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region 8寸灯具解析
        private void EightInchesDataAnalysis(byte[] CompleteData)
        {
            for(int i=0;i<CompleteData.Length/32;i++)
            {
                if(CompleteData[32*i]==0x02&& CompleteData[32*i+1]==0xAA&& CompleteData[32*i+2]==0x08&& CompleteData[32*i+3]==0x08)
                {
                    RMS1Eightinches.Add(CompleteData[32 * i + 5] * 1100);
                    Val2Eightinches.Add(CompleteData[32 * i + 6] * 20);
                    Val3Eightinches.Add(CompleteData[32 * i + 7]);
                    RMSEightinches.Add(CompleteData[32 * i + 8] * 4);
                    CurrentRatio1Eightinches.Add((float)(CompleteData[32 * i + 9] / 10.0));
                    CurrentRatio3Eightinches.Add((float)(CompleteData[32 * i + 10] / 10.0));
                    RESIAEightinches.Add(CompleteData[32 * i + 11] * 124);
                    RESIIAEightinches.Add(CompleteData[32 * i + 12] * 124);
                    SNSIAEightinches.Add(CompleteData[32 * i + 13] * 16);
                    SNSIIAEightinches.Add(CompleteData[32 * i + 14] * 16);
                    LEDF1Eightinches.Add(CompleteData[32 * i + 15]);
                    TEightinches.Add(CompleteData[32 * i + 16]);

                    int SecondResult = 0;
                    for (int j = 0; j < 4; j++)
                    {
                        int SecondOrigin = CompleteData[32*i+17+j];
                        SecondResult |= SecondOrigin;
                        if (j < 3)
                        {
                            SecondResult <<= 8;
                        }
                    }
                    SecondEightinches.Add(SecondResult);
                }
                else
                {
                    MessageBox.Show("接收指令错误！请重新查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    ClearEightInchesLampsParameter();
                }
            }
        }
        #endregion

        #region 12寸灯具解析
        private void TwelveInchesDataAnalysis(byte[] CompleteData)
        {
            for (int i = 0; i < CompleteData.Length / 32; i++)
            {
                if (CompleteData[32 * i] == 0x02 && CompleteData[32 * i + 1] == 0xAA && CompleteData[32 * i + 2] == 0x01 && CompleteData[32 * i + 3] == 0x0C&&CompleteData[32*i+4]==0x0C)
                {
                    RMS1Twelveinches.Add(CompleteData[32 * i + 5] * 500);
                    RMS2Twelveinches.Add(CompleteData[32 * i + 6] * 500);
                    Val2Twelveinches.Add(CompleteData[32 * i + 7] * 20);
                    Val3Twelveinches.Add(CompleteData[32 * i + 8]);
                    RMSMID1Twelveinches.Add(CompleteData[32 * i + 9] * 16);
                    RMSMID2Twelveinches.Add(CompleteData[32 * i + 10] * 16);
                    RMSTwelveinches.Add(CompleteData[32 * i + 11] * 4);
                    CurrentRatio1Twelveinches.Add((float)(CompleteData[32 * i + 12] / 10.0));
                    CurrentRatio2Twelveinches.Add((float)(CompleteData[32 * i + 13] / 10.0));
                    CurrentRatio3Twelveinches.Add((float)(CompleteData[32 * i + 14] / 10.0));
                    CurrentRatio4Twelveinches.Add((float)(CompleteData[32 * i + 15] / 10.0));
                    RESIATwelveinches.Add(CompleteData[32 * i + 16] * 124);
                    RESIBTwelveinches.Add(CompleteData[32 * i + 17] * 124);
                    RESIIATwelveinches.Add(CompleteData[32 * i + 18] * 124);
                    RESIIBTwelveinches.Add(CompleteData[32 * i + 19] * 124);
                    SNSIATwelveinches.Add(CompleteData[32 * i + 20] * 16);
                    SNSIBTwelveinches.Add(CompleteData[32 * i + 21] * 16);
                    SNSIIATwelveinches.Add(CompleteData[32 * i + 22] * 16);
                    SNSIIBTwelveinches.Add(CompleteData[32 * i + 23] * 16);
                    LEDF1Twelveinches.Add(CompleteData[32 * i + 24]);
                    TTwelveinches.Add(CompleteData[32 * i + 25]);

                    int SecondResult = 0;
                    for (int j = 0; j < 4; j++)
                    {
                        int SecondOrigin = CompleteData[32 * i + 26 + j];
                        SecondResult |= SecondOrigin;
                        if (j < 3)
                        {
                            SecondResult <<= 8;
                        }
                    }
                    SecondTwelveinches.Add(SecondResult);
                }
                else
                {
                    MessageBox.Show("12寸灯具接收指令错误！请重新查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    ClearTwelveInchesLampsParameters();
                }
            }
        }
        #endregion

        #region 8寸灯具数据保存至Excel
        void EightInchesLampParametersCreatExcel()
        {
            //创建excel模板
            str_fileName = "d:\\8寸灯具参数解析" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
            ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
            ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

            //设置Excel列名           
            ExcelSheet.Cells[1, 1] = "8寸灯具参数解析";
            ExcelSheet.Cells[2, 1] = "序号";
            ExcelSheet.Cells[2, 2] = "RMS1";
            ExcelSheet.Cells[2, 3] = "Val2";
            ExcelSheet.Cells[2, 4] = "Val3";
            ExcelSheet.Cells[2, 5] = "RMS";
            ExcelSheet.Cells[2, 6] = "Current_Ratio1";
            ExcelSheet.Cells[2, 7] = "Current_Ratio3";
            ExcelSheet.Cells[2, 8] = "RES_IA";
            ExcelSheet.Cells[2, 9] = "RES_IIA";
            ExcelSheet.Cells[2, 10] = "SNS_IA";
            ExcelSheet.Cells[2, 11] = "SNS_IIA";
            ExcelSheet.Cells[2, 12] = "LED_F1";
            ExcelSheet.Cells[2, 13] = "T";
            ExcelSheet.Cells[2, 14] = "Second";

            //输出各个参数值
            for (int i = 0; i < RMS1Eightinches.Count; i++)
            {
                ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                ExcelSheet.Cells[3 + i, 2] = RMS1Eightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 3] = Val2Eightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 4] = Val3Eightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 5] = RMSEightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 6] = CurrentRatio1Eightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 7] = CurrentRatio3Eightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 8] = RESIAEightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 9] = RESIIAEightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 10] = SNSIAEightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 11] = SNSIIAEightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 12] = LEDF1Eightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 13] = TEightinches[i].ToString();
                ExcelSheet.Cells[3 + i, 14] = ((int)SecondEightinches[i] / 3600).ToString() + ":" + (((int)SecondEightinches[i] % 3600) / 60).ToString() + ":" + (((int)SecondEightinches[i] % 3600) % 60).ToString();
            }

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序    

            ClearEightInchesLampsParameter();
            MessageBox.Show("数据已导出!保存至D盘的Excel文档", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        #endregion

        #region 12寸灯具数据保存至Excel
        void TwelveInchesLampParametersCreatExcel()
        {
            //创建excel模板
            str_fileName = "d:\\12寸灯具参数解析" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
            ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
            ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

            //设置Excel列名           
            ExcelSheet.Cells[1, 1] = "12寸灯具参数解析";
            ExcelSheet.Cells[2, 1] = "序号";
            ExcelSheet.Cells[2, 2] = "RMS1";
            ExcelSheet.Cells[2, 3] = "RMS2";
            ExcelSheet.Cells[2, 4] = "Val2";
            ExcelSheet.Cells[2, 5] = "Val3";
            ExcelSheet.Cells[2, 6] = "RMSMID1";
            ExcelSheet.Cells[2, 7] = "RMSMID2";
            ExcelSheet.Cells[2, 8] = "RMS";
            ExcelSheet.Cells[2, 9] = "Current_Ratio1";
            ExcelSheet.Cells[2, 10] = "Current_Ratio2";
            ExcelSheet.Cells[2, 11] = "Current_Ratio3";
            ExcelSheet.Cells[2, 12] = "Current_Ratio4";
            ExcelSheet.Cells[2, 13] = "RES_IA";
            ExcelSheet.Cells[2, 14] = "RES_IB";
            ExcelSheet.Cells[2, 15] = "RES_IIA";
            ExcelSheet.Cells[2, 16] = "RES_IIB";
            ExcelSheet.Cells[2, 17] = "SNS_IA";
            ExcelSheet.Cells[2, 18] = "SNS_IB";
            ExcelSheet.Cells[2, 19] = "SNS_IIA";
            ExcelSheet.Cells[2, 20] = "SNS_IIB";
            ExcelSheet.Cells[2, 21] = "LED_F1";
            ExcelSheet.Cells[2, 22] = "T";
            ExcelSheet.Cells[2, 23] = "Second";

            //输出各个参数值
            for (int i = 0; i < RMS1Twelveinches.Count; i++)
            {
                ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                ExcelSheet.Cells[3 + i, 2] = RMS1Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 3] = RMS2Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 4] = Val2Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 5] = Val3Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 6] = RMSMID1Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 7] = RMSMID2Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 8] = RMSTwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 9] = CurrentRatio1Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 10] = CurrentRatio2Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 11] = CurrentRatio3Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 12] = CurrentRatio4Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 13] = RESIATwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 14] = RESIBTwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 15] = RESIIATwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 16] = RESIIBTwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 17] = SNSIATwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 18] = SNSIBTwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 19] = SNSIIATwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 20] = SNSIIBTwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 21] = LEDF1Twelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 22] = TTwelveinches[i].ToString();
                ExcelSheet.Cells[3 + i, 23] = ((int)SecondTwelveinches[i] / 3600).ToString() + ":" + (((int)SecondTwelveinches[i] % 3600) / 60).ToString() + ":" + (((int)SecondTwelveinches[i] % 3600) % 60).ToString();
            }

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序    

            ClearTwelveInchesLampsParameters();

            MessageBox.Show("数据已导出! 保存至D盘Excel文档", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
       }
        #endregion

        #region 清空12寸灯具参数
        private void ClearTwelveInchesLampsParameters()
        {
            RMS1Twelveinches.Clear();
            RMS2Twelveinches.Clear();
            Val2Twelveinches.Clear();
            Val3Twelveinches.Clear();
            RMSMID1Twelveinches.Clear();
            RMSMID2Twelveinches.Clear();
            RMSTwelveinches.Clear();
            CurrentRatio1Twelveinches.Clear();
            CurrentRatio2Twelveinches.Clear();
            CurrentRatio3Twelveinches.Clear();
            CurrentRatio4Twelveinches.Clear();
            RESIATwelveinches.Clear();
            RESIBTwelveinches.Clear();
            RESIIATwelveinches.Clear();
            RESIIBTwelveinches.Clear();
            SNSIATwelveinches.Clear();
            SNSIBTwelveinches.Clear();
            SNSIIATwelveinches.Clear();
            SNSIIBTwelveinches.Clear();
            LEDF1Twelveinches.Clear();
            TTwelveinches.Clear();
            SecondTwelveinches.Clear();

            ReceivedStatusFeedbackCommand.Clear();
        }

        #region 清空8寸灯具参数
        private void ClearEightInchesLampsParameter()
        {
            RMS1Eightinches.Clear();
            Val2Eightinches.Clear();
            Val3Eightinches.Clear();
            RMSEightinches.Clear();
            CurrentRatio1Eightinches.Clear();
            CurrentRatio3Eightinches.Clear();
            RESIAEightinches.Clear();
            RESIIAEightinches.Clear();
            SNSIAEightinches.Clear();
            SNSIIAEightinches.Clear();
            LEDF1Eightinches.Clear();
            TEightinches.Clear();
            SecondEightinches.Clear();

            ReceivedStatusFeedbackCommand.Clear();
        }
        #endregion

        #endregion

        #endregion

        #region 工厂模式参数设置
        private void SetLightParametersInFactoryMode_Click(object sender, RoutedEventArgs e)
        {
            
            if(ConfirmSettingLampParameter.Text!=""&&ConfirmSettingSpecialLampParameter.Text!=""&&ConfirmSettingOpenCircuitParameter.Text!="")
            {
                ConfigureSettingParametersCommand();               

                if (MessageBox.Show("是否将指令写入灯具？", "问询", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    if (lampsPort.IsOpen)
                    {
                        isQueryStatus = false;
                        lampsPort.Write(settingParameterCommand, 0, 28);
                    }
                    else
                    {
                        MessageBox.Show("串口未打开！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("参数选择不能为空！请完成配置", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        #region 灯具类型选择
        private void SelectApproachChenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectApproachChenterlineLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureApproachChenterlineLightParameters();
        }

        private void SelectApproachCrossbarLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectApproachCrossbarLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureApproachCrossbarLightParameters();
        }

        private void SelectApproachSideRowLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectApproachSideRowLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureApproachSideRowLightParameters();
        }

        private void SelectRWYThresholdWingBarLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRWYThresholdWingBarLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRWYThresholdWingBarLightParameters();
        }

        private void SelectRWYThresholdLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRWYThresholdLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRWYThresholdLightParameters();
        }

        private void SelectRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRWYEdgeLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRWYEdgeLightParameters();
        }

        private void Select12inchesRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = Select12inchesRWYEndLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            Configure12inchesRWYEndLightParameters();
        }

        private void SelectRWYThresholdEndLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRWYThresholdEndLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRWYThresholdEndLightParameters();
        }

        private void SelectRWYCenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRWYCenterlineLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRWYCenterlineLightParameters();
        }

        private void SelectRWYTouchdownZoneLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRWYTouchdownZoneLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRWYTouchdownZoneLightParameters();
        }

        private void Select8inchesRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = Select8inchesRWYEndLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            Configure8inchesRWYEndLightParameters();
        }

        private void SelectRapidExitTWYIndicatorLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Collapsed;
            NoSpecialLight.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectRapidExitTWYIndicatorLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = NoSpecialLight.Content.ToString();
            }));

            ConfigureRapidExitTWYIndicatorLightParameters();
        }

        private void SelectCombinedRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupSpecialLights.Visibility = Visibility.Visible;
            NoSpecialLight.Visibility = Visibility.Collapsed;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectCombinedRWYEdgeLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = "";
            }));
        }
    #endregion

        #region 特殊灯具选择
        private void SelectSpecialWhiteYellowAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {               
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteYellowAllAround.Content.ToString();
            }));

            ConfigureSpecialWhiteYellowAllAroundParameters();
        }

        private void SelectSpecialWhiteWhiteAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteWhiteAllAround.Content.ToString();
            }));

            ConfigureSpecialWhiteWhiteAllAroundParameters();
        }

        private void SelectSpecialWhiteRedAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteRedAllAround.Content.ToString();
            }));

            ConfigureSpecialWhiteRedAllAroundParameters();
        }

        private void SelectSpecialRedYellowAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialRedYellowAllAround.Content.ToString();
            }));

            ConfigureSpecialRedYellowAllAroundParameters();
        }

        private void SelectSpecialWhiteYellow_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteYellow.Content.ToString();
            }));

            ConfigureSpecialWhiteYellowParameters();
        }

        private void SelectSpecialWhiteWhite_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteWhite.Content.ToString();
            }));

            ConfigureSpecialWhiteWhiteParameters();
        }

        private void SelectSpecialWhiteRed_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteRed.Content.ToString();
            }));

            ConfigureSpecialWhiteRedParameters();
        }

        private void SelectSpecialRedYellow_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialRedYellow.Content.ToString();
            }));

            ConfigureSpecialRedYellowParameters();
        }
    #endregion

        #region 开路选择
        private void SelectOpenCircuitTrue_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingOpenCircuitParameter.Text = SelectOpenCircuitTrue.Content.ToString();
            }));

            ConfigureOpenCircuitTrue();
        }

        private void SelectOpenCircuitFalse_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmSettingOpenCircuitParameter.Text = SelectOpenCircuitFalse.Content.ToString();
            }));

            ConfigureOpenCircuitFalse();
        }
    #endregion

        #region 工厂模式配置一般灯具参数
        private void ConfigureApproachChenterlineLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x01;               
        }

        private void ConfigureApproachCrossbarLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x02;
        }

        private void ConfigureApproachSideRowLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x03;
        }

        private void ConfigureRWYThresholdWingBarLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x04;
        }

        private void ConfigureRWYThresholdLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x05;
        }

        private void ConfigureRWYEdgeLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x06;
        }

        private void Configure12inchesRWYEndLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x07;
        }

        private void ConfigureRWYThresholdEndLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x08;
        }

        private void ConfigureRWYCenterlineLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x09;
        }

        private void ConfigureRWYTouchdownZoneLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x0A;
        }

        private void Configure8inchesRWYEndLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x0B;
        }

        private void ConfigureRapidExitTWYIndicatorLightParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x0C;
        }

        #endregion

        #region 工厂模式配置特殊灯具参数
        private void ConfigureSpecialWhiteYellowAllAroundParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x07;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x04;
            settingIIA[2] = 0x05;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x0D;
        }

        private void ConfigureSpecialWhiteWhiteAllAroundParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x09;
            settingIB[2] = 0x03;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x04;
            settingIIA[2] = 0x05;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x0E;
        }

        private void ConfigureSpecialWhiteRedAllAroundParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x05;
            settingIB[2] = 0x05;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x04;
            settingIIA[2] = 0x05;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x0F;
        }

        private void ConfigureSpecialRedYellowAllAroundParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x05;
            settingIA[2] = 0x05;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x07;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x04;
            settingIIA[2] = 0x05;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x10;
        }

        private void ConfigureSpecialWhiteYellowParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x07;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x11;
        }

        private void ConfigureSpecialWhiteWhiteParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x09;
            settingIB[2] = 0x03;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x12;
        }

        private void ConfigureSpecialWhiteRedParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x05;
            settingIB[2] = 0x05;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x13;
        }

        private void ConfigureSpecialRedYellowParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x05;
            settingIA[2] = 0x05;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x07;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x14;
        }
        #endregion

        #region 工厂模式配置灯具参数是否开路
        private void ConfigureOpenCircuitTrue()
        {
            settingBreakFlag = 0x01;
        }

        private void ConfigureOpenCircuitFalse()
        {
            settingBreakFlag = 0x00;
        }

        #endregion      

        #endregion

        #region 开发者模式设置参数
        private void SetLightParametersInDeveloperMode_Click(object sender, RoutedEventArgs e)
        {            
            CalculateSetParameterCommand();
        }

        #region 参数设置
        private void SetParameterIA_TextChanged(object sender, TextChangedEventArgs e)
        {
            settingIA = CalculateCurrentBuffer(SetParameterIA.Text.ToString(), 1);
        }

        private void SetParameterIB_TextChanged(object sender, TextChangedEventArgs e)
        {
            settingIB = CalculateCurrentBuffer(SetParameterIB.Text.ToString(), 2);
        }

        private void SetParameterIIA_TextChanged(object sender, TextChangedEventArgs e)
        {
            settingIIA = CalculateCurrentBuffer(SetParameterIIA.Text.ToString(), 3);
        }

        private void SetParameterIIB_TextChanged(object sender, TextChangedEventArgs e)
        {
            settingIIB = CalculateCurrentBuffer(SetParameterIIB.Text.ToString(), 4);
        }

        private void SetReadResistanceFalse_Checked(object sender, RoutedEventArgs e)
        {
            settingReadRFlag = 0x00;
        }

        private void SetReadResistanceTrue_Checked(object sender, RoutedEventArgs e)
        {
            settingReadRFlag = 0x01;
        }

        private void SetMosfetTrue_Checked(object sender, RoutedEventArgs e)
        {
            settingMosFlag = 0x01;
        }

        private void SetMosfetFalse_Checked(object sender, RoutedEventArgs e)
        {
            settingMosFlag = 0x00;
        }

        private void SetBreakTrue_Checked(object sender, RoutedEventArgs e)
        {
            settingBreakFlag = 0x01;
        }

        private void SetBreakFalse_Checked(object sender, RoutedEventArgs e)
        {
            settingBreakFlag = 0x00;
        }

        private void SetLightNumber_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            settingLampsNumber = Convert.ToByte(SetLightNumber.SelectedIndex+1);
        }
        #endregion

        #region 生成电流数组
        private byte[] CalculateCurrentBuffer(string stringCurrentValue, int textboxNumber)
        {
            stringCurrentValue = stringCurrentValue.Trim();
            char[] charCurrentValue = stringCurrentValue.ToCharArray();
            byte[] result = new byte[4] { 0x00, 0x00, 0x00, 0x00 };

            for (int i = 0; i < charCurrentValue.Length; i++)
            {
                if ((charCurrentValue[i] >= '0' && charCurrentValue[i] <= '9') || charCurrentValue[i] == '.')
                {

                }
                else
                {
                    MessageBox.Show("非法输入！请输入数字 0~9、字符 '.'", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    PrugeTextBoxContent(textboxNumber);
                    return result;
                }
            }
            if (charCurrentValue.Length >= 1)
            {
                if (charCurrentValue[0] == '0' || charCurrentValue[0] == '1')
                {
                    if (charCurrentValue.Length > 1 && charCurrentValue.Length < 6)
                    {
                        if (charCurrentValue[1] == '.')
                        {
                            if (charCurrentValue.Length == 2)
                            {
                                result[0] = Convert.ToByte(charCurrentValue[0] - '0');
                            }
                            else if (charCurrentValue.Length == 3)
                            {
                                if (charCurrentValue[2] == '.')
                                {
                                    MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                    PrugeTextBoxContent(textboxNumber);
                                    return result;
                                }
                                else
                                {
                                    if (charCurrentValue[0] == '1' && charCurrentValue[2] != '0')
                                    {
                                        MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                        PrugeTextBoxContent(textboxNumber);
                                        return result;
                                    }
                                    else
                                    {
                                        result[0] = Convert.ToByte(charCurrentValue[0] - '0');
                                        result[1] = Convert.ToByte(charCurrentValue[2] - '0');
                                    }
                                }

                            }
                            else if (charCurrentValue.Length == 4)
                            {
                                if (charCurrentValue[2] == '.' || charCurrentValue[3] == '.')
                                {
                                    MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                    PrugeTextBoxContent(textboxNumber);
                                    return result;
                                }
                                else
                                {
                                    if (charCurrentValue[0] == '1' && (charCurrentValue[2] != '0' || charCurrentValue[3] != '0'))
                                    {
                                        MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                        PrugeTextBoxContent(textboxNumber);
                                        return result;
                                    }
                                    else
                                    {
                                        result[0] = Convert.ToByte(charCurrentValue[0] - '0');
                                        result[1] = Convert.ToByte(charCurrentValue[2] - '0');
                                        result[2] = Convert.ToByte(charCurrentValue[3] - '0');
                                    }
                                }
                            }
                            else
                            {
                                if (charCurrentValue[2] == '.' || charCurrentValue[3] == '.' || charCurrentValue[4] == '.')
                                {
                                    MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                    PrugeTextBoxContent(textboxNumber);
                                    return result;
                                }
                                else
                                {
                                    if (charCurrentValue[0] == '1' && (charCurrentValue[2] != '0' || charCurrentValue[3] != '0' || charCurrentValue[4] != '0'))
                                    {
                                        MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                        PrugeTextBoxContent(textboxNumber);
                                        return result;
                                    }
                                    else
                                    {
                                        result[0] = Convert.ToByte(charCurrentValue[0] - '0');
                                        result[1] = Convert.ToByte(charCurrentValue[2] - '0');
                                        result[2] = Convert.ToByte(charCurrentValue[3] - '0');
                                        result[3] = Convert.ToByte(charCurrentValue[4] - '0');
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                            PrugeTextBoxContent(textboxNumber);
                            return result;
                        }
                    }
                    else if (charCurrentValue.Length == 1)
                    {
                        result[0] = Convert.ToByte(charCurrentValue[0] - '0');
                    }
                    else
                    {
                        MessageBox.Show("位数错误！小数点后最多保留三位，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                        PrugeTextBoxContent(textboxNumber);
                        return result;
                    }
                }
                else if (charCurrentValue[0] == '.')
                {
                    MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    PrugeTextBoxContent(textboxNumber);
                    return result;
                }
                else
                {
                    MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    PrugeTextBoxContent(textboxNumber);
                    return result;
                }

            }

            return result;
        }

        #endregion

        #region 在生产电流数组函数中清空指定文本框
        private void PrugeTextBoxContent(int Number)
        {
            switch (Number)
            {
                case 1: SetParameterIA.Text = ""; break;
                case 2: SetParameterIB.Text = ""; break;
                case 3: SetParameterIIA.Text = ""; break;
                case 4: SetParameterIIB.Text = ""; break;
            }
        }


        #endregion

        #region 开发者模式 生成并发送设置参数指令
        private void CalculateSetParameterCommand()
        {
            ShowSetParameterCommand.Text = "";
            if (SetParameterIA.Text != "" && SetParameterIB.Text != "" && SetParameterIIA.Text != "" && SetParameterIIB.Text != "")
            {
                ConfigureSettingParametersCommand();

                for (int i=0;i<settingParameterCommand.Length;i++)
                {                              
                        ShowSetParameterCommand.Text += Convert.ToString(settingParameterCommand[i], 16).PadLeft(2, '0').ToUpper() + " ";                                
                }

                if(MessageBox.Show("是否将指令写入灯具？","问询",MessageBoxButton.YesNo,MessageBoxImage.Question)==MessageBoxResult.Yes)
                {
                    if(lampsPort.IsOpen)
                    {
                        isQueryStatus = false;
                        lampsPort.Write(settingParameterCommand, 0, 28);
                    }
                    else
                    {
                        MessageBox.Show("串口未打开！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }                                
            else
            {
                MessageBox.Show("文本框不能为空！请输入电流值", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }
        #endregion
        #endregion

        #region 串口数据接收函数
        private void lampsPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (lampsPort.IsOpen)
            {
                byte[] dataReceived = new byte[lampsPort.BytesToRead];
                lampsPort.Read(dataReceived, 0, dataReceived.Length);

                if (isQueryStatus == true)
                {
                    ReceivedStatusFeedbackCommand.AddRange(dataReceived);

                    QueryStatusTimeSpan = DateTime.Now - StartQueryStatus;
                    this.Dispatcher.Invoke(new System.Action(() =>
                    {
                        AnswerStatus.Text = "灯具状态查询已耗时: " + QueryStatusTimeSpan.Milliseconds.ToString()+" ms";
                    }));
                }
                else
                {
                    if (dataReceived.Length != 0)
                    {
                        if (dataReceived[0] == 0x02)
                        {
                            //byte checkOutValue = CalculateCheckOutValue(dataReceived);
                            //if(checkOutValue==dataReceived[dataReceived.Length-1])
                            //{
                            if (dataReceived[1] == 0x89 && dataReceived[2] == 0x22 && dataReceived[3] == 0x85)  //版本反馈指令
                            {
                                year = dataReceived[4];
                                month = dataReceived[5];
                                date = dataReceived[6];
                                hardwareVersion = dataReceived[7];
                                versionBigNumber = dataReceived[8];
                                versionSmallNumber = dataReceived[9];

                                this.Dispatcher.Invoke(new System.Action(() =>
                                {
                                    AnswerVersion.Text = "20" + year.ToString() + "年" + month.ToString() + "月" + date.ToString() + "日 " + hardwareVersion.ToString() + "寸 " + versionBigNumber.ToString() + "." + versionSmallNumber.ToString() + "版";

                                }));
                            }
                            else if (dataReceived[1] == 0x55 && dataReceived[2] == 0x11)  //设置反馈指令
                            {
                                MessageBox.Show("参数设置成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                                return;
                            }
                            else
                            {
                                MessageBox.Show("解析错误", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                            //}
                            //else
                            //{
                            //    MessageBox.Show("校验错误", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                            //}
                        }
                        else
                        {
                            MessageBox.Show("帧头错误", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("未接收到串口指令！", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("串口未打开！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }



        #endregion

        #region 生成设置参数指令
        public void ConfigureSettingParametersCommand()
        {
            for (int i = 0; i < settingIA.Length; i++)
            {
                settingParameterCommand[5 + i] = settingIA[i];
            }
            for (int i = 0; i < settingIB.Length; i++)
            {
                settingParameterCommand[9 + i] = settingIB[i];
            }
            for (int i = 0; i < settingIIA.Length; i++)
            {
                settingParameterCommand[13 + i] = settingIIA[i];
            }
            for (int i = 0; i < settingIIB.Length; i++)
            {
                settingParameterCommand[17 + i] = settingIIB[i];
            }
            settingParameterCommand[21] = settingReadRFlag;
            settingParameterCommand[22] = settingMosFlag;
            settingParameterCommand[23] = settingBreakFlag;
            settingParameterCommand[24] = settingLampsNumber;
            settingParameterCommand[27] = CalculateCheckOutValue(settingParameterCommand);
        }



        #endregion

        #region 用户登录
        private void LogIn_Click(object sender, RoutedEventArgs e)
        {
            userName = UserName.Text.ToString().Trim();
            password = Password.Password;
            
            if (userName =="Airsafe")
            {
                if(password == "Airsafe")
                {
                    DeveloperMode.Visibility = Visibility.Visible;                   
                }
                else
                {
                    MessageBox.Show("密码错误！请重新输入密码", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                    UserName.Text = "";
                    Password.Password = "";
                }                
            }
            else
            {
                MessageBox.Show("账号错误！请重新输入账号", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                UserName.Text = "";
                Password.Password = "";
            }
        }       

        private void LogOut_Click(object sender, RoutedEventArgs e)
        {
            if(DeveloperMode.Visibility==Visibility.Visible)
            {
                DeveloperMode.Visibility = Visibility.Collapsed;                
            }
            UserName.Text = "";
            Password.Password = "";
        }

        #endregion
    }
}
