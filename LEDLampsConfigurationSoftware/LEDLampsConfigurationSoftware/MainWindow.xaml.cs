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
using System.Threading;
using System.IO;

namespace LEDLampsConfigurationSoftware
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    
    public partial class MainWindow : System.Windows.Window
    {
        #region 设置全局变量
        SerialPort lampsPort = new SerialPort();  //定义串口       
        int judgeFeedbackCommand = 0;  //设置参数反馈指令为1，版本查询反馈指令为2，状态查询反馈指令为3，无反馈指令为0，打开串口时发送版本查询指令为4
                
        #endregion

        #region 版本反馈指令参数
        int year = 0;
        int month = 0;
        int date = 0;
        int hardwareVersion = 0;
        int versionBigNumber = 0;
        int versionSmallNumber = 0;
        double currentRatio1 = 0;
        double currentRatio2 = 0;
        double currentRatio3 = 0;
        double currentRatio4 = 0;
        int breakFlag = 0;
        byte lampsNumber = 0;
        #endregion

        #region 工厂模式下，设置参数指令参数
        byte[] settingIA = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] settingIB = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] settingIIA = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] settingIIB = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte settingReadRFlag = 0x00;  //读取电阻
        byte settingMosFlag = 0x00;  //不开MOSFET
        byte settingBreakFlag = 0x00;  //不带开路
        byte settingLampsNumber = 0x00;  //灯具编号
        #endregion

        #region 开发者模式下，设置参数指令参数
        byte[] InDeveloperModeSettingIA = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] InDeveloperModeSettingIB = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] InDeveloperModeSettingIIA = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte[] InDeveloperModeSettingIIB = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
        byte InDeveloperModeSettingReadRFlag = 0x00;  //读取电阻
        byte InDeveloperModeSettingMosFlag = 0x00;  //不开MOSFET
        byte InDeveloperModeSettingBreakFlag = 0x00;  //不带开路
        byte InDeveloperModeSettingLampsNumber = 0x00;  //灯具编号
        #endregion


        #region 发送指令集（尚未计算校验值）
        Byte[] queryStatusCommand = new Byte[28] { 0x02,0x89,0x11,0x58,0x12,0x00,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00 };
        Byte[] queryVersionCommand = new Byte[28] { 0x02, 0x89, 0x22, 0x85, 0x12, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00 };
        Byte[] settingParameterCommand = new Byte[28] { 0x02, 0x55, 0x11, 0x58, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        Byte[] InDeveloperModeSettingParameterCommand = new Byte[28] { 0x02, 0x55, 0x11, 0x58, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

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
        ArrayList ErrorCodeEightinches = new ArrayList();
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
        ArrayList ErrorCodeTwelveinches = new ArrayList();
        #endregion

        public MainWindow()
        {                      
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {           
            lampsPort.BaudRate = 921600;
            lampsPort.Parity = Parity.None;
            lampsPort.DataBits = 8;
            lampsPort.StopBits = StopBits.One;
            lampsPort.DataReceived += new SerialDataReceivedEventHandler(lampsPortDataReceived);
            
            string[] portNames = SerialPort.GetPortNames();      //获取当前电脑所有串口号
            PortNameSelect.ItemsSource = portNames;    //将串口号显示在ComboBox
            PortNameSelect.SelectedIndex = portNames.Length - 1;

            queryVersionCommand[27] = CalculateCheckOutValue(queryVersionCommand);  //计算版本查询指令的校验值
            queryStatusCommand[27] = CalculateCheckOutValue(queryStatusCommand);  //计算状态查询指令的校验值    

            SettingSerialPort.IsSelected = true;                             
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("是否关闭软件", "问询", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result==MessageBoxResult.Yes)
            {
                ConfigurationWindow.IsEnabled = true;

                if (lampsPort.IsOpen==true)
                {
                    if(MessageBox.Show("串口未关闭！请先关闭串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.None)
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                    else
                    {
                        e.Cancel = true;
                        ConfigurationWindow.IsEnabled = true;
                    }                    
                }                
            }
            else if(result == MessageBoxResult.No)
            {
                e.Cancel = true;
                ConfigurationWindow.IsEnabled = true;
            }
            else if(result == MessageBoxResult.None)
            {
                ConfigurationWindow.IsEnabled = false;
            }                       
        }

        #region 串口操作
        //刷新串口号        
        private void PortNameSelect_DropDownOpened(object sender, EventArgs e)
        {
            string[] portNames = SerialPort.GetPortNames();      //获取当前电脑所有串口号
            PortNameSelect.ItemsSource = portNames;    //将串口号显示在ComboBox
            PortNameSelect.SelectedIndex = portNames.Length - 1;
        }

        //打开串口
        private void OpenSerialPort_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lampsPort.IsOpen == false)
                {
                    lampsPort.PortName = PortNameSelect.SelectedItem.ToString();
                    lampsPort.Open();                   
                }
                if(lampsPort.IsOpen==true)
                {
                    PurgingDeveloperMode();
                    PurgingFactoryMode();
                    judgeFeedbackCommand = 4;
                    LampInchesLabel.Content = "";
                    lampsPort.Write(queryVersionCommand, 0, 28);
                    Thread.Sleep(50);
                    if (judgeFeedbackCommand == 4)
                    {
                        if( MessageBox.Show("查询灯具尺寸失败!请重新打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }

                        SelectApproachChenterlineLight.IsEnabled = false;
                        SelectApproachCrossbarLight.IsEnabled = false;
                        SelectApproachSideRowLight.IsEnabled = false;
                        SelectRWYThresholdWingBarLight.IsEnabled = false;
                        SelectRWYThresholdLight.IsEnabled = false;
                        SelectRWYEdgeLight.IsEnabled = false;
                        Select12inchesRWYEndLight.IsEnabled = false;
                        SelectRWYThresholdEndLight.IsEnabled = false;
                        SelectRWYCenterlineLight.IsEnabled = false;
                        SelectRWYTouchdownZoneLight.IsEnabled = false;
                        Select8inchesRWYEndLight.IsEnabled = false;
                        SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                        SelectCombinedRWYEdgeLight.IsEnabled = false;

                        LampInchesLabel.Content = "";
                        LampInchesLabel.Content = "查询当前灯具驱动板尺寸失败";
                    }
                    else
                    {
                        if (MessageBox.Show("串口已打开！", "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                            FactoryMode.IsSelected = true;                            
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                    }                                                      
                }
            }
           catch
            {
                if( MessageBox.Show("串口未打开！请选择正确的串口号", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
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
                    LampInchesLabel.Content = "";
                    PurgingDeveloperMode();
                    PurgingFactoryMode();
                }
                if(lampsPort.IsOpen==false)
                {
                    if( MessageBox.Show("串口已关闭！", "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }
            }
            catch
            {
                if( MessageBox.Show("关闭串口失败！请重启软件", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
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
                judgeFeedbackCommand = 2;                
                lampsPort.Write(queryVersionCommand, 0, 28);
                Thread.Sleep(50);
                if(judgeFeedbackCommand==2)
                {
                    if( MessageBox.Show("未接收到反馈指令！请重新查询版本", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }
            }
            else
            {
                if( MessageBox.Show("未打开串口！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
            }                       
        }
        #endregion

        #region 状态查询       
        DateTime StartQueryStatus;
        TimeSpan QueryStatusTimeSpan;
        private void QueryStatus_Click(object sender, RoutedEventArgs e)
        {
            if (lampsPort.IsOpen)
            {
                if(ShowEXCELHandleProcess.Visibility==Visibility.Hidden&&ShowTXTHandleProcess.Visibility==Visibility.Hidden)
                {
                    ReceivedStatusFeedbackCommand.Clear();
                    judgeFeedbackCommand = 3;
                    AnswerStatus.Text = "";
                    lampsPort.Write(queryStatusCommand, 0, 28);
                    StartQueryStatus = DateTime.Now;
                }
                else
                {
                    if( MessageBox.Show("正在导出数据！请稍候进行状态查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;                        
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }                
            }
            else
            {
                if( MessageBox.Show("未打开串口！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
            }
        }
        #endregion

        #region 串口数据接收函数
        byte[] dataReceived;
        private void lampsPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (lampsPort.IsOpen)
            {
                dataReceived = new byte[lampsPort.BytesToRead];
                lampsPort.Read(dataReceived, 0, dataReceived.Length);

                if (dataReceived.Length != 0)
                {
                    switch (judgeFeedbackCommand)
                    {
                        case 0: NoFeedbackCommand(); break;
                        case 1: SetParameterFeedbackCommand(); break;
                        case 2: QueryVersionFeedbackCommand(); break;
                        case 3: QueryStatusFeedbackCommand(); break;
                        case 4: ConfirmLampInches(); break;
                    }
                }
                
            }
            else
            {                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("串口未打开！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
                return;
            }
        }

        private void NoFeedbackCommand()
        {
            if(dataReceived.Length==1&&dataReceived[0]==0x00)
            {
                return;
            }
            else
            {
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("指令错误！", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
            }
        }

        private void SetParameterFeedbackCommand()
        {
            judgeFeedbackCommand = 0;            
            
            if (dataReceived.Length == 4)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x55 && dataReceived[2] == 0x11)
                    {
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if ( MessageBox.Show("参数设置成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                            {                           
                                    ConfigurationWindow.IsEnabled = true;                                                     
                            }
                            else
                            {                           
                                    ConfigurationWindow.IsEnabled = false;                            
                            }
                        }));
                        return;
                    }
                    else
                    {                       
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show("帧头错误!请重新设置参数", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }
                        }));
                        return;
                    }
                }
                else
                {                   
                    this.Dispatcher.Invoke(new System.Action(() =>
                    {
                        if (MessageBox.Show("校验错误!请重新设置参数", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                    }));
                    return;
                }
            }
            else
            {               
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("指令长度错误!请重新设置参数", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
                return;
            }
        }

        private void QueryVersionFeedbackCommand()
        {
            judgeFeedbackCommand = 0;
            if (dataReceived.Length == 20)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x22 && dataReceived[3] == 0x85)
                    {
                        year = dataReceived[4];
                        month = dataReceived[5];
                        date = dataReceived[6];
                        hardwareVersion = dataReceived[7];
                        versionBigNumber = dataReceived[8];
                        versionSmallNumber = dataReceived[9];
                        breakFlag = dataReceived[14];
                        lampsNumber = dataReceived[15];                       

                        currentRatio1 = CalculateRealCurrentValue(dataReceived[10]);
                        currentRatio2 = CalculateRealCurrentValue(dataReceived[11]);
                        currentRatio3 = CalculateRealCurrentValue(dataReceived[12]);
                        currentRatio4 = CalculateRealCurrentValue(dataReceived[13]);
                        
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            PurgeAnswerVersionTextblock();

                            AnswerVersion.Text = hardwareVersion.ToString() + "寸 " + versionBigNumber.ToString() + "." + versionSmallNumber.ToString() + "版 " + " 20" + year.ToString() + "年" + month.ToString() + "月" + date.ToString() + "日";
                            AnswerLampModel.Text= LampsContentShow(lampsNumber);
                            AnswerIA.Text = currentRatio1.ToString();
                            AnswerIB.Text = currentRatio2.ToString();
                            AnswerIIA.Text = currentRatio3.ToString();
                            AnswerIIB.Text = currentRatio4.ToString();                           
                            if (breakFlag == 0)
                            {
                                AnswerOpenCircuit.Text = "不带开路";
                            }
                            else
                            {
                                AnswerOpenCircuit.Text = "带开路";
                            }

                            
                        }));
                        
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show("版本查询成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }
                        }));
                    }
                    else
                    {                        
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show("帧头错误!请重新查询版本", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }
                        }));
                        return;
                    }
                }
                else
                {                    
                    this.Dispatcher.Invoke(new System.Action(() =>
                    {
                        if (MessageBox.Show("校验错误!请重新查询版本", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                    }));
                    return;
                }
            }
            else
            {               
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("指令长度错误!请重新查询版本", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
                return;
            }
        }

        private void ConfirmLampInches()
        {
            judgeFeedbackCommand = 0;
            if (dataReceived.Length == 20)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x22 && dataReceived[3] == 0x85)
                    {                        
                        hardwareVersion = dataReceived[7];                        

                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            LampInchesLabel.Content = "";
                            LampInchesLabel.Content = "当前灯具驱动板尺寸为" + hardwareVersion.ToString() + "寸";
                                              
                            if(hardwareVersion==8)
                            {
                                SelectApproachChenterlineLight.IsEnabled = false;
                                SelectApproachCrossbarLight.IsEnabled = false;
                                SelectApproachSideRowLight.IsEnabled = false;
                                SelectRWYThresholdWingBarLight.IsEnabled = false;
                                SelectRWYThresholdLight.IsEnabled = false;
                                SelectRWYEdgeLight.IsEnabled = false;
                                Select12inchesRWYEndLight.IsEnabled = false;
                                SelectRWYThresholdEndLight.IsEnabled = false;
                                SelectRWYCenterlineLight.IsEnabled = true;
                                SelectRWYTouchdownZoneLight.IsEnabled = true;
                                Select8inchesRWYEndLight.IsEnabled = true;
                                SelectRapidExitTWYIndicatorLight.IsEnabled = true;
                                SelectCombinedRWYEdgeLight.IsEnabled = false;
                            }
                            else if (hardwareVersion == 12)
                            {
                                SelectApproachChenterlineLight.IsEnabled = true;
                                SelectApproachCrossbarLight.IsEnabled = true;
                                SelectApproachSideRowLight.IsEnabled = true;
                                SelectRWYThresholdWingBarLight.IsEnabled = true;
                                SelectRWYThresholdLight.IsEnabled = true;
                                SelectRWYEdgeLight.IsEnabled = true;
                                Select12inchesRWYEndLight.IsEnabled = true;
                                SelectRWYThresholdEndLight.IsEnabled = true;
                                SelectRWYCenterlineLight.IsEnabled = false;
                                SelectRWYTouchdownZoneLight.IsEnabled = false;
                                Select8inchesRWYEndLight.IsEnabled = false;
                                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                                SelectCombinedRWYEdgeLight.IsEnabled = true;
                            }
                            else
                            {
                                SelectApproachChenterlineLight.IsEnabled = false;
                                SelectApproachCrossbarLight.IsEnabled = false;
                                SelectApproachSideRowLight.IsEnabled = false;
                                SelectRWYThresholdWingBarLight.IsEnabled = false;
                                SelectRWYThresholdLight.IsEnabled = false;
                                SelectRWYEdgeLight.IsEnabled = false;
                                Select12inchesRWYEndLight.IsEnabled = false;
                                SelectRWYThresholdEndLight.IsEnabled = false;
                                SelectRWYCenterlineLight.IsEnabled = false;
                                SelectRWYTouchdownZoneLight.IsEnabled = false;
                                Select8inchesRWYEndLight.IsEnabled = false;
                                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                                SelectCombinedRWYEdgeLight.IsEnabled = false;
                            }
                        }));
                    }
                    else
                    {                        
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show("查询灯具尺寸失败!请重新打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }
                        }));
                        return;
                    }
                }
                else
                {                    
                    this.Dispatcher.Invoke(new System.Action(() =>
                    {
                        if (MessageBox.Show("查询灯具尺寸失败!请重新打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                    }));
                    return;
                }
            }
            else
            {                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("查询灯具尺寸失败!请重新打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
                return;
            }
        }

        ArrayList ReceivedStatusFeedbackCommand = new ArrayList();  //定义接收到的状态反馈指令        
        private void QueryStatusFeedbackCommand()
        {
            ReceivedStatusFeedbackCommand.AddRange(dataReceived);

            QueryStatusTimeSpan = DateTime.Now - StartQueryStatus;
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                AnswerStatus.Text = "灯具状态查询已耗时: " + QueryStatusTimeSpan.Hours.ToString().PadLeft(2, '0') + ":" + QueryStatusTimeSpan.Minutes.ToString().PadLeft(2, '0') + ":" + QueryStatusTimeSpan.Seconds.ToString().PadLeft(2, '0') + ":" + QueryStatusTimeSpan.Milliseconds.ToString().PadLeft(3, '0');
            }));
        }

        public string LampsContentShow(byte lampNumber)
        {
            string result = "";
            switch (lampNumber)
            {
                case 0: result = "无灯具型号"; break;
                case 1: result = "进近中线灯(APPS-12-S-LED-C)"; break;
                case 2: result = "进近横排灯(APPS-12-L-LED-C)"; break;
                case 3: result = "进近横排灯(APPS-12-R-LED-C)"; break;               
                case 4: result = "进近侧边灯(APSS-12-L-LED-R)"; break;
                case 5: result = "进近侧边灯(APSS-12-R-LED-R)"; break;
                case 6: result = "跑道入口翼排灯(THWS-12-L-LED-G)"; break;
                case 7: result = "跑道入口翼排灯(THWS-12-R-LED-G)"; break;
                case 8: result = "跑道入口灯(THRS-12-L-LED-G)"; break;
                case 9: result = "跑道入口灯(THRS-12-R-LED-G)"; break;
                case 10: result = "跑道入口灯(THRS-12-S-LED-G)"; break;
                case 11: result = "跑道边灯(RELS-12-L-LED-YC)"; break;
                case 12: result = "跑道边灯(RELS-12-R-LED-YC)"; break;
                case 13: result = "跑道边灯(RELS-12-L-LED-CY)"; break;
                case 14: result = "跑道边灯(RELS-12-R-LED-CY)"; break;
                case 15: result = "跑道边灯(RELS-12-L-LED-CC)"; break;
                case 16: result = "跑道边灯(RELS-12-R-LED-CC)"; break;
                case 17: result = "跑道边灯(RELS-12-L-LED-CR)"; break;
                case 18: result = "跑道边灯(RELS-12-R-LED-CR)"; break;
                case 19: result = "跑道边灯(RELS-12-L-LED-RC)"; break;
                case 20: result = "跑道边灯(RELS-12-R-LED-RC)"; break;
                case 21: result = "跑道末端灯（ENDS-12-LED-R）"; break;
                case 22: result = "跑道入口/末端灯(TAES-12-L-LED-GR-1P)"; break;
                case 23: result = "跑道入口/末端灯(TAES-12-R-LED-GR-1P)"; break;
                case 24: result = "跑道入口/末端灯(TAES-12-S-LED-GR-1P)"; break;
                case 25: result = "跑道中线灯(RCLS-08-LED-CB-1P)"; break;
                case 26: result = "跑道中线灯(RCLS-08-LED-RB-1P)"; break;
                case 27: result = "跑道中线灯(RCLS-08-LED-CC-1P)"; break;
                case 28: result = "跑道中线灯(RCLS-08-LED-RC-1P)"; break;
                case 29: result = "跑道接地带灯(TDZS-08-L-LED-C)"; break;
                case 30: result = "跑道接地带灯(TDZS-08-R-LED-C)"; break;
                case 31: result = "跑道末端灯（ENDS-08-LED-R）"; break;
                case 32: result = "快速出口滑行道指示灯(RAPS-08-LED-Y)"; break;
                case 33: result = "组合跑道边灯（RELC-12-LED-CY-C-1P）"; break;
                case 34: result = "组合跑道边灯（RELC-12-LED-CC-C-1P）"; break;
                case 35: result = "组合跑道边灯（RELC-12-LED-CR-C-1P）"; break;
                case 36: result = "组合跑道边灯（RELC-12-LED-RY-C-1P）"; break;
                case 37: result = "组合跑道边灯（RELC-12-LED-CY-B-1P）"; break;
                case 38: result = "组合跑道边灯（RELC-12-LED-CC-B-1P）"; break;
                case 39: result = "组合跑道边灯（RELC-12-LED-CR-B-1P）"; break;
                case 40: result = "组合跑道边灯（RELC-12-LED-RY-B-1P）"; break;
            }
            return result;
        }

        private double CalculateRealCurrentValue(byte originalData)
        {
            double original= originalData / 10.0;
            double result = 0.0;

            if(hardwareVersion == 12)
            {
                result = original;

                if (lampsNumber == 1 || lampsNumber == 2 ||lampsNumber==3)
                {
                    result = original * 1.3;
                }
                if(lampsNumber>=33&&lampsNumber<=40)
                {
                    if(original==0.9)
                    {
                        result = 0.93 * 1.3;
                    }
                    if(original==0.7)
                    {
                        result = 0.7 * 1;
                    }
                    if(original==0.6)
                    {
                        result = 0.55 * 1;
                    }
                    if(original==0.4)
                    {
                        result = 0.45 * 1;
                    }
                }
            }
            if(hardwareVersion == 8)
            {
                result = original * 0.66;
            }           

            result = Math.Round(result, 2,MidpointRounding.AwayFromZero);

            return result;

        }

        private void PurgeAnswerVersionTextblock()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                AnswerVersion.Text = "";
                AnswerLampModel.Text = "";
                AnswerIA.Text = "";
                AnswerIB.Text = "";
                AnswerIIA.Text = "";
                AnswerIIB.Text = "";
                AnswerOpenCircuit.Text = "";
            }));
        }
        #endregion

        #region 导出最终数据
        byte[] receivedStatusFeedbackCommand;
        Thread CreateExcelThread;
        private void CreatExcel_Click(object sender, RoutedEventArgs e)
        {
            ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
            {
                ShowEXCELHandleProcess.Visibility = Visibility.Visible;
            }));

            CreateExcelThread = new Thread(new ThreadStart(CreateExcelThreadProcess));
            CreateExcelThread.Start();            
        }

        private void CreateExcelThreadProcess()
        {
            if (judgeFeedbackCommand == 3)
            {
                if (ReceivedStatusFeedbackCommand.Count != 0)
                {
                    receivedStatusFeedbackCommand = new byte[ReceivedStatusFeedbackCommand.Count];

                    for (int i = 0; i < receivedStatusFeedbackCommand.Length; i++)
                    {
                        receivedStatusFeedbackCommand[i] = (byte)ReceivedStatusFeedbackCommand[i];
                    }
                    
                    if (receivedStatusFeedbackCommand[0] == 0x02 && receivedStatusFeedbackCommand[1] == 0xAA && receivedStatusFeedbackCommand[2] == 0x01 && receivedStatusFeedbackCommand[3] == 0x0C && receivedStatusFeedbackCommand[4] == 0x0C)
                    {
                        TwelveInchesLampDataAnalysis(receivedStatusFeedbackCommand);
                        TwelveInchesLampParametersCreatExcel();
                    }
                    else if (receivedStatusFeedbackCommand[0] == 0x02 && receivedStatusFeedbackCommand[1] == 0xAA && receivedStatusFeedbackCommand[2] == 0x01 && receivedStatusFeedbackCommand[3] == 0x08 && receivedStatusFeedbackCommand[4] == 0x08)
                    {
                        EightInchesLampDataAnalysis(receivedStatusFeedbackCommand);
                        EightInchesLampParametersCreatExcel();
                    }
                    else
                    {                       
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show("接收指令帧头错误！请重新查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }
                        }));

                        ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                        {
                            ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                        }));
                    }
                }
                else
                {                    
                    this.Dispatcher.Invoke(new System.Action(() =>
                    {
                        if (MessageBox.Show("接收指令不能为空！请重新查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                    }));

                    ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                    {
                        ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                    }));
                }
            }
            else
            {                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("未进行状态查询！请先进行状态查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));
            }
            CreateExcelThread.Abort();
        }

        #region 12寸灯具状态信息解析
        private void TwelveInchesLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] received12InchesStatusFeedbackCommandArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x0C && CompleteData[i + 4] == 0x0C)
                {
                    commandCount.Add(i);
                }
            }

            received12InchesStatusFeedbackCommandArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    received12InchesStatusFeedbackCommandArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    received12InchesStatusFeedbackCommandArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < received12InchesStatusFeedbackCommandArray[i].Length; j++)
                {
                    received12InchesStatusFeedbackCommandArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < received12InchesStatusFeedbackCommandArray.Length; i++)
            {
                if (received12InchesStatusFeedbackCommandArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(received12InchesStatusFeedbackCommandArray[i]);
                    if (checkOutValue == received12InchesStatusFeedbackCommandArray[i][received12InchesStatusFeedbackCommandArray[i].Length - 1])
                    {
                        RMS1Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][5] * 500);
                        RMS2Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][6] * 500);
                        Val2Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][7] * 20);
                        Val3Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][8]);
                        RMSMID1Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][9] * 16);
                        RMSMID2Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][10] * 16);
                        RMSTwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][11] * 4);
                        CurrentRatio1Twelveinches.Add((float)(received12InchesStatusFeedbackCommandArray[i][12] / 10.0));
                        CurrentRatio2Twelveinches.Add((float)(received12InchesStatusFeedbackCommandArray[i][13] / 10.0));
                        CurrentRatio3Twelveinches.Add((float)(received12InchesStatusFeedbackCommandArray[i][14] / 10.0));
                        CurrentRatio4Twelveinches.Add((float)(received12InchesStatusFeedbackCommandArray[i][15] / 10.0));
                        RESIATwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][16] * 124);
                        RESIBTwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][17] * 124);
                        RESIIATwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][18] * 124);
                        RESIIBTwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][19] * 124);
                        SNSIATwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][20] * 16);
                        SNSIBTwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][21] * 16);
                        SNSIIATwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][22] * 16);
                        SNSIIBTwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][23] * 16);
                        LEDF1Twelveinches.Add(received12InchesStatusFeedbackCommandArray[i][24]);
                        TTwelveinches.Add(received12InchesStatusFeedbackCommandArray[i][25]);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = received12InchesStatusFeedbackCommandArray[i][26 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondTwelveinches.Add(SecondResult);
                        ErrorCodeTwelveinches.Add("No Error");
                    }
                    else
                    {
                        TwelveInchesLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    TwelveInchesLampCommandLengthErrorHandle();
                }
            }
           
        }

        private void TwelveInchesLampCheckValueErrorHandle()
        {
            RMS1Twelveinches.Add("Null");
            RMS2Twelveinches.Add("Null");
            Val2Twelveinches.Add("Null");
            Val3Twelveinches.Add("Null");
            RMSMID1Twelveinches.Add("Null");
            RMSMID2Twelveinches.Add("Null");
            RMSTwelveinches.Add("Null");
            CurrentRatio1Twelveinches.Add("Null");
            CurrentRatio2Twelveinches.Add("Null");
            CurrentRatio3Twelveinches.Add("Null");
            CurrentRatio4Twelveinches.Add("Null");
            RESIATwelveinches.Add("Null");
            RESIBTwelveinches.Add("Null");
            RESIIATwelveinches.Add("Null");
            RESIIBTwelveinches.Add("Null");
            SNSIATwelveinches.Add("Null");
            SNSIBTwelveinches.Add("Null");
            SNSIIATwelveinches.Add("Null");
            SNSIIBTwelveinches.Add("Null");
            LEDF1Twelveinches.Add("Null");
            TTwelveinches.Add("Null");
            SecondTwelveinches.Add("Null");
            ErrorCodeTwelveinches.Add("Check Value Error");
        }

        private void TwelveInchesLampCommandLengthErrorHandle()
        {
            RMS1Twelveinches.Add("Null");
            RMS2Twelveinches.Add("Null");
            Val2Twelveinches.Add("Null");
            Val3Twelveinches.Add("Null");
            RMSMID1Twelveinches.Add("Null");
            RMSMID2Twelveinches.Add("Null");
            RMSTwelveinches.Add("Null");
            CurrentRatio1Twelveinches.Add("Null");
            CurrentRatio2Twelveinches.Add("Null");
            CurrentRatio3Twelveinches.Add("Null");
            CurrentRatio4Twelveinches.Add("Null");
            RESIATwelveinches.Add("Null");
            RESIBTwelveinches.Add("Null");
            RESIIATwelveinches.Add("Null");
            RESIIBTwelveinches.Add("Null");
            SNSIATwelveinches.Add("Null");
            SNSIBTwelveinches.Add("Null");
            SNSIIATwelveinches.Add("Null");
            SNSIIBTwelveinches.Add("Null");
            LEDF1Twelveinches.Add("Null");
            TTwelveinches.Add("Null");
            SecondTwelveinches.Add("Null");
            ErrorCodeTwelveinches.Add("Command Length Error");
        }

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
        }

        string str_fileName;                                                  //定义变量Excel文件名
        Microsoft.Office.Interop.Excel.Application ExcelApp;                  //声明Excel应用程序
        Workbook ExcelDoc;                                                    //声明工作簿
        Worksheet ExcelSheet;                                                 //声明工作表
        void TwelveInchesLampParametersCreatExcel()
        {
            try
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
                ExcelSheet.Cells[2, 24] = "Error Code";

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
                    if (SecondTwelveinches[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 23] = SecondTwelveinches[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 23] = ((int)SecondTwelveinches[i] / 3600).ToString() + ":" + (((int)SecondTwelveinches[i] % 3600) / 60).ToString() + ":" + (((int)SecondTwelveinches[i] % 3600) % 60).ToString();
                    }
                    ExcelSheet.Cells[3 + i, 24] = ErrorCodeTwelveinches[i].ToString();
                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearTwelveInchesLampsParameters();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));
                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("数据已导出! 保存至D盘Excel文档", "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
            }
            catch
            {
                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));
                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("导出数据失败！请重新导出", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
            }          
        }
        #endregion

        #region 8寸灯具状态信息解析
        private void EightInchesLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] received8InchesStatusFeedbackCommandArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x08 && CompleteData[i + 4] == 0x08)
                {
                    commandCount.Add(i);
                }
            }

            received8InchesStatusFeedbackCommandArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    received8InchesStatusFeedbackCommandArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    received8InchesStatusFeedbackCommandArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < received8InchesStatusFeedbackCommandArray[i].Length; j++)
                {
                    received8InchesStatusFeedbackCommandArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < received8InchesStatusFeedbackCommandArray.Length; i++)
            {
                if (received8InchesStatusFeedbackCommandArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(received8InchesStatusFeedbackCommandArray[i]);
                    if (checkOutValue == received8InchesStatusFeedbackCommandArray[i][received8InchesStatusFeedbackCommandArray[i].Length - 1])
                    {
                        RMS1Eightinches.Add(received8InchesStatusFeedbackCommandArray[i][5] * 1100);
                        Val2Eightinches.Add(received8InchesStatusFeedbackCommandArray[i][6] * 20);
                        Val3Eightinches.Add(received8InchesStatusFeedbackCommandArray[i][7]);
                        RMSEightinches.Add(received8InchesStatusFeedbackCommandArray[i][8] * 4);
                        CurrentRatio1Eightinches.Add((float)(received8InchesStatusFeedbackCommandArray[i][9] / 10.0));
                        CurrentRatio3Eightinches.Add((float)(received8InchesStatusFeedbackCommandArray[i][10] / 10.0));
                        RESIAEightinches.Add(received8InchesStatusFeedbackCommandArray[i][11] * 124);
                        RESIIAEightinches.Add(received8InchesStatusFeedbackCommandArray[i][12] * 124);
                        SNSIAEightinches.Add(received8InchesStatusFeedbackCommandArray[i][13] * 16);
                        SNSIIAEightinches.Add(received8InchesStatusFeedbackCommandArray[i][14] * 16);
                        LEDF1Eightinches.Add(received8InchesStatusFeedbackCommandArray[i][15]);
                        TEightinches.Add(received8InchesStatusFeedbackCommandArray[i][16]);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = received8InchesStatusFeedbackCommandArray[i][17 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondEightinches.Add(SecondResult);

                        ErrorCodeEightinches.Add("No Error");
                    }
                    else
                    {
                        EightInchesLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    EightInchesLampCommandLengthErrorHandle();
                }
            }
        }

        private void EightInchesLampCheckValueErrorHandle()
        {
            RMS1Eightinches.Add("Null");
            Val2Eightinches.Add("Null");
            Val3Eightinches.Add("Null");
            RMSEightinches.Add("Null");
            CurrentRatio1Eightinches.Add("Null");
            CurrentRatio3Eightinches.Add("Null");
            RESIAEightinches.Add("Null");
            RESIIAEightinches.Add("Null");
            SNSIAEightinches.Add("Null");
            SNSIIAEightinches.Add("Null");
            LEDF1Eightinches.Add("Null");
            TEightinches.Add("Null");
            SecondEightinches.Add("Null");
            ErrorCodeEightinches.Add("Check Value Error");
        }

        private void EightInchesLampCommandLengthErrorHandle()
        {
            RMS1Eightinches.Add("Null");
            Val2Eightinches.Add("Null");
            Val3Eightinches.Add("Null");
            RMSEightinches.Add("Null");
            CurrentRatio1Eightinches.Add("Null");
            CurrentRatio3Eightinches.Add("Null");
            RESIAEightinches.Add("Null");
            RESIIAEightinches.Add("Null");
            SNSIAEightinches.Add("Null");
            SNSIIAEightinches.Add("Null");
            LEDF1Eightinches.Add("Null");
            TEightinches.Add("Null");
            SecondEightinches.Add("Null");
            ErrorCodeEightinches.Add("Command Length Error");
        }

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
        }

        void EightInchesLampParametersCreatExcel()
        {
            try
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
                ExcelSheet.Cells[2, 15] = "Error Code";

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
                    if (SecondEightinches[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 14] = SecondEightinches[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 14] = ((int)SecondEightinches[i] / 3600).ToString() + ":" + (((int)SecondEightinches[i] % 3600) / 60).ToString() + ":" + (((int)SecondEightinches[i] % 3600) % 60).ToString();
                    }
                    ExcelSheet.Cells[3 + i, 15] = ErrorCodeEightinches[i].ToString();
                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearEightInchesLampsParameter();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));
                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("数据已导出!保存至D盘的Excel文档", "提示", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
            }
            catch
            {
                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));
                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show("导出数据失败！请重新导出", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }));
            }
        }

        #endregion

        #endregion

        #region 导出原始数据
        private void CreatTXT_Click(object sender, RoutedEventArgs e)
        {
            if(judgeFeedbackCommand==3)
            {
                if(ReceivedStatusFeedbackCommand.Count!=0)
                {
                    ShowTXTHandleProcess.Visibility = Visibility.Visible;
                    byte[] receivedOriginalData = new byte[ReceivedStatusFeedbackCommand.Count];
                    string OriginalData = "";

                    for (int i = 0; i < receivedOriginalData.Length; i++)
                    {
                        receivedOriginalData[i] = (byte)ReceivedStatusFeedbackCommand[i];
                        OriginalData += Convert.ToString(receivedOriginalData[i], 16).PadLeft(2, '0').ToUpper() + " ";
                    }

                    string FileName = "d:\\灯具原始数据" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                    FileStream aFile = new FileStream(FileName, FileMode.Create);
                    StreamWriter sw = new StreamWriter(aFile);

                    sw.Write(OriginalData);
                    sw.Close();

                    ShowTXTHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                    {
                        ShowTXTHandleProcess.Visibility = Visibility.Hidden;
                    }));
                    if( MessageBox.Show("数据已导出! 保存至D盘TXT文档", "提示", MessageBoxButton.OK, MessageBoxImage.Information)==MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                }
                else
                {
                    if( MessageBox.Show("接收指令不能为空！请重新查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                    ShowTXTHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                    {
                        ShowTXTHandleProcess.Visibility = Visibility.Hidden;
                    }));
                }
            }
            else
            {
                if( MessageBox.Show("未进行状态查询！请先进行状态查询", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
                ShowTXTHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowTXTHandleProcess.Visibility = Visibility.Hidden;
                }));
            }
        }                  

        #endregion

        #region 工厂模式参数设置
        private void SetLightParametersInFactoryMode_Click(object sender, RoutedEventArgs e)
        {
            
            if(ConfirmLampName.Text!=""&&ConfirmLampModel.Text!=""&&ConfirmSettingOpenCircuitParameter.Text!="")
            {
                ConfigureSettingParametersCommand();

                MessageBoxResult result = MessageBox.Show("是否将指令写入灯具？", "问询", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    ConfigurationWindow.IsEnabled = true;

                    if (lampsPort.IsOpen)
                    {
                        judgeFeedbackCommand = 1;
                        lampsPort.Write(settingParameterCommand, 0, 28);
                        
                    }
                    else
                    {
                        if( MessageBox.Show("串口未打开！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                        return;
                    }
                }

                if(result==MessageBoxResult.No)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                if(result==MessageBoxResult.None)
                {
                    ConfigurationWindow.IsEnabled = false;
                }
            }
            else
            {
                if( MessageBox.Show("参数选择不能为空！请完成配置", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
                return;
            }
        }

        #region 灯具名称
        private void SelectApproachChenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Visible;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectAPPS12SLEDC.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectApproachChenterlineLight.Content.ToString();
                ConfirmLampModel.Text = "";             
            }));
        }

        private void SelectApproachCrossbarLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Visible;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectAPPS12LLEDC.IsChecked = false;
            SelectAPPS12RLEDC.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectApproachCrossbarLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectApproachSideRowLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Visible;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectAPSS12LLEDR.IsChecked = false;
            SelectAPSS12RLEDR.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectApproachSideRowLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRWYThresholdWingBarLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Visible;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectTHWS12LLEDG.IsChecked = false;
            SelectTHWS12RLEDG.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRWYThresholdWingBarLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRWYThresholdLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Visible;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectTHRS12LLEDG.IsChecked = false;
            SelectTHRS12RLEDG.IsChecked = false;
            SelectTHRS12SLEDG.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRWYThresholdLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Visible;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectRELS12LLEDYC.IsChecked = false;
            SelectRELS12RLEDYC.IsChecked = false;
            SelectRELS12LLEDCY.IsChecked = false;
            SelectRELS12RLEDCY.IsChecked = false;
            SelectRELS12LLEDCC.IsChecked = false;
            SelectRELS12RLEDCC.IsChecked = false;
            SelectRELS12LLEDCR.IsChecked = false;
            SelectRELS12RLEDCR.IsChecked = false;
            SelectRELS12LLEDRC.IsChecked = false;
            SelectRELS12RLEDRC.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRWYEdgeLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void Select12inchesRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Visible;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectENDS12LEDR.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = Select12inchesRWYEndLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRWYThresholdEndLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Visible;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectTAES12LLEDGR1P.IsChecked = false;
            SelectTAES12RLEDGR1P.IsChecked = false;
            SelectTAES12SLEDGR1P.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRWYThresholdEndLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRWYCenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Visible;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectRCLS08LEDCB1P.IsChecked = false;
            SelectRCLS08LEDRB1P.IsChecked = false;
            SelectRCLS08LEDCC1P.IsChecked = false;
            SelectRCLS08LEDRC1P.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRWYCenterlineLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRWYTouchdownZoneLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Visible;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectTDZS08LLEDC.IsChecked = false;
            SelectTDZS08RLEDC.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRWYTouchdownZoneLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void Select8inchesRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Visible;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectENDS08LEDR.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = Select8inchesRWYEndLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectRapidExitTWYIndicatorLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Visible;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;

            SelectRAPS08LEDY.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectRapidExitTWYIndicatorLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }

        private void SelectCombinedRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            GroupApproachChenterlineLight.Visibility = Visibility.Collapsed;
            GroupApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupRWYEdgeLight.Visibility = Visibility.Collapsed;
            Group12inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRWYThresholdEndLight.Visibility = Visibility.Collapsed;
            GroupRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Visible;

            SelectRELC12LEDCYC1P.IsChecked = false;
            SelectRELC12LEDCCC1P.IsChecked = false;
            SelectRELC12LEDCRC1P.IsChecked = false;
            SelectRELC12LEDRYC1P.IsChecked = false;
            SelectRELC12LEDCYB1P.IsChecked = false;
            SelectRELC12LEDCCB1P.IsChecked = false;
            SelectRELC12LEDCRB1P.IsChecked = false;
            SelectRELC12LEDRYB1P.IsChecked = false;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = SelectCombinedRWYEdgeLight.Content.ToString();
                ConfirmLampModel.Text = "";
            }));
        }
        #endregion

        #region 灯具型号
        private void SelectAPPS12SLEDC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPPS12SLEDC.Content.ToString();
            }));

            ConfigureAPPS12SLEDCParameters();
        }

        private void SelectAPPS12LLEDC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPPS12LLEDC.Content.ToString();
            }));

            ConfigureAPPS12LLEDCParameters();
        }

        private void SelectAPPS12RLEDC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPPS12RLEDC.Content.ToString();
            }));

            ConfigureAPPS12RLEDCParameters();
        }

        private void SelectAPSS12LLEDR_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPSS12LLEDR.Content.ToString();
            }));

            ConfigureAPSS12LLEDRParameters();
        }

        private void SelectAPSS12RLEDR_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPSS12RLEDR.Content.ToString();
            }));

            ConfigureAPSS12RLEDRParameters();
        }

        private void SelectTHWS12LLEDG_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHWS12LLEDG.Content.ToString();
            }));

            ConfigureTHWS12LLEDGParameters();
        }

        private void SelectTHWS12RLEDG_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHWS12RLEDG.Content.ToString();
            }));

            ConfigureTHWS12RLEDGParameters();
        }

        private void SelectTHRS12LLEDG_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHRS12LLEDG.Content.ToString();
            }));

            ConfigureTHRS12LLEDGParameters();
        }

        private void SelectTHRS12RLEDG_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHRS12RLEDG.Content.ToString();
            }));

            ConfigureTHRS12RLEDGParameters();
        }

        private void SelectTHRS12SLEDG_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHRS12SLEDG.Content.ToString();
            }));

            ConfigureTHRS12SLEDGParameters();
        }

        private void SelectRELS12LLEDYC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDYC.Content.ToString();
            }));

            ConfigureRELS12LLEDYCParameters();
        }

        private void SelectRELS12RLEDYC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDYC.Content.ToString();
            }));

            ConfigureRELS12RLEDYCParameters();
        }

        private void SelectRELS12LLEDCY_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDCY.Content.ToString();
            }));

            ConfigureRELS12LLEDCYParameters();
        }

        private void SelectRELS12RLEDCY_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDCY.Content.ToString();
            }));

            ConfigureRELS12RLEDCYParameters();
        }

        private void SelectRELS12LLEDCC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDCC.Content.ToString();
            }));

            ConfigureRELS12LLEDCCParameters();
        }

        private void SelectRELS12RLEDCC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDCC.Content.ToString();
            }));

            ConfigureRELS12RLEDCCParameters();
        }

        private void SelectRELS12LLEDCR_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDCR.Content.ToString();
            }));

            ConfigureRELS12LLEDCRParameters();
        }

        private void SelectRELS12RLEDCR_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDCR.Content.ToString();
            }));

            ConfigureRELS12RLEDCRParameters();
        }

        private void SelectRELS12LLEDRC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDRC.Content.ToString();
            }));

            ConfigureRELS12LLEDRCParameters();
        }

        private void SelectRELS12RLEDRC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDRC.Content.ToString();
            }));

            ConfigureRELS12RLEDRCParameters();
        }

        private void SelectENDS12LEDR_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectENDS12LEDR.Content.ToString();
            }));

            ConfigureENDS12LEDRParameters();
        }

        private void SelectTAES12LLEDGR1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTAES12LLEDGR1P.Content.ToString();
            }));

            ConfigureTAES12LLEDGR1PParameters();
        }

        private void SelectTAES12RLEDGR1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTAES12RLEDGR1P.Content.ToString();
            }));

            ConfigureTAES12RLEDGR1PParameters();
        }

        private void SelectTAES12SLEDGR1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTAES12SLEDGR1P.Content.ToString();
            }));

            ConfigureTAES12SLEDGR1PParameters();
        }

        private void SelectRCLS08LEDCB1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDCB1P.Content.ToString();
            }));

            ConfigureRCLS08LEDCB1PParameters();
        }

        private void SelectRCLS08LEDRB1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDRB1P.Content.ToString();
            }));

            ConfigureRCLS08LEDRB1PParameters();
        }

        private void SelectRCLS08LEDCC1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDCC1P.Content.ToString();
            }));

            ConfigureRCLS08LEDCC1PParameters();
        }

        private void SelectRCLS08LEDRC1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDRC1P.Content.ToString();
            }));

            ConfigureRCLS08LEDRC1PParameters();
        }

        private void SelectTDZS08LLEDC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTDZS08LLEDC.Content.ToString();
            }));

            ConfigureTDZS08LLEDCParameters();
        }

        private void SelectTDZS08RLEDC_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTDZS08RLEDC.Content.ToString();
            }));

            ConfigureTDZS08RLEDCParameters();
        }

        private void SelectENDS08LEDR_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectENDS08LEDR.Content.ToString();
            }));

            ConfigureENDS08LEDRParameters();
        }

        private void SelectRAPS08LEDY_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRAPS08LEDY.Content.ToString();
            }));

            ConfigureRAPS08LEDYParameters();
        }

        private void SelectRELC12LEDCYC1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCYC1P.Content.ToString();
            }));

            ConfigureRELC12LEDCYC1PParameters();
        }

        private void SelectRELC12LEDCCC1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCCC1P.Content.ToString();
            }));

            ConfigureRELC12LEDCCC1PParameters();
        }

        private void SelectRELC12LEDCRC1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCRC1P.Content.ToString();
            }));

            ConfigureRELC12LEDCRC1PParameters();
        }

        private void SelectRELC12LEDRYC1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDRYC1P.Content.ToString();
            }));

            ConfigureRELC12LEDRYC1PParameters();
        }

        private void SelectRELC12LEDCYB1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCYB1P.Content.ToString();
            }));

            ConfigureRELC12LEDCYB1PParameters();
        }

        private void SelectRELC12LEDCCB1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCCB1P.Content.ToString();
            }));

            ConfigureRELC12LEDCCB1PParameters();
        }

        private void SelectRELC12LEDCRB1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCRB1P.Content.ToString();
            }));

            ConfigureRELC12LEDCRB1PParameters();
        }

        private void SelectRELC12LEDRYB1P_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDRYB1P.Content.ToString();
            }));

            ConfigureRELC12LEDRYB1PParameters();
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

        #region 工厂模式配置灯具参数
        private void ConfigureAPPS12SLEDCParameters()
        {
            settingIA[0] = 0x01;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x01;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x01;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x01;               
        }

        private void ConfigureAPPS12LLEDCParameters()
        {
            settingIA[0] = 0x01;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x01;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x01;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x02;
        }

        private void ConfigureAPPS12RLEDCParameters()
        {
            settingIA[0] = 0x01;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x01;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x01;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x03;
        }

        private void ConfigureAPSS12LLEDRParameters()
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

        private void ConfigureAPSS12RLEDRParameters()
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

        private void ConfigureTHWS12LLEDGParameters()
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

        private void ConfigureTHWS12RLEDGParameters()
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

        private void ConfigureTHRS12LLEDGParameters()
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

        private void ConfigureTHRS12RLEDGParameters()
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

        private void ConfigureTHRS12SLEDGParameters()
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

        private void ConfigureRELS12LLEDYCParameters()
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

        private void ConfigureRELS12RLEDYCParameters()
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

        private void ConfigureRELS12LLEDCYParameters()
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
            settingLampsNumber = 0x0D;
        }

        private void ConfigureRELS12RLEDCYParameters()
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
            settingLampsNumber = 0x0E;
        }

        private void ConfigureRELS12LLEDCCParameters()
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
            settingLampsNumber = 0x0F;
        }

        private void ConfigureRELS12RLEDCCParameters()
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
            settingLampsNumber = 0x10;
        }

        private void ConfigureRELS12LLEDCRParameters()
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
            settingLampsNumber = 0x11;
        }

        private void ConfigureRELS12RLEDCRParameters()
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
            settingLampsNumber = 0x12;
        }

        private void ConfigureRELS12LLEDRCParameters()
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
            settingLampsNumber = 0x13;
        }

        private void ConfigureRELS12RLEDRCParameters()
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
            settingLampsNumber = 0x14;
        }

        private void ConfigureENDS12LEDRParameters()
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
            settingLampsNumber = 0x15;
        }

        private void ConfigureTAES12LLEDGR1PParameters()
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
            settingLampsNumber = 0x16;
        }

        private void ConfigureTAES12RLEDGR1PParameters()
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
            settingLampsNumber = 0x17;
        }

        private void ConfigureTAES12SLEDGR1PParameters()
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
            settingLampsNumber = 0x18;
        }

        private void ConfigureRCLS08LEDCB1PParameters()
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
            settingLampsNumber = 0x19;
        }

        private void ConfigureRCLS08LEDRB1PParameters()
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
            settingLampsNumber = 0x1A;
        }

        private void ConfigureRCLS08LEDCC1PParameters()
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
            settingLampsNumber = 0x1B;
        }

        private void ConfigureRCLS08LEDRC1PParameters()
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
            settingLampsNumber = 0x1C;
        }

        private void ConfigureTDZS08LLEDCParameters()
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
            settingLampsNumber = 0x1D;
        }

        private void ConfigureTDZS08RLEDCParameters()
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
            settingLampsNumber = 0x1E;
        }

        private void ConfigureENDS08LEDRParameters()
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
            settingLampsNumber = 0x1F;
        }

        private void ConfigureRAPS08LEDYParameters()
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
            settingLampsNumber = 0x20;
        }

        private void ConfigureRELC12LEDCYC1PParameters()
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
            settingLampsNumber = 0x21;
        }

        private void ConfigureRELC12LEDCCC1PParameters()
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
            settingLampsNumber = 0x22;
        }

        private void ConfigureRELC12LEDCRC1PParameters()
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
            settingLampsNumber = 0x23;
        }

        private void ConfigureRELC12LEDRYC1PParameters()
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
            settingLampsNumber = 0x24;
        }

        private void ConfigureRELC12LEDCYB1PParameters()
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
            settingLampsNumber = 0x25;
        }

        private void ConfigureRELC12LEDCCB1PParameters()
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
            settingLampsNumber = 0x26;
        }

        private void ConfigureRELC12LEDCRB1PParameters()
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
            settingLampsNumber = 0x27;
        }

        private void ConfigureRELC12LEDRYB1PParameters()
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
            settingLampsNumber = 0x28;
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

        #region 工厂模式下，生成设置参数指令
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

        #endregion

        #region 开发者模式设置参数
        private void SetLightParametersInDeveloperMode_Click(object sender, RoutedEventArgs e)
        {
            ShowSetParameterCommand.Text = "";
            if (SetParameterIA.Text != "" && SetParameterIB.Text != "" && SetParameterIIA.Text != "" && SetParameterIIB.Text != "")
            {
                InDeveloperModeConfigureSettingParametersCommand();

                for (int i = 0; i < InDeveloperModeSettingParameterCommand.Length; i++)
                {
                    ShowSetParameterCommand.Text += Convert.ToString(InDeveloperModeSettingParameterCommand[i], 16).PadLeft(2, '0').ToUpper() + " ";
                }

                MessageBoxResult result = MessageBox.Show("是否将指令写入灯具？", "问询", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if ( result== MessageBoxResult.Yes)
                {
                    ConfigurationWindow.IsEnabled = true;

                    if (lampsPort.IsOpen)
                    {
                        judgeFeedbackCommand = 1;
                        lampsPort.Write(InDeveloperModeSettingParameterCommand, 0, 28);
                    }
                    else
                    {
                        if( MessageBox.Show("串口未打开！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                        return;
                    }
                }

                if(result==MessageBoxResult.None)
                {
                    ConfigurationWindow.IsEnabled = false;
                }
                if(result==MessageBoxResult.No)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
            }
            else
            {
                if( MessageBox.Show("文本框不能为空！请输入电流值", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
                return;
            }
        }

        #region 参数设置
        private void SetParameterIA_TextChanged(object sender, TextChangedEventArgs e)
        {
            InDeveloperModeSettingIA = CalculateCurrentBuffer(SetParameterIA.Text.ToString(), 1);
        }

        private void SetParameterIB_TextChanged(object sender, TextChangedEventArgs e)
        {
            InDeveloperModeSettingIB = CalculateCurrentBuffer(SetParameterIB.Text.ToString(), 2);
        }

        private void SetParameterIIA_TextChanged(object sender, TextChangedEventArgs e)
        {
            InDeveloperModeSettingIIA = CalculateCurrentBuffer(SetParameterIIA.Text.ToString(), 3);
        }

        private void SetParameterIIB_TextChanged(object sender, TextChangedEventArgs e)
        {
            InDeveloperModeSettingIIB = CalculateCurrentBuffer(SetParameterIIB.Text.ToString(), 4);
        }

        private void SetReadResistanceFalse_Checked(object sender, RoutedEventArgs e)
        {
            InDeveloperModeSettingReadRFlag = 0x00;
        }

        private void SetReadResistanceTrue_Checked(object sender, RoutedEventArgs e)
        {
            InDeveloperModeSettingReadRFlag = 0x01;
        }

        private void SetMosfetTrue_Checked(object sender, RoutedEventArgs e)
        {
            InDeveloperModeSettingMosFlag = 0x01;
        }

        private void SetMosfetFalse_Checked(object sender, RoutedEventArgs e)
        {
            InDeveloperModeSettingMosFlag = 0x00;
        }

        private void SetBreakTrue_Checked(object sender, RoutedEventArgs e)
        {
            InDeveloperModeSettingBreakFlag = 0x01;
        }

        private void SetBreakFalse_Checked(object sender, RoutedEventArgs e)
        {
            InDeveloperModeSettingBreakFlag = 0x00;
        }

        private void SetLightNumber_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            InDeveloperModeSettingLampsNumber = Convert.ToByte(SetLightNumber.SelectedIndex+1);
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
                    if( MessageBox.Show("非法输入！请输入数字 0~9、字符 '.'", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
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
                                    if( MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                                    {
                                        ConfigurationWindow.IsEnabled = true;
                                    }
                                    else
                                    {
                                        ConfigurationWindow.IsEnabled = false;
                                    }
                                    PrugeTextBoxContent(textboxNumber);
                                    return result;
                                }
                                else
                                {
                                    if (charCurrentValue[0] == '1' && charCurrentValue[2] != '0')
                                    {
                                        if( MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                                        {
                                            ConfigurationWindow.IsEnabled = true;
                                        }
                                        else
                                        {
                                            ConfigurationWindow.IsEnabled = false;
                                        }
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
                                    if( MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                                    {
                                        ConfigurationWindow.IsEnabled = true;
                                    }
                                    else
                                    {
                                        ConfigurationWindow.IsEnabled = false;
                                    }
                                    PrugeTextBoxContent(textboxNumber);
                                    return result;
                                }
                                else
                                {
                                    if (charCurrentValue[0] == '1' && (charCurrentValue[2] != '0' || charCurrentValue[3] != '0'))
                                    {
                                        if( MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                                        {
                                            ConfigurationWindow.IsEnabled = true;
                                        }
                                        else
                                        {
                                            ConfigurationWindow.IsEnabled = false;
                                        }
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
                                    if( MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                                    {
                                        ConfigurationWindow.IsEnabled = true;
                                    }
                                    else
                                    {
                                        ConfigurationWindow.IsEnabled = false;
                                    }
                                    PrugeTextBoxContent(textboxNumber);
                                    return result;
                                }
                                else
                                {
                                    if (charCurrentValue[0] == '1' && (charCurrentValue[2] != '0' || charCurrentValue[3] != '0' || charCurrentValue[4] != '0'))
                                    {
                                        if( MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                                        {
                                            ConfigurationWindow.IsEnabled = true;
                                        }
                                        else
                                        {
                                            ConfigurationWindow.IsEnabled = false;
                                        }
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
                            if( MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }
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
                        if( MessageBox.Show("位数错误！小数点后最多保留三位，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }
                        PrugeTextBoxContent(textboxNumber);
                        return result;
                    }
                }
                else if (charCurrentValue[0] == '.')
                {
                    if( MessageBox.Show("格式错误！请输入正确格式，例：0.123", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                    PrugeTextBoxContent(textboxNumber);
                    return result;
                }
                else
                {
                    if( MessageBox.Show("已超出量程！请输入数值范围为0~1", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                    PrugeTextBoxContent(textboxNumber);
                    return result;
                }

            }

            return result;
        }

        #endregion

        #region 在生成电流数组函数中清空指定文本框
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

        #region 开发者模式下，生成设置参数指令
        public void InDeveloperModeConfigureSettingParametersCommand()
        {
            for (int i = 0; i < InDeveloperModeSettingIA.Length; i++)
            {
                InDeveloperModeSettingParameterCommand[5 + i] = InDeveloperModeSettingIA[i];
            }
            for (int i = 0; i < InDeveloperModeSettingIB.Length; i++)
            {
                InDeveloperModeSettingParameterCommand[9 + i] = InDeveloperModeSettingIB[i];
            }
            for (int i = 0; i < InDeveloperModeSettingIIA.Length; i++)
            {
                InDeveloperModeSettingParameterCommand[13 + i] = InDeveloperModeSettingIIA[i];
            }
            for (int i = 0; i < InDeveloperModeSettingIIB.Length; i++)
            {
                InDeveloperModeSettingParameterCommand[17 + i] = InDeveloperModeSettingIIB[i];
            }
            InDeveloperModeSettingParameterCommand[21] = InDeveloperModeSettingReadRFlag;
            InDeveloperModeSettingParameterCommand[22] = InDeveloperModeSettingMosFlag;
            InDeveloperModeSettingParameterCommand[23] = InDeveloperModeSettingBreakFlag;
            InDeveloperModeSettingParameterCommand[24] = InDeveloperModeSettingLampsNumber;
            InDeveloperModeSettingParameterCommand[27] = CalculateCheckOutValue(InDeveloperModeSettingParameterCommand);
        }
        #endregion

        #endregion
      
        #region 用户登录
        string userName;
        string password;
        private void LogIn_Click(object sender, RoutedEventArgs e)
        {
            userName = UserName.Text.ToString().Trim();
            password = Password.Password;
            
            if (userName =="Airsafe")
            {
                if(password == "Airsafe")
                {
                    DeveloperMode.Visibility = Visibility.Visible;
                    DeveloperMode.IsSelected = true;
                    LogIn.IsEnabled = false;                  
                }
                else
                {
                    if(MessageBox.Show("密码错误！请重新输入密码", "提示", MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
                    
                    Password.Password = "";
                }                
            }
            else
            {
                if( MessageBox.Show("账号错误！请重新输入账号", "提示", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                {
                    ConfigurationWindow.IsEnabled = true;
                }
                else
                {
                    ConfigurationWindow.IsEnabled = false;
                }
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
            PurgingDeveloperMode();
            LogIn.IsEnabled = true;
            FactoryMode.IsSelected = true;
        }
        #endregion

        #region 清空页面       
        private void PurgingFactoryMode()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = "";
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
                AnswerIA.Text = "";
                AnswerIB.Text = "";
                AnswerIIA.Text = "";
                AnswerIIB.Text = "";
                AnswerLampModel.Text = "";
                AnswerOpenCircuit.Text = "";
                AnswerVersion.Text = "";

                SelectApproachChenterlineLight.IsChecked = false;
                SelectApproachCrossbarLight.IsChecked = false;
                SelectApproachSideRowLight.IsChecked = false;
                SelectRWYThresholdWingBarLight.IsChecked = false;
                SelectRWYThresholdLight.IsChecked = false;
                SelectRWYEdgeLight.IsChecked = false;
                Select12inchesRWYEndLight.IsChecked = false;
                SelectRWYThresholdEndLight.IsChecked = false;
                SelectRWYCenterlineLight.IsChecked = false;
                SelectRWYTouchdownZoneLight.IsChecked = false;
                Select8inchesRWYEndLight.IsChecked = false;
                SelectRapidExitTWYIndicatorLight.IsChecked = false;
                SelectCombinedRWYEdgeLight.IsChecked = false;

                SelectAPPS12SLEDC.IsChecked = false;
                SelectAPPS12LLEDC.IsChecked = false;
                SelectAPPS12RLEDC.IsChecked = false;
                SelectAPSS12LLEDR.IsChecked = false;
                SelectAPSS12RLEDR.IsChecked = false;
                SelectTHWS12LLEDG.IsChecked = false;
                SelectTHWS12RLEDG.IsChecked = false;
                SelectTHRS12LLEDG.IsChecked = false;
                SelectTHRS12RLEDG.IsChecked = false;
                SelectTHRS12SLEDG.IsChecked = false;
                SelectRELS12LLEDYC.IsChecked = false;
                SelectRELS12RLEDYC.IsChecked = false;
                SelectRELS12LLEDCY.IsChecked = false;
                SelectRELS12RLEDCY.IsChecked = false;
                SelectRELS12LLEDCC.IsChecked = false;
                SelectRELS12RLEDCC.IsChecked = false;
                SelectRELS12LLEDCR.IsChecked = false;
                SelectRELS12RLEDCR.IsChecked = false;
                SelectRELS12LLEDRC.IsChecked = false;
                SelectRELS12RLEDRC.IsChecked = false;
                SelectENDS12LEDR.IsChecked = false;
                SelectTAES12LLEDGR1P.IsChecked = false;
                SelectTAES12RLEDGR1P.IsChecked = false;
                SelectTAES12SLEDGR1P.IsChecked = false;
                SelectRCLS08LEDCB1P.IsChecked = false;
                SelectRCLS08LEDRB1P.IsChecked = false;
                SelectRCLS08LEDCC1P.IsChecked = false;
                SelectRCLS08LEDRC1P.IsChecked = false;
                SelectTDZS08LLEDC.IsChecked = false;
                SelectTDZS08RLEDC.IsChecked = false;
                SelectENDS08LEDR.IsChecked = false;
                SelectRAPS08LEDY.IsChecked = false;
                SelectRELC12LEDCYC1P.IsChecked = false;
                SelectRELC12LEDCCC1P.IsChecked = false;
                SelectRELC12LEDCRC1P.IsChecked = false;
                SelectRELC12LEDRYC1P.IsChecked = false;
                SelectRELC12LEDCYB1P.IsChecked = false;
                SelectRELC12LEDCCB1P.IsChecked = false;
                SelectRELC12LEDCRB1P.IsChecked = false;
                SelectRELC12LEDRYB1P.IsChecked = false;

                SelectOpenCircuitTrue.IsChecked = false;
                SelectOpenCircuitFalse.IsChecked = false;

                SelectApproachChenterlineLight.IsEnabled = true;
                SelectApproachCrossbarLight.IsEnabled = true;
                SelectApproachSideRowLight.IsEnabled = true;
                SelectRWYThresholdWingBarLight.IsEnabled = true;
                SelectRWYThresholdLight.IsEnabled = true;
                SelectRWYEdgeLight.IsEnabled = true;
                Select12inchesRWYEndLight.IsEnabled = true;
                SelectRWYThresholdEndLight.IsEnabled = true;
                SelectRWYCenterlineLight.IsEnabled = true;
                SelectRWYTouchdownZoneLight.IsEnabled = true;
                Select8inchesRWYEndLight.IsEnabled = true;
                SelectRapidExitTWYIndicatorLight.IsEnabled = true;
                SelectCombinedRWYEdgeLight.IsEnabled = true;
            }));
        }

        private void PurgingDeveloperMode()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                SetParameterIA.Text = "";
                SetParameterIB.Text = "";
                SetParameterIIA.Text = "";
                SetParameterIIB.Text = "";
                ShowSetParameterCommand.Text = "";
                AnswerStatus.Text = "";
            }));
        }
        #endregion

        #region UI操作 按Enter键得到Tab的效果
        private void Grid_PreviewKeyDown(object sender,KeyEventArgs e)
        {
            var uie = e.OriginalSource as UIElement;

            if(e.Key==Key.Enter)
            {
                e.Handled = true;
                uie.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
            }
        }

        #endregion

        #region UI操作 鼠标单击获取焦点后的全选
        private void SetParameterIA_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.SelectAll();
            tb.PreviewMouseDown -= new MouseButtonEventHandler(SetParameterIA_PreviewMouseDown);
        }

        private void SetParameterIA_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.Focus();
            e.Handled = true;
        }

        private void SetParameterIA_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.PreviewMouseDown += new MouseButtonEventHandler(SetParameterIA_PreviewMouseDown);
        }

       

        private void SetParameterIB_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.SelectAll();
            tb.PreviewMouseDown -= new MouseButtonEventHandler(SetParameterIB_PreviewMouseDown);
        }

        private void SetParameterIB_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.Focus();
            e.Handled = true;
        }

        private void SetParameterIB_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.PreviewMouseDown += new MouseButtonEventHandler(SetParameterIB_PreviewMouseDown);
        }
     

        private void SetParameterIIA_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.PreviewMouseDown += new MouseButtonEventHandler(SetParameterIIA_PreviewMouseDown);
        }

        private void SetParameterIIA_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.Focus();
            e.Handled = true;
        }

        private void SetParameterIIA_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.SelectAll();
            tb.PreviewMouseDown -= new MouseButtonEventHandler(SetParameterIIA_PreviewMouseDown);
        }


        private void SetParameterIIB_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.SelectAll();
            tb.PreviewMouseDown -= new MouseButtonEventHandler(SetParameterIIB_PreviewMouseDown);
        }

        private void SetParameterIIB_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.Focus();
            e.Handled = true;
        }

        private void SetParameterIIB_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox tb = e.Source as System.Windows.Controls.TextBox;
            tb.PreviewMouseDown += new MouseButtonEventHandler(SetParameterIIB_PreviewMouseDown);
        }

        #endregion
    }
}
