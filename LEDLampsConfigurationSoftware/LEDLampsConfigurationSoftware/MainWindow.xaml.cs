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

namespace LEDLampsConfigurationSoftware
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 设置全局变量
        SerialPort lampsPort = new SerialPort();  //定义串口
        bool isQueryStatus = false;  //定义状态查询标识符
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

            ////test
            //double test = 0.123;
            //byte[] bytetest;
            //bytetest = CalculateCurrentBuffer(test);
            //for(int i=0;i<bytetest.Length;i++)
            //{
            //    AnswerStatus.Text += bytetest[i].ToString()+"";
            //}

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
                lampsPort.Write(queryStatusCommand, 0, 28);
            }
            else
            {
                MessageBox.Show("未打开串口！请打开串口", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region 工厂模式参数设置
        private void SetLightParametersInFactoryMode_Click(object sender, RoutedEventArgs e)
        {
            AnswerStatus.Text = "";
            if(ConfirmSettingLampParameter.Text!=""&&ConfirmSettingSpecialLampParameter.Text!=""&&ConfirmSettingOpenCircuitParameter.Text!="")
            {
                ConfigureSettingParametersCommand();

                for (int i = 0; i < settingParameterCommand.Length; i++)
                {
                    AnswerStatus.Text += Convert.ToString(settingParameterCommand[i], 16).PadLeft(2, '0').ToUpper() + " ";
                }

                if (MessageBox.Show("是否将指令写入灯具？", "问询", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    if (lampsPort.IsOpen)
                    {
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
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

            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingLampParameter.Text = SelectCombinedRWYEdgeLight.Content.ToString();
                ConfirmSettingSpecialLampParameter.Text = "";
            }));
        }
    #endregion

        #region 特殊灯具选择
        private void SelectSpecialWhiteYellowAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {               
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteYellowAllAround.Content.ToString();
            }));

            ConfigureSpecialWhiteYellowAllAroundParameters();
        }

        private void SelectSpecialWhiteWhiteAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteWhiteAllAround.Content.ToString();
            }));

            ConfigureSpecialWhiteWhiteAllAroundParameters();
        }

        private void SelectSpecialWhiteRedAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteRedAllAround.Content.ToString();
            }));

            ConfigureSpecialWhiteRedAllAroundParameters();
        }

        private void SelectSpecialRedYellowAllAround_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialRedYellowAllAround.Content.ToString();
            }));

            ConfigureSpecialRedYellowAllAroundParameters();
        }

        private void SelectSpecialWhiteYellow_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteYellow.Content.ToString();
            }));

            ConfigureSpecialWhiteYellowParameters();
        }

        private void SelectSpecialWhiteWhite_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteWhite.Content.ToString();
            }));

            ConfigureSpecialWhiteWhiteParameters();
        }

        private void SelectSpecialWhiteRed_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialWhiteRed.Content.ToString();
            }));

            ConfigureSpecialWhiteRedParameters();
        }

        private void SelectSpecialRedYellow_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingSpecialLampParameter.Text = SelectSpecialRedYellow.Content.ToString();
            }));

            ConfigureSpecialRedYellowParameters();
        }
    #endregion

        #region 开路选择
        private void SelectOpenCircuitTrue_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                ConfirmSettingOpenCircuitParameter.Text = SelectOpenCircuitTrue.Content.ToString();
            }));

            ConfigureOpenCircuitTrue();
        }

        private void SelectOpenCircuitFalse_Checked(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(() =>
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

                                this.Dispatcher.Invoke(new Action(() =>
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


    }
}
