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
        int judgeFeedbackCommand = 0;  //设置参数指令为1，版本查询指令为2，状态查询指令为3，无反馈指令为0，打开串口时发送版本查询指令为4,总时间查询指令为5

        #endregion

        #region 工厂模式
        #region 版本反馈指令参数
        int year = 0;
        int month = 0;
        int date = 0;
        int softwareVersion1 = 0;
        int softwareVersion2 = 0;
        int softwareVersion3 = 0;
        double currentRatio1 = 0;
        double currentRatio2 = 0;
        double currentRatio3 = 0;
        double currentRatio4 = 0;
        int breakFlag = 0;
        byte lampsNumber = 0;
        int hardwareVersion1 = 0;
        int hardwareVersion2 = 0;
        int hardwareVersion3 = 0;
        int softwareNumber = 0;

        //跑道警戒灯特有
        int channelSelect = 0;
        int flashFrequency = 0;
        int waveformSelect = 0;
        #endregion

        #region 总时间查询指令参数
        int breakDownCount = 0;
        int totalTime = 0;
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
        byte settingChannel = 0x00;  //通道选择
        byte settingFlashFrequency = 0x00;  //闪光频率
        byte settingWaveform = 0x00;  //波形选择
        byte settingBreakVal1 = 0x00;
        byte settingBreakVal2 = 0x00;
        byte settingORIVOLT = 0X00;
        byte[] settingRMSSET = new byte[2] { 0X00, 0X00 };
        byte settingIICFLAG = 0X00;
        #endregion

        #region 工厂模式下，其他设置参数
        int[] FlashFrequencyArray = new int[31];
        #endregion

        #region 总时间查询按钮适用对象，以及版本查询电流值计算方法
        /* 第一列为<hardware_version1>，驱动识别号的一部分
         * 第二列为<S_Version>，驱动识别号的一部分
         * 第三列为<software_version1>，软件版本号的一部分
         * 第四列为<software_version2>，软件版本号的一部分
         * 第五列为总时间查询功能是否适用，适用为：1，不适用为：2
         * 第六列为真实电流计算公式的符号，乘号为：1，除号为：2
         * 第七列为真实电流计算公式的系数         
         */
        double[][] totalTimeObject = new double[][]
        {
            new double[7]{4,5,1,3,1,2,1.72},
            new double[7]{5,6,1,4,1,2,1.6 },
            new double[7]{5,6,1,3,1,2,1.96 },
            new double[7]{5,6,1,2,1,2,1.65 },
            new double[7]{8,0,1,3,2,1,0.66 },
            new double[7]{8,0,1,4,1,1,0.66},
            new double[7]{8,4,1,1,2,1,0.66 },
            new double[7]{9,9,2,0,1,1,0.66},
            new double[7]{12,0,1,3,2,1,1 },
            new double[7]{12,0,1,2,2,1,1 },
            new double[7]{12,3,1,0,2,1,1 },
            new double[7]{13,2,1,0,2,1,0.66 }            
        };
        #endregion
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

        //跑道警戒灯设置参数
        byte InDeveloperModeChannelSelectContent = 0x00;
        byte InDeveloperModeFlashFrequencyContent = 0x00;
        byte InDeveloperModeWaveformSelectContent = 0x00;
        #endregion

        #region 发送指令集（尚未计算校验值）
        Byte[] queryStatusCommand = new Byte[28] { 0x02,0x89,0x11,0x58,0x12,0x00,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x06,0x00,0x00,0x00,0x00 };
        Byte[] queryVersionCommand = new Byte[28] { 0x02, 0x89, 0x22, 0x85, 0x12, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00 };
        Byte[] settingParameterCommand = new Byte[28] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        Byte[] InDeveloperModeSettingParameterCommand = new Byte[28] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        Byte[] InFactoryModeCommonLightRestoreOriginalCommand = new Byte[28] { 0x02, 0x55, 0x11, 0x58, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00 };
        Byte[] InFactoryModeRWYGuardLightRestoreOriginalCommand = new Byte[28] { 0x55, 0x02, 0x11, 0x58, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x1F, 0x00, 0x00 };
        Byte[] queryTotalTimeCommand = new Byte[28] { 0x02, 0x89, 0x58, 0x22, 0x12, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00 };

        #endregion

        #region 状态反馈指令参数集
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

        #region 8寸V2灯具各项参数存储集合
        ArrayList RMS1EightinchesV2 = new ArrayList();
        ArrayList Val2EightinchesV2 = new ArrayList();
        ArrayList TCheckEightinchesV2 = new ArrayList();
        ArrayList RMSEightinchesV2 = new ArrayList();
        ArrayList CurrentRatio1EightinchesV2 = new ArrayList();
        ArrayList CurrentRatio3EightinchesV2 = new ArrayList();
        ArrayList RESIAEightinchesV2 = new ArrayList();
        ArrayList RESIIAEightinchesV2 = new ArrayList();
        ArrayList SNSIAEightinchesV2 = new ArrayList();
        ArrayList SNSIIAEightinchesV2 = new ArrayList();
        ArrayList LEDF1EightinchesV2 = new ArrayList();
        ArrayList TEightinchesV2 = new ArrayList();
        ArrayList SecondEightinchesV2 = new ArrayList();
        ArrayList Shock1EightinchesV2 = new ArrayList();
        ArrayList ErrorCodeEightinchesV2 = new ArrayList();
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

        #region 双路跑中驱动灯具各项参数存储集合
        ArrayList RMS1DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList Val2DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RMS2DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList CurrentRatio1DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList CurrentRatio2DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList CurrentRatio3DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList CurrentRatio4DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RMS1LASTDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RMS2LASTDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList SNSIADoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList SNSIBDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList SNSIIADoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList SNSIIBDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList LEDF1DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList LEDF2DoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RESIADoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RESIBDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RESIIADoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList RESIIBDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList TDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList SecondDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList TCHECKDoubleCircuitRWYCenterDrive = new ArrayList();
        ArrayList ErrorCodeDoubleCircuitRWYCenterDrive = new ArrayList();
        #endregion

        #region 12寸双路跑中驱动灯具各项参数存储集合
        ArrayList RMS1DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RMS2DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList Val2DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList Val3DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RMSMID1DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RMSMID2DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RMS1LASTDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList CurrentRatio1DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList CurrentRatio2DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList CurrentRatio3DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList CurrentRatio4DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RESIADoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RESIBDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RESIIADoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RESIIBDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList SNSIADoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList SNSIBDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList SNSIIADoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList LEDF2DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList LEDF1DoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList TDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList SecondDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList RMS2LASTDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        ArrayList ErrorCodeDoubleCircuitRWYCenterDrive12inches = new ArrayList();
        #endregion

        #region 8寸警戒灯灯具各项参数存储集合
        ArrayList RMS1RWYGuardLight = new ArrayList();
        ArrayList Val2RWYGuardLight = new ArrayList();
        ArrayList Val3RWYGuardLight = new ArrayList();
        ArrayList RMSRWYGuardLight = new ArrayList();
        ArrayList CurrentRatio1RWYGuardLight = new ArrayList();
        ArrayList CurrentRatio3RWYGuardLight = new ArrayList();
        ArrayList WaveformRWYGuardLight = new ArrayList();
        ArrayList ChannelRWYGuardLight = new ArrayList();
        ArrayList SNSIARWYGuardLight = new ArrayList();
        ArrayList SNSIIARWYGuardLight = new ArrayList();
        ArrayList LEDF1RWYGuardLight = new ArrayList();
        ArrayList TRWYGuardLight = new ArrayList();
        ArrayList SecondRWYGuardLight = new ArrayList();
        ArrayList FlashFrequencyRWYGuardLight = new ArrayList();
        ArrayList ModeRWYGuardLight = new ArrayList();
        ArrayList ErrorCodeRWYGuardLight = new ArrayList();
        #endregion

        #region 双路滑中驱动各项参数存储集合
        ArrayList RMS1DoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList RMS2DoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList RMS1LASTDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList RMS2LASTDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList Val2DoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList TCHECKDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList LEDF1DoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList LEDF2DoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList Shock1DoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList ShockDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList ShortFlagDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList TDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList SecondDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList AMaxDoubleCircuitTWYCenterDrive = new ArrayList();
        ArrayList ErrorCodeDoubleCircuitTWYCenterDrive = new ArrayList();
        #endregion

        #region 单路滑中驱动各项参数存储集合
        ArrayList RMS1SingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList RMS2SingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList RMS1LASTSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList RMS2LASTSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList Val2SingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList TCHECKSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList LEDF1SingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList LEDF2SingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList Shock1SingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList ShockSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList ShortFlagSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList TSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList SecondSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList AMaxSingleCircuitTWYCenterDrive = new ArrayList();
        ArrayList ErrorCodeSingleCircuitTWYCenterDrive = new ArrayList();
        #endregion

        #region 立式跑道灯具驱动各项参数存储集合
        ArrayList RMS1ElevatedRWYLightDrive = new ArrayList();
        ArrayList RMS1LASTElevatedRWYLightDrive = new ArrayList();
        ArrayList Val2ElevatedRWYLightDrive = new ArrayList();
        ArrayList TCHECKElevatedRWYLightDrive = new ArrayList();
        ArrayList LEDF1ElevatedRWYLightDrive = new ArrayList();
        ArrayList LEDVSNS1ElevatedRWYLightDrive = new ArrayList();
        ArrayList LEDVSNS2ElevatedRWYLightDrive = new ArrayList();
        ArrayList TempertureElevatedRWYLightDrive = new ArrayList();
        ArrayList HumidityElevatedRWYLightDrive = new ArrayList();
        ArrayList AMaxElevatedRWYLightDrive = new ArrayList();
        ArrayList ShortFlagElevatedRWYLightDrive = new ArrayList();
        ArrayList TElevatedRWYLightDrive = new ArrayList();
        ArrayList SecondElevatedRWYLightDrive = new ArrayList();
        ArrayList ErrorCodeElevatedRWYLightDrive = new ArrayList();
        #endregion
        #endregion

        #region 中英文切换字符
        #region 后台代码，串口设置页面，中英文切换字符串
        string LampInchesLabel1 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel1");
        string LampInchesLabel2 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel2");
        string LampInchesLabel3 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel3");
        string LampInchesLabel4 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel4");
        string LampInchesLabel5 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel5");
        string LampInchesLabel6 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel6");
        string LampInchesLabel7 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel7");
        string LampInchesLabel8 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel8");
        string LampInchesLabel9 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel9");
        string LampInchesLabel10 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel10");



        #endregion

        #region 后台代码，工厂模式页面，中英文切换字符串
        string AnswerHardwareVersion0 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion0");
        string AnswerHardwareVersion1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion1");
        string AnswerHardwareVersion2 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion2");
        string AnswerHardwareVersion3 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion3");
        string AnswerHardwareVersion4 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion4");
        string AnswerHardwareVersion5 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion5");
        string AnswerHardwareVersion6 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion6");
        string AnswerHardwareVersion7 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion7");
        string AnswerHardwareVersion8 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion8");


        string AnswerLampModel0 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel0");
        string AnswerLampModel1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel1");
        string AnswerLampModel2 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel2");
        string AnswerLampModel3 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel3");
        string AnswerLampModel4 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel4");
        string AnswerLampModel5 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel5");
        string AnswerLampModel6 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel6");
        string AnswerLampModel7 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel7");
        string AnswerLampModel8 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel8");
        string AnswerLampModel9 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel9");
        string AnswerLampModel10 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel10");
        string AnswerLampModel11 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel11");
        string AnswerLampModel12 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel12");
        string AnswerLampModel13 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel13");
        string AnswerLampModel14 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel14");
        string AnswerLampModel15 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel15");
        string AnswerLampModel16 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel16");
        string AnswerLampModel17 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel17");
        string AnswerLampModel18 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel18");
        string AnswerLampModel19 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel19");
        string AnswerLampModel20 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel20");
        string AnswerLampModel21 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel21");
        string AnswerLampModel22 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel22");
        string AnswerLampModel23 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel23");
        string AnswerLampModel24 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel24");
        string AnswerLampModel25 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel25");
        string AnswerLampModel26 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel26");
        string AnswerLampModel27 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel27");
        string AnswerLampModel28 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel28");
        string AnswerLampModel29 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel29");
        string AnswerLampModel30 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel30");
        string AnswerLampModel31 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel31");
        string AnswerLampModel32 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel32");
        string AnswerLampModel33 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel33");
        string AnswerLampModel34 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel34");
        string AnswerLampModel35 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel35");
        string AnswerLampModel36 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel36");
        string AnswerLampModel37 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel37");
        string AnswerLampModel38 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel38");
        string AnswerLampModel39 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel39");
        string AnswerLampModel40 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel40");
        string AnswerLampModel41 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel41");
        string AnswerLampModel42 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel42");
        string AnswerLampModel43 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel43");
        string AnswerLampModel44 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel44");
        string AnswerLampModel45 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel45");
        string AnswerLampModel46 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel46");
        string AnswerLampModel47 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel47");
        string AnswerLampModel48 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel48");
        string AnswerLampModel49 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel49");
        string AnswerLampModel50 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel50");
        string AnswerLampModel51 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel51");
        string AnswerLampModel52 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel52");
        string AnswerLampModel53 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel53");
        string AnswerLampModel54 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel54");
        string AnswerLampModel55 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel55");
        string AnswerLampModel56 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel56");
        string AnswerLampModel57 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel57");
        string AnswerLampModel58 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel58");
        string AnswerLampModel59 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel59");
        string AnswerLampModel60 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel60");
        string AnswerLampModel61 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel61");
        string AnswerLampModel62 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel62");
        string AnswerLampModel63 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel63");
        string AnswerLampModel64 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel64");
        string AnswerLampModel65 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel65");
        string AnswerLampModel66 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel66");
        string AnswerLampModel67 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel67");
        string AnswerLampModel69 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel69");
        string AnswerLampModel70 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel70");
        string AnswerLampModel71 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel71");
        string AnswerLampModel68 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel68");
        string AnswerLampModel72 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel72");
        string AnswerLampModel73 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel73");
        string AnswerLampModel74 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel74");
        string AnswerLampModel75 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel75");
        string AnswerLampModel76 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel76");
        string AnswerLampModel77 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel77");
        string AnswerLampModel78 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel78");
        string AnswerLampModel79 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel79");
        string AnswerLampModel80 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel80");
        string AnswerLampModel81 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel81");
        string AnswerLampModel82 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel82");
        string AnswerLampModel83 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel83");
        string AnswerLampModel84 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel84");

        string AnswerOpenCircuit1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerOpenCircuit1");
        string AnswerOpenCircuit2 = (string)System.Windows.Application.Current.FindResource("LangsAnswerOpenCircuit2");
        #endregion

        #region 后台代码，开发者模式页面，中英文切换字符串
        string AnswerStatus1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerStatus1");
        string CreateExcel1 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel1");
        string CreateExcel2 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel2");
        string CreateExcel3 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel3");
        string CreateExcel4 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel4");
        string CreateExcel5 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel5");
        string CreateExcel6 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel6");
        string CreateExcel7 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel7");
        string CreateExcel8 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel8");
        string CreateExcel9 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel9");
        string CreateExcel10 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel10");
        string CreateExcel11 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel11");


        string CreateTxt1 = (string)System.Windows.Application.Current.FindResource("LangsCreateTxt1");

        #endregion

        #region 后台代码，Messagebox，中英文切换字符串
        string MessageboxHeader1 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxHeader1");
        string MessageboxHeader2 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxHeader2");

        string MessageboxContent1 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent1");
        string MessageboxContent2 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent2");
        string MessageboxContent3 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent3");
        string MessageboxContent4 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent4");
        string MessageboxContent5 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent5");
        string MessageboxContent6 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent6");
        string MessageboxContent7 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent7");
        string MessageboxContent8 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent8");
        string MessageboxContent9 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent9");
        string MessageboxContent10 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent10");
        string MessageboxContent11 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent11");
        string MessageboxContent12 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent12");
        string MessageboxContent13 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent13");
        string MessageboxContent14 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent14");
        string MessageboxContent15 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent15");
        string MessageboxContent16 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent16");
        string MessageboxContent17 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent17");
        string MessageboxContent18 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent18");
        string MessageboxContent19 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent19");
        string MessageboxContent20 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent20");
        string MessageboxContent21 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent21");
        string MessageboxContent22 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent22");
        string MessageboxContent23 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent23");
        string MessageboxContent24 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent24");
        string MessageboxContent25 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent25");
        string MessageboxContent26 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent26");
        string MessageboxContent27 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent27");
        string MessageboxContent28 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent28");
        string MessageboxContent29 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent29");
        string MessageboxContent30 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent30");
        string MessageboxContent31 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent31");
        string MessageboxContent32 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent32");
        string MessageboxContent33 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent33");
        string MessageboxContent34 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent34");
        string MessageboxContent35 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent35");
        string MessageboxContent36 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent36");
        string MessageboxContent37 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent37");
        string MessageboxContent38 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent38");
        string MessageboxContent39 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent39");
        string MessageboxContent40 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent40");
        string MessageboxContent41 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent41");
        string MessageboxContent42 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent42");
        string MessageboxContent43 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent43");
        string MessageboxContent44 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent44");
        string MessageboxContent45 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent45");
        string MessageboxContent46 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent46");

        #endregion
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

            for(int i=0;i<FlashFrequencyArray.Length;i++)
            {
                FlashFrequencyArray[i] = 30 + i;
            }
            FlashFrequencySelect.ItemsSource = FlashFrequencyArray;
            FlashFrequencySelect.SelectedIndex = 0;

            InDeveloperModeFlashFrequencySelect.ItemsSource = FlashFrequencyArray;
            InDeveloperModeFlashFrequencySelect.SelectedIndex = 0;

            queryVersionCommand[27] = CalculateCheckOutValue(queryVersionCommand);  //计算版本查询指令的校验值
            queryStatusCommand[27] = CalculateCheckOutValue(queryStatusCommand);  //计算状态查询指令的校验值    
            InFactoryModeCommonLightRestoreOriginalCommand[27] = CalculateCheckOutValue(InFactoryModeCommonLightRestoreOriginalCommand);
            InFactoryModeRWYGuardLightRestoreOriginalCommand[27] = CalculateCheckOutValue(InFactoryModeRWYGuardLightRestoreOriginalCommand);
            queryTotalTimeCommand[27] = CalculateCheckOutValue(queryTotalTimeCommand);

            SettingSerialPort.IsSelected = true;
            
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            RefreshStringMessageLanguage();
            MessageBoxResult result = MessageBox.Show(MessageboxContent1, MessageboxHeader2, MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result==MessageBoxResult.Yes)
            {
                ConfigurationWindow.IsEnabled = true;

                if (lampsPort.IsOpen==true)
                {
                    if(MessageBox.Show(MessageboxContent2, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.None)
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
            RefreshStringMessageLanguage();
            try
            {
                if (lampsPort.IsOpen == false)
                {
                    lampsPort.PortName = PortNameSelect.SelectedItem.ToString();
                    lampsPort.Open();

                    PurgingDeveloperMode();
                    PurgingFactoryMode();
                }
                if(lampsPort.IsOpen==true)
                {                   
                    judgeFeedbackCommand = 4;
                    LampInchesLabel.Content = "";
                    lampsPort.Write(queryVersionCommand, 0, 28);

                    Thread.Sleep(1000);
                    if (judgeFeedbackCommand == 4)
                    {                        
                        lampsPort.Close();                           
                        
                        if(lampsPort.IsOpen == false)
                        {
                            if (MessageBox.Show(MessageboxContent3, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }

                            NoneLampSelect();
                            LampInchesLabel.Content = LampInchesLabel1;
                        }
                                                                         
                    }
                    else
                    {
                        if (MessageBox.Show(MessageboxContent4, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                if( MessageBox.Show(MessageboxContent5, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
            RefreshStringMessageLanguage();
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
                    if( MessageBox.Show(MessageboxContent6, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }

                    LampInchesLabel.Content = "";
                }
            }
            catch
            {
                if( MessageBox.Show(MessageboxContent7, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            RefreshStringMessageLanguage();
            if (lampsPort.IsOpen)
            {
                judgeFeedbackCommand = 2;                
                lampsPort.Write(queryVersionCommand, 0, 28);
                Thread.Sleep(1000);
                if(judgeFeedbackCommand==2)
                {
                    if( MessageBox.Show(MessageboxContent8, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                if( MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            RefreshStringMessageLanguage();
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
                    if( MessageBox.Show(MessageboxContent10, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                if( MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 总时间查询
        private void QueryTotalTime_Click(object sender, RoutedEventArgs e)
        {
            RefreshStringMessageLanguage();
            if (lampsPort.IsOpen)
            {
                judgeFeedbackCommand = 5;
                lampsPort.Write(queryTotalTimeCommand, 0, 28);
                Thread.Sleep(1000);
                if (judgeFeedbackCommand == 5)
                {
                    if (MessageBox.Show(MessageboxContent44, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                if (MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            RefreshStringMessageLanguage();
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
                        case 5: QueryTotalTimeFeedbackCommand();break;
                    }
                }
                
            }
            else
            {                
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region NoFeedbackCommand
        private void NoFeedbackCommand()
        {
            if (dataReceived.Length == 1 && dataReceived[0] == 0x00)
            {
                return;
            }
            else
            {
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent11, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region SetParameterFeedbackCommand
        private void SetParameterFeedbackCommand()
        {
            judgeFeedbackCommand = 0;

            if (dataReceived.Length == 4)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if ((dataReceived[0] == 0x02 && dataReceived[1] == 0x55 && dataReceived[2] == 0x11) || (dataReceived[0] == 0x55 && dataReceived[1] == 0x02 && dataReceived[2] == 0x11) || (dataReceived[0] == 0x03 && dataReceived[1] == 0x66 && dataReceived[2] == 0x11))
                    {
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show(MessageboxContent12, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                            if (MessageBox.Show(MessageboxContent13, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                        if (MessageBox.Show(MessageboxContent14, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent15, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
        #endregion

        #region QueryVersionFeedbackCommand
        private void QueryVersionFeedbackCommand()
        {
            judgeFeedbackCommand = 0;
            if (dataReceived.Length == 24)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x22 && dataReceived[3] == 0x85)
                    {
                        year = dataReceived[4];
                        month = dataReceived[5];
                        date = dataReceived[6];
                        softwareVersion1 = dataReceived[7];
                        softwareVersion2 = dataReceived[8];
                        softwareVersion3 = dataReceived[9];
                        breakFlag = dataReceived[14];
                        lampsNumber = dataReceived[15];
                        hardwareVersion1 = dataReceived[16];
                        hardwareVersion2 = dataReceived[17];
                        hardwareVersion3 = dataReceived[18];
                        softwareNumber = dataReceived[19];
                        channelSelect = dataReceived[20];
                        flashFrequency = dataReceived[21];
                        waveformSelect = dataReceived[22];

                        currentRatio1 = CalculateRealCurrentValue(dataReceived[10]);
                        currentRatio2 = CalculateRealCurrentValue(dataReceived[11]);
                        currentRatio3 = CalculateRealCurrentValue(dataReceived[12]);
                        currentRatio4 = CalculateRealCurrentValue(dataReceived[13]);

                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            PurgeAnswerVersionTextblock();
                            RefreshStringMessageLanguage();

                            AnswerSoftwareVersion.Text = "V" + softwareVersion1.ToString() + "." + softwareVersion2.ToString() + "." + softwareVersion3.ToString() + " " + " 20" + year.ToString() + "/" + month.ToString() + "/" + date.ToString();
                            AnswerLampModel.Text = LampsContentShow(lampsNumber);
                            AnswerIA.Text = currentRatio1.ToString();
                            AnswerIB.Text = currentRatio2.ToString();
                            AnswerIIA.Text = currentRatio3.ToString();
                            AnswerIIB.Text = currentRatio4.ToString();
                            if (breakFlag == 0)
                            {
                                AnswerOpenCircuit.Text = AnswerOpenCircuit1;
                            }
                            else
                            {
                                AnswerOpenCircuit.Text = AnswerOpenCircuit2;
                            }

                            string driveName = "";
                            if (hardwareVersion1 == 12 && softwareNumber == 0)
                            {
                                driveName = AnswerHardwareVersion0;
                            }
                            else if (hardwareVersion1 == 8 && softwareNumber == 0)
                            {
                                driveName = AnswerHardwareVersion1;
                            }
                            else if (hardwareVersion1 == 13 && softwareNumber == 2)
                            {
                                driveName = AnswerHardwareVersion2;
                            }
                            else if (hardwareVersion1 == 12 && softwareNumber == 3)
                            {
                                driveName = AnswerHardwareVersion3;
                            }
                            else if (hardwareVersion1 == 8 && softwareNumber == 4)
                            {
                                driveName = AnswerHardwareVersion4;
                            }
                            else if (hardwareVersion1 == 5 && softwareNumber == 6)
                            {
                                driveName = AnswerHardwareVersion6;
                            }
                            else if (hardwareVersion1 == 4 && softwareNumber == 5)
                            {
                                driveName = AnswerHardwareVersion7;
                            }
                            else if (hardwareVersion1 == 9 && softwareNumber == 9)
                            {
                                driveName = AnswerHardwareVersion8;
                            }
                            else
                            {
                                driveName = AnswerHardwareVersion5;
                            }

                            AnswerHardwareVersion.Text = "V" + hardwareVersion2.ToString() + "." + hardwareVersion3.ToString() + "  " + driveName;

                            if (hardwareVersion1 == 8 && softwareNumber == 4)
                            {
                                RWYGuardLightVersionStatus.Visibility = Visibility.Visible;
                                AnswerChannelSelect.Text = channelSelect.ToString();
                                AnswerFlashFrequency.Text = flashFrequency.ToString();
                                AnswerWaveformSelect.Text = waveformSelect.ToString();
                            }
                            else
                            {
                                RWYGuardLightVersionStatus.Visibility = Visibility.Collapsed;
                            }

                        }));

                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show(MessageboxContent16, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                            if (MessageBox.Show(MessageboxContent13, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                        if (MessageBox.Show(MessageboxContent14, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent15, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        public string LampsContentShow(byte lampNumber)
        {
            string result = "";

            RefreshStringMessageLanguage();
            switch (lampNumber)
            {
                case 0: result = AnswerLampModel0; break;
                case 1: result = AnswerLampModel1; break;
                case 2: result = AnswerLampModel2; break;
                case 3: result = AnswerLampModel3; break;
                case 4: result = AnswerLampModel4; break;
                case 5: result = AnswerLampModel5; break;
                case 6: result = AnswerLampModel6; break;
                case 7: result = AnswerLampModel7; break;
                case 8: result = AnswerLampModel8; break;
                case 9: result = AnswerLampModel9; break;
                case 10: result = AnswerLampModel10; break;
                case 11: result = AnswerLampModel11; break;
                case 12: result = AnswerLampModel12; break;
                case 13: result = AnswerLampModel13; break;
                case 14: result = AnswerLampModel14; break;
                case 15: result = AnswerLampModel15; break;
                case 16: result = AnswerLampModel16; break;
                case 17: result = AnswerLampModel17; break;
                case 18: result = AnswerLampModel18; break;
                case 19: result = AnswerLampModel19; break;
                case 20: result = AnswerLampModel20; break;
                case 21: result = AnswerLampModel21; break;
                case 22: result = AnswerLampModel22; break;
                case 23: result = AnswerLampModel23; break;
                case 24: result = AnswerLampModel24; break;
                case 25: result = AnswerLampModel25; break;
                case 26: result = AnswerLampModel26; break;
                case 27: result = AnswerLampModel27; break;
                case 28: result = AnswerLampModel28; break;
                case 29: result = AnswerLampModel29; break;
                case 30: result = AnswerLampModel30; break;
                case 31: result = AnswerLampModel31; break;
                case 32: result = AnswerLampModel32; break;
                case 33: result = AnswerLampModel33; break;
                case 34: result = AnswerLampModel34; break;
                case 35: result = AnswerLampModel35; break;
                case 36: result = AnswerLampModel36; break;
                case 37: result = AnswerLampModel37; break;
                case 38: result = AnswerLampModel38; break;
                case 39: result = AnswerLampModel39; break;
                case 40: result = AnswerLampModel40; break;
                case 41: result = AnswerLampModel41; break;
                case 42: result = AnswerLampModel42; break;
                case 43: result = AnswerLampModel43; break;
                case 44: result = AnswerLampModel44; break;
                case 45: result = AnswerLampModel45; break;
                case 46: result = AnswerLampModel46; break;
                case 47: result = AnswerLampModel47; break;
                case 48: result = AnswerLampModel48; break;
                case 49: result = AnswerLampModel49; break;
                case 50: result = AnswerLampModel50; break;
                case 51: result = AnswerLampModel51; break;
                case 52: result = AnswerLampModel52; break;
                case 53: result = AnswerLampModel53; break;
                case 54: result = AnswerLampModel54; break;
                case 55: result = AnswerLampModel55; break;
                case 56: result = AnswerLampModel56; break;
                case 57: result = AnswerLampModel57; break;
                case 58: result = AnswerLampModel58; break;
                case 59: result = AnswerLampModel59; break;
                case 60: result = AnswerLampModel60; break;
                case 61: result = AnswerLampModel61; break;
                case 62: result = AnswerLampModel62; break;
                case 63: result = AnswerLampModel63; break;
                case 64: result = AnswerLampModel64; break;
                case 65: result = AnswerLampModel65; break;
                case 66: result = AnswerLampModel66; break;
                case 67: result = AnswerLampModel67; break;
                case 69: result = AnswerLampModel69; break;
                case 70: result = AnswerLampModel70; break;
                case 71: result = AnswerLampModel71; break;
                case 68: result = AnswerLampModel68; break;
                case 72: result = AnswerLampModel72; break;
                case 73: result = AnswerLampModel73; break;
                case 74: result = AnswerLampModel74; break;
                case 75: result = AnswerLampModel75; break;
                case 76: result = AnswerLampModel76; break;
                case 77: result = AnswerLampModel77; break;
                case 78: result = AnswerLampModel78; break;
                case 79: result = AnswerLampModel79; break;
                case 80: result = AnswerLampModel80; break;
                case 81: result = AnswerLampModel81; break;
                case 82: result = AnswerLampModel82; break;
                case 83: result = AnswerLampModel83; break;
                case 84: result = AnswerLampModel84; break;


            }
            return result;
        }

        private double CalculateRealCurrentValue(byte originalData)
        {
            double original = originalData / 10.0;
            double result = 0.0;  
           

            for (int i = 0; i < totalTimeObject.Length; i++)
            {
                if (hardwareVersion1 == totalTimeObject[i][0] && softwareNumber == totalTimeObject[i][1] && softwareVersion1 == totalTimeObject[i][2] && softwareVersion2 == totalTimeObject[i][3])
                {
                    if(totalTimeObject[i][5]==1)
                    {
                        result= original * totalTimeObject[i][6];
                    }
                    else if(totalTimeObject[i][5] == 2)
                    {
                        result = original / totalTimeObject[i][6];
                    }
                    
                }
            }



            if ((hardwareVersion1 == 12 && softwareNumber == 0&& softwareVersion1==1&& softwareVersion2==3)|| (hardwareVersion1 == 12 && softwareNumber == 0 && softwareVersion1 == 1 && softwareVersion2 == 2))
            {                

                if (lampsNumber == 1 || lampsNumber == 2 || lampsNumber == 3)
                {
                    result = result * 1.3;
                }
                if ((lampsNumber >= 33 && lampsNumber <= 40) || lampsNumber == 47)
                {
                    if (result == 0.9)
                    {
                        result = 0.93 * 1.3;
                    }
                    if (result == 0.7)
                    {
                        result = 0.7 * 1;
                    }
                    if (result == 0.6)
                    {
                        result = 0.55 * 1;
                    }
                    if (result == 0.4)
                    {
                        result = 0.45 * 1;
                    }
                }
            }

            if (hardwareVersion1 == 12 && softwareNumber == 3 && softwareVersion1 == 1 && softwareVersion2 == 0)
            {
                if(result==0.4)
                {
                    result = 0.45;
                }

            }

            //for temporary
            if (hardwareVersion1 == 4 && softwareNumber == 5 && softwareVersion1 == 1 && softwareVersion2 == 3)
            {
                result *= 0.75;
            }


            result = Math.Round(result, 2, MidpointRounding.AwayFromZero);

            return result;

        }

        private void PurgeAnswerVersionTextblock()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                AnswerSoftwareVersion.Text = "";
                AnswerHardwareVersion.Text = "";
                AnswerLampModel.Text = "";
                AnswerIA.Text = "";
                AnswerIB.Text = "";
                AnswerIIA.Text = "";
                AnswerIIB.Text = "";
                AnswerOpenCircuit.Text = "";
            }));
        }
        #endregion

        #region QueryStatusFeedbackCommand
        ArrayList ReceivedStatusFeedbackCommand = new ArrayList();  //定义接收到的状态反馈指令        
        private void QueryStatusFeedbackCommand()
        {
            ReceivedStatusFeedbackCommand.AddRange(dataReceived);

            RefreshStringMessageLanguage();
            QueryStatusTimeSpan = DateTime.Now - StartQueryStatus;
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                AnswerStatus.Text = AnswerStatus1 + " " + QueryStatusTimeSpan.Hours.ToString().PadLeft(2, '0') + ":" + QueryStatusTimeSpan.Minutes.ToString().PadLeft(2, '0') + ":" + QueryStatusTimeSpan.Seconds.ToString().PadLeft(2, '0') + ":" + QueryStatusTimeSpan.Milliseconds.ToString().PadLeft(3, '0');
            }));

            if (dataReceived.Length == 4)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x11)
                    {
                        QueryStatusNoContentFeedbackCommand();
                    }
                }
            }

        }

        private void QueryStatusNoContentFeedbackCommand()
        {
            judgeFeedbackCommand = 0;

            if (dataReceived.Length == 4)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x11)
                    {
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show(MessageboxContent21, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                            if (MessageBox.Show(MessageboxContent13, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                        if (MessageBox.Show(MessageboxContent14, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent15, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #endregion

        #region ConfirmLampInches
        private void ConfirmLampInches()
        {
            judgeFeedbackCommand = 0;
            RefreshStringMessageLanguage();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                LampInchesLabel.Content = "";

                if (dataReceived.Length == 24)
                {
                    byte checkOutValue = CalculateCheckOutValue(dataReceived);
                    if (checkOutValue == dataReceived[dataReceived.Length - 1])
                    {
                        if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x22 && dataReceived[3] == 0x85)
                        {
                            softwareVersion1 = dataReceived[7];
                            softwareVersion2 = dataReceived[8];                            
                            hardwareVersion1 = dataReceived[16];
                            softwareNumber = dataReceived[19];

                            ConfirmQueryTotalTimeIsUseOrNot(hardwareVersion1, softwareNumber, softwareVersion1, softwareVersion2);

                            if (hardwareVersion1 == 12 && softwareNumber == 0)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel3;
                                TwelveInchesLampSelect();
                            }
                            else if (hardwareVersion1 == 8 && softwareNumber == 0)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel4;
                                EightInchesLampSelect();
                            }
                            else if (hardwareVersion1 == 13 && softwareNumber == 2)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel5;
                                DoubleCircuitRWYCenterLampSelect();
                            }
                            else if (hardwareVersion1 == 12 && softwareNumber == 3)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel6;
                                DoubleCircuitRWYCenter12inchesLampSelect();
                            }
                            else if (hardwareVersion1 == 8 && softwareNumber == 4)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel7;
                                RWYGuardLampSelect();
                            }
                            else if (hardwareVersion1 == 5 && softwareNumber == 6)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel8;
                                DoubleCircuitTWYCenterLampSelect();
                            }
                            else if (hardwareVersion1 == 4 && softwareNumber == 5)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel9;
                                SingleCircuitTWYCenterLampSelect();
                            }
                            else if (hardwareVersion1 == 9 && softwareNumber == 9)
                            {
                                LampInchesLabel.Content = LampInchesLabel2 + " " + LampInchesLabel10;
                                ElevatedRWYLampSelect();
                            }
                            else
                            {
                                LampInchesLabel.Content = LampInchesLabel1;
                                NoneLampSelect();
                            }

                            
                            if (hardwareVersion1 == 8 && softwareNumber == 4)
                            {
                                this.Dispatcher.Invoke(new System.Action(() =>
                                {
                                    InDeveloperModeRWYGuardLightParametersSetting.Visibility = Visibility.Visible;
                                    RestoreOriginalStatus.IsEnabled = false;
                                }));
                            }
                            else if(hardwareVersion1 == 9 && softwareNumber == 9)
                            {
                                this.Dispatcher.Invoke(new System.Action(() =>
                                {
                                    InDeveloperModeRWYGuardLightParametersSetting.Visibility = Visibility.Collapsed;
                                    RestoreOriginalStatus.IsEnabled = false;
                                }));
                            }
                            else
                            {
                                this.Dispatcher.Invoke(new System.Action(() =>
                                {
                                    InDeveloperModeRWYGuardLightParametersSetting.Visibility = Visibility.Collapsed;
                                    RestoreOriginalStatus.IsEnabled = true;
                                }));
                            }

                        }
                        else
                        {
                            if (MessageBox.Show(MessageboxContent3, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                            {
                                ConfigurationWindow.IsEnabled = true;
                            }
                            else
                            {
                                ConfigurationWindow.IsEnabled = false;
                            }

                            NoneLampSelect();
                            LampInchesLabel.Content = LampInchesLabel1;
                        }
                    }
                    else
                    {
                        if (MessageBox.Show(MessageboxContent3, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                        {
                            ConfigurationWindow.IsEnabled = true;
                        }
                        else
                        {
                            ConfigurationWindow.IsEnabled = false;
                        }

                        NoneLampSelect();
                        LampInchesLabel.Content = LampInchesLabel1;
                    }
                }
                else
                {
                    if (MessageBox.Show(MessageboxContent3, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }

                    NoneLampSelect();
                    LampInchesLabel.Content = LampInchesLabel1;
                }
            }));
        }

        #region 不同驱动可选的灯具类型
        public void EightInchesLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = true;
                Select8inchesRWYEndLight.IsEnabled = true;
                SelectRapidExitTWYIndicatorLight.IsEnabled = true;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = true;

            }));
        }

        public void TwelveInchesLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = true;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = true;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }

        public void DoubleCircuitRWYCenterLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = true;
                Select8inchesRWYEndLight.IsEnabled = true;
                SelectRapidExitTWYIndicatorLight.IsEnabled = true;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }

        public void DoubleCircuitRWYCenter12inchesLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                SelectApproachChenterlineLight.IsEnabled = false;
                SelectApproachCrossbarLight.IsEnabled = false;
                SelectApproachSideRowLight.IsEnabled = false;
                SelectRWYThresholdWingBarLight.IsEnabled = false;
                SelectRWYThresholdLight.IsEnabled = false;
                SelectRWYEdgeLight.IsEnabled = false;
                Select12inchesRWYEndLight.IsEnabled = false;
                SelectRWYThresholdEndLight.IsEnabled = true;
                SelectRWYCenterlineLight.IsEnabled = false;
                Select12inchesRWYCenterlineLight.IsEnabled = true;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }

        public void RWYGuardLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = true;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }

        public void NoneLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }

        public void DoubleCircuitTWYCenterLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = true;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;



            }));
        }

        public void SingleCircuitTWYCenterLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = true;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = true;
                SelectIntermediateHoldingPositionLight.IsEnabled = true;
                SelectTWYIntersectionsLight.IsEnabled = true;
                SelectTWYEdgeLight.IsEnabled = true;
                SelectElevatedApproachCenterlineLight.IsEnabled = false;
                SelectElevatedApproachCrossbarLight.IsEnabled = false;
                SelectElevatedApproachSideRowLight.IsEnabled = false;
                SelectElevatedRWYEdgeLight.IsEnabled = false;
                SelectElevatedRWYEndLight.IsEnabled = false;
                SelectElevatedRWYThresholdLight.IsEnabled = false;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = false;
                SelectElevatedTWYStopBarLight.IsEnabled = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }

        public void ElevatedRWYLampSelect()
        {
            this.Dispatcher.Invoke(new System.Action(() =>
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
                Select12inchesRWYCenterlineLight.IsEnabled = false;
                SelectRWYTouchdownZoneLight.IsEnabled = false;
                Select8inchesRWYEndLight.IsEnabled = false;
                SelectRapidExitTWYIndicatorLight.IsEnabled = false;
                SelectCombinedRWYEdgeLight.IsEnabled = false;
                SelectRWYGuardLight.IsEnabled = false;
                SelectTWYCenterLight.IsEnabled = false;
                SelectTWYCenterLight2P.IsEnabled = false;
                SelectTWYStopBarLight.IsEnabled = false;
                SelectIntermediateHoldingPositionLight.IsEnabled = false;
                SelectTWYIntersectionsLight.IsEnabled = false;
                SelectTWYEdgeLight.IsEnabled = false;
                SelectElevatedApproachCenterlineLight.IsEnabled = true;
                SelectElevatedApproachCrossbarLight.IsEnabled = true;
                SelectElevatedApproachSideRowLight.IsEnabled = true;
                SelectElevatedRWYEdgeLight.IsEnabled = true;
                SelectElevatedRWYEndLight.IsEnabled = true;
                SelectElevatedRWYThresholdLight.IsEnabled = true;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = true;
                SelectElevatedTWYStopBarLight.IsEnabled = true;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = false;


            }));
        }
        #endregion

        public void ConfirmQueryTotalTimeIsUseOrNot(int hardware,int s,int software1,int software2)
        {
            double result = 0;

            for(int i=0;i<totalTimeObject.Length;i++)
            {
                if(hardware==totalTimeObject[i][0]&&s==totalTimeObject[i][1]&&software1==totalTimeObject[i][2]&&software2==totalTimeObject[i][3])
                {
                    result = totalTimeObject[i][4];
                }
            }

            if(result==1)
            {
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    QueryTotalTime.IsEnabled = true;
                }));
            }
            else
            {
                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    QueryTotalTime.IsEnabled = false;
                }));
            }
        }

        #endregion

        #region QueryTotalTimeFeedbackCommand
        private void QueryTotalTimeFeedbackCommand()
        {
            judgeFeedbackCommand = 0;

            if (dataReceived.Length == 16)
            {
                byte checkOutValue = CalculateCheckOutValue(dataReceived);
                if (checkOutValue == dataReceived[dataReceived.Length - 1])
                {
                    if (dataReceived[0] == 0x02 && dataReceived[1] == 0x89 && dataReceived[2] == 0x22 && dataReceived[3] == 0x85)
                    {                        
                        uint breakDownCount = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            uint SecondOrigin = dataReceived[4 + j];
                            breakDownCount |= SecondOrigin;
                            if (j < 3)
                            {
                                breakDownCount <<= 8;
                            }
                        }

                        uint totalTime = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            uint SecondOrigin = dataReceived[8 + j];
                            totalTime |= SecondOrigin;
                            if (j < 3)
                            {
                                totalTime <<= 8;
                            }
                        }

                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            TotalTimeEnquiryWindows myTotalTimeEnquiryWindows = new TotalTimeEnquiryWindows();
                            myTotalTimeEnquiryWindows.breakDownCount = breakDownCount;
                            myTotalTimeEnquiryWindows.totalTime = totalTime;
                            myTotalTimeEnquiryWindows.ShowDialog();                            
                        }));
                        return;
                    }
                    else
                    {
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show(MessageboxContent13, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                        if (MessageBox.Show(MessageboxContent14, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent15, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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


        #endregion

  
        #endregion

        #region 导出最终数据
        byte[] receivedStatusFeedbackCommand;
        Thread CreateExcelThread;
        private void CreatExcel_Click(object sender, RoutedEventArgs e)
        {
            RefreshStringMessageLanguage();
            CreateExcelThread = new Thread(new ThreadStart(CreateExcelThreadProcess));
            CreateExcelThread.Start();            
        }

        private void CreateExcelThreadProcess()
        {
            if (judgeFeedbackCommand == 3)
            {
                if (ReceivedStatusFeedbackCommand.Count != 0)
                {
                    ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                    {
                        ShowEXCELHandleProcess.Visibility = Visibility.Visible;
                    }));

                    receivedStatusFeedbackCommand = new byte[ReceivedStatusFeedbackCommand.Count];

                    for (int i = 0; i < receivedStatusFeedbackCommand.Length; i++)
                    {
                        receivedStatusFeedbackCommand[i] = (byte)ReceivedStatusFeedbackCommand[i];
                    }
                    
                    if (hardwareVersion1==12&&softwareNumber==0)
                    {
                        TwelveInchesLampDataAnalysis(receivedStatusFeedbackCommand);
                        TwelveInchesLampParametersCreatExcel();
                    }
                    else if (hardwareVersion1 == 8 && softwareNumber == 0)
                    {
                        if(receivedStatusFeedbackCommand[2]==0x01)
                        {
                            EightInchesLampDataAnalysis(receivedStatusFeedbackCommand);
                            EightInchesLampParametersCreatExcel();
                        }
                        else if(receivedStatusFeedbackCommand[2] == 0x02)
                        {
                            EightInchesV2LampDataAnalysis(receivedStatusFeedbackCommand);
                            EightInchesV2LampParametersCreatExcel();
                        }                       
                    }
                    else if (hardwareVersion1 == 13 && softwareNumber == 2)
                    {
                        DoubleCircuitRWYCenterDriveLampDataAnalysis(receivedStatusFeedbackCommand);
                        DoubleCircuitRWYCenterDriveLampParametersCreatExcel();
                    }                   
                    else if (hardwareVersion1 == 12 && softwareNumber == 3)
                    {
                        DoubleCircuitRWYCenterDrive12inchesLampDataAnalysis(receivedStatusFeedbackCommand);
                        DoubleCircuitRWYCenterDrive12inchesLampParametersCreatExcel();
                    }
                    else if (hardwareVersion1 == 8 && softwareNumber == 4)
                    {
                        RWYGuardLampDataAnalysis(receivedStatusFeedbackCommand);
                        RWYGuardLampParametersCreatExcel();
                    }
                    else if (hardwareVersion1 == 5 && softwareNumber == 6)
                    {
                        DoubleCircuitTWYCenterDriveLampDataAnalysis(receivedStatusFeedbackCommand);
                        DoubleCircuitTWYCenterDriveParametersCreatExcel();
                    }
                    else if (hardwareVersion1 == 4 && softwareNumber == 5)
                    {
                        SingleCircuitTWYCenterDriveLampDataAnalysis(receivedStatusFeedbackCommand);
                        SingleCircuitTWYCenterDriveParametersCreatExcel();
                    }
                    else if (hardwareVersion1 == 9 && softwareNumber == 9)
                    {
                        ElevatedRWYLightDriveLampDataAnalysis(receivedStatusFeedbackCommand);
                        ElevatedRWYLightDriveParametersCreatExcel();
                    }
                    else
                    {                       
                        this.Dispatcher.Invoke(new System.Action(() =>
                        {
                            if (MessageBox.Show(MessageboxContent13, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                        if (MessageBox.Show(MessageboxContent26, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent27, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x0C && CompleteData[i + 4] == 0x0C)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1Twelveinches.Add(DataArray[i][5] * 500);
                        RMS2Twelveinches.Add(DataArray[i][6] * 500);
                        Val2Twelveinches.Add(DataArray[i][7] * 20);
                        Val3Twelveinches.Add(DataArray[i][8]);
                        RMSMID1Twelveinches.Add(DataArray[i][9] * 16);
                        RMSMID2Twelveinches.Add(DataArray[i][10] * 16);
                        RMSTwelveinches.Add(DataArray[i][11] * 4);
                        CurrentRatio1Twelveinches.Add((float)(DataArray[i][12] / 10.0));
                        CurrentRatio2Twelveinches.Add((float)(DataArray[i][13] / 10.0));
                        CurrentRatio3Twelveinches.Add((float)(DataArray[i][14] / 10.0));
                        CurrentRatio4Twelveinches.Add((float)(DataArray[i][15] / 10.0));
                        RESIATwelveinches.Add(DataArray[i][16] * 124);
                        RESIBTwelveinches.Add(DataArray[i][17] * 124);
                        RESIIATwelveinches.Add(DataArray[i][18] * 124);
                        RESIIBTwelveinches.Add(DataArray[i][19] * 124);
                        SNSIATwelveinches.Add(DataArray[i][20] * 16);
                        SNSIBTwelveinches.Add(DataArray[i][21] * 16);
                        SNSIIATwelveinches.Add(DataArray[i][22] * 16);
                        SNSIIBTwelveinches.Add(DataArray[i][23] * 16);
                        LEDF1Twelveinches.Add(DataArray[i][24]);
                        TTwelveinches.Add((SByte)DataArray[i][25]);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][26 + j];
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
            ErrorCodeTwelveinches.Clear();
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
                str_fileName = "d:\\" + CreateExcel3 + " " + CreateExcel1 + " "+DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel3 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
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
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x08 && CompleteData[i + 4] == 0x08)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1Eightinches.Add(DataArray[i][5] * 1100);
                        Val2Eightinches.Add(DataArray[i][6] * 20);
                        Val3Eightinches.Add(DataArray[i][7]);
                        RMSEightinches.Add(DataArray[i][8] * 4);
                        CurrentRatio1Eightinches.Add((float)(DataArray[i][9] / 10.0));
                        CurrentRatio3Eightinches.Add((float)(DataArray[i][10] / 10.0));
                        RESIAEightinches.Add(DataArray[i][11] * 124);
                        RESIIAEightinches.Add(DataArray[i][12] * 124);
                        SNSIAEightinches.Add(DataArray[i][13] * 16);
                        SNSIIAEightinches.Add(DataArray[i][14] * 16);
                        LEDF1Eightinches.Add(DataArray[i][15]);
                        TEightinches.Add((SByte)DataArray[i][16]);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][17 + j];
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
            ErrorCodeEightinches.Clear();
        }

        void EightInchesLampParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\" +CreateExcel11+" " + CreateExcel1 +" "+ DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel11 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
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
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 8寸V2灯具状态信息解析
        private void EightInchesV2LampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x02 && CompleteData[i + 3] == 0x08 && CompleteData[i + 4] == 0x08)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            
            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1EightinchesV2.Add(DataArray[i][5] * 1100);
                        Val2EightinchesV2.Add(DataArray[i][6] * 20);
                        TCheckEightinchesV2.Add(DataArray[i][7]*16);
                        RMSEightinchesV2.Add(DataArray[i][8] * 4);
                        CurrentRatio1EightinchesV2.Add((float)(DataArray[i][9] / 10.0));
                        CurrentRatio3EightinchesV2.Add((float)(DataArray[i][10] / 10.0));
                        RESIAEightinchesV2.Add(DataArray[i][11] * 124);
                        RESIIAEightinchesV2.Add(DataArray[i][12] * 124);
                        SNSIAEightinchesV2.Add(DataArray[i][13] * 16);
                        SNSIIAEightinchesV2.Add(DataArray[i][14] * 16);
                        LEDF1EightinchesV2.Add(DataArray[i][15]);
                        TEightinchesV2.Add((SByte)DataArray[i][16]);
                        Shock1EightinchesV2.Add(DataArray[i][22]);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][17 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondEightinchesV2.Add(SecondResult);

                        ErrorCodeEightinchesV2.Add("No Error");
                    }
                    else
                    {
                        EightInchesV2LampCheckValueErrorHandle();
                    }
                }
                else
                {
                    EightInchesV2LampCommandLengthErrorHandle();
                }
            }
        }

        private void EightInchesV2LampCheckValueErrorHandle()
        {
            RMS1EightinchesV2.Add("Null");
            Val2EightinchesV2.Add("Null");
            TCheckEightinchesV2.Add("Null");
            RMSEightinchesV2.Add("Null");
            CurrentRatio1EightinchesV2.Add("Null");
            CurrentRatio3EightinchesV2.Add("Null");
            RESIAEightinchesV2.Add("Null");
            RESIIAEightinchesV2.Add("Null");
            SNSIAEightinchesV2.Add("Null");
            SNSIIAEightinchesV2.Add("Null");
            LEDF1EightinchesV2.Add("Null");
            TEightinchesV2.Add("Null");
            SecondEightinchesV2.Add("Null");
            Shock1EightinchesV2.Add("Null");
            ErrorCodeEightinchesV2.Add("Check Value Error");
        }

        private void EightInchesV2LampCommandLengthErrorHandle()
        {
            RMS1EightinchesV2.Add("Null");
            Val2EightinchesV2.Add("Null");
            TCheckEightinchesV2.Add("Null");
            RMSEightinchesV2.Add("Null");
            CurrentRatio1EightinchesV2.Add("Null");
            CurrentRatio3EightinchesV2.Add("Null");
            RESIAEightinchesV2.Add("Null");
            RESIIAEightinchesV2.Add("Null");
            SNSIAEightinchesV2.Add("Null");
            SNSIIAEightinchesV2.Add("Null");
            LEDF1EightinchesV2.Add("Null");
            TEightinchesV2.Add("Null");
            SecondEightinchesV2.Add("Null");
            Shock1EightinchesV2.Add("Null");
            ErrorCodeEightinchesV2.Add("Command Length Error");
        }

        private void ClearEightInchesV2LampsParameter()
        {
            RMS1EightinchesV2.Clear();
            Val2EightinchesV2.Clear();
            TCheckEightinchesV2.Clear();
            RMSEightinchesV2.Clear();
            CurrentRatio1EightinchesV2.Clear();
            CurrentRatio3EightinchesV2.Clear();
            RESIAEightinchesV2.Clear();
            RESIIAEightinchesV2.Clear();
            SNSIAEightinchesV2.Clear();
            SNSIIAEightinchesV2.Clear();
            LEDF1EightinchesV2.Clear();
            TEightinchesV2.Clear();
            SecondEightinchesV2.Clear();
            Shock1EightinchesV2.Clear();
            ErrorCodeEightinchesV2.Clear();
        }

        void EightInchesV2LampParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\" + CreateExcel10 + " " + CreateExcel1 + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel10 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "Val2";
                ExcelSheet.Cells[2, 4] = "T_CHECK";
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
                ExcelSheet.Cells[2, 15] = "shock1";
                ExcelSheet.Cells[2, 16] = "Error Code";

                //输出各个参数值
                for (int i = 0; i < RMS1EightinchesV2.Count; i++)
                {
                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1EightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = Val2EightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = TCheckEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = RMSEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = CurrentRatio1EightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = CurrentRatio3EightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = RESIAEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = RESIIAEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = SNSIAEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = SNSIIAEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = LEDF1EightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = TEightinchesV2[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = Shock1EightinchesV2[i].ToString();

                    if (SecondEightinchesV2[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 14] = SecondEightinchesV2[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 14] = ((int)SecondEightinchesV2[i] / 3600).ToString() + ":" + (((int)SecondEightinchesV2[i] % 3600) / 60).ToString() + ":" + (((int)SecondEightinchesV2[i] % 3600) % 60).ToString();
                    }
                    ExcelSheet.Cells[3 + i, 16] = ErrorCodeEightinchesV2[i].ToString();
                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearEightInchesV2LampsParameter();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 双路跑中驱动灯具状态信息解析
        private void DoubleCircuitRWYCenterDriveLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x0D && CompleteData[i + 4] == 0x0D)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1DoubleCircuitRWYCenterDrive.Add(DataArray[i][5] * 1100);
                        Val2DoubleCircuitRWYCenterDrive.Add(DataArray[i][6] * 20);
                        RMS2DoubleCircuitRWYCenterDrive.Add(DataArray[i][7] * 1100);
                        CurrentRatio1DoubleCircuitRWYCenterDrive.Add((float)(DataArray[i][8] / 10.0));
                        CurrentRatio2DoubleCircuitRWYCenterDrive.Add((float)(DataArray[i][9] / 10.0));
                        CurrentRatio3DoubleCircuitRWYCenterDrive.Add((float)(DataArray[i][10] / 10.0));
                        CurrentRatio4DoubleCircuitRWYCenterDrive.Add((float)(DataArray[i][11] / 10.0));
                        RMS1LASTDoubleCircuitRWYCenterDrive.Add(DataArray[i][12] * 4);
                        RMS2LASTDoubleCircuitRWYCenterDrive.Add(DataArray[i][13] * 4);
                        SNSIADoubleCircuitRWYCenterDrive.Add(DataArray[i][14] * 16);
                        SNSIBDoubleCircuitRWYCenterDrive.Add(DataArray[i][15] * 16);
                        SNSIIADoubleCircuitRWYCenterDrive.Add(DataArray[i][16] * 16);
                        SNSIIBDoubleCircuitRWYCenterDrive.Add(DataArray[i][17] * 16);
                        LEDF1DoubleCircuitRWYCenterDrive.Add(DataArray[i][18]);
                        LEDF2DoubleCircuitRWYCenterDrive.Add(DataArray[i][19]);
                        RESIADoubleCircuitRWYCenterDrive.Add(DataArray[i][20] * 124);
                        RESIBDoubleCircuitRWYCenterDrive.Add(DataArray[i][21] * 124);
                        RESIIADoubleCircuitRWYCenterDrive.Add(DataArray[i][22] * 124);
                        RESIIBDoubleCircuitRWYCenterDrive.Add(DataArray[i][23] * 124);
                        TDoubleCircuitRWYCenterDrive.Add((SByte)DataArray[i][24]);
                        TCHECKDoubleCircuitRWYCenterDrive.Add(DataArray[i][29]);                                               

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][26 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondDoubleCircuitRWYCenterDrive.Add(SecondResult);
                        ErrorCodeDoubleCircuitRWYCenterDrive.Add("No Error");
                    }
                    else
                    {
                        DoubleCircuitRWYCenterDriveLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    DoubleCircuitRWYCenterDriveLampCommandLengthErrorHandle();
                }
            }

        }

        private void DoubleCircuitRWYCenterDriveLampCheckValueErrorHandle()
        {
            RMS1DoubleCircuitRWYCenterDrive.Add("Null");
            Val2DoubleCircuitRWYCenterDrive.Add("Null");
            RMS2DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio1DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio2DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio3DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio4DoubleCircuitRWYCenterDrive.Add("Null");
            RMS1LASTDoubleCircuitRWYCenterDrive.Add("Null");
            RMS2LASTDoubleCircuitRWYCenterDrive.Add("Null");
            SNSIADoubleCircuitRWYCenterDrive.Add("Null");
            SNSIBDoubleCircuitRWYCenterDrive.Add("Null");
            SNSIIADoubleCircuitRWYCenterDrive.Add("Null");
            SNSIIBDoubleCircuitRWYCenterDrive.Add("Null");
            LEDF1DoubleCircuitRWYCenterDrive.Add("Null");
            LEDF2DoubleCircuitRWYCenterDrive.Add("Null");
            RESIADoubleCircuitRWYCenterDrive.Add("Null");
            RESIBDoubleCircuitRWYCenterDrive.Add("Null");
            RESIIADoubleCircuitRWYCenterDrive.Add("Null");
            RESIIBDoubleCircuitRWYCenterDrive.Add("Null");
            TDoubleCircuitRWYCenterDrive.Add("Null");
            SecondDoubleCircuitRWYCenterDrive.Add("Null");
            TCHECKDoubleCircuitRWYCenterDrive.Add("Null");
            ErrorCodeDoubleCircuitRWYCenterDrive.Add("Check Value Error");          
        }

        private void DoubleCircuitRWYCenterDriveLampCommandLengthErrorHandle()
        {
            RMS1DoubleCircuitRWYCenterDrive.Add("Null");
            Val2DoubleCircuitRWYCenterDrive.Add("Null");
            RMS2DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio1DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio2DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio3DoubleCircuitRWYCenterDrive.Add("Null");
            CurrentRatio4DoubleCircuitRWYCenterDrive.Add("Null");
            RMS1LASTDoubleCircuitRWYCenterDrive.Add("Null");
            RMS2LASTDoubleCircuitRWYCenterDrive.Add("Null");
            SNSIADoubleCircuitRWYCenterDrive.Add("Null");
            SNSIBDoubleCircuitRWYCenterDrive.Add("Null");
            SNSIIADoubleCircuitRWYCenterDrive.Add("Null");
            SNSIIBDoubleCircuitRWYCenterDrive.Add("Null");
            LEDF1DoubleCircuitRWYCenterDrive.Add("Null");
            LEDF2DoubleCircuitRWYCenterDrive.Add("Null");
            RESIADoubleCircuitRWYCenterDrive.Add("Null");
            RESIBDoubleCircuitRWYCenterDrive.Add("Null");
            RESIIADoubleCircuitRWYCenterDrive.Add("Null");
            RESIIBDoubleCircuitRWYCenterDrive.Add("Null");
            TDoubleCircuitRWYCenterDrive.Add("Null");
            SecondDoubleCircuitRWYCenterDrive.Add("Null");
            TCHECKDoubleCircuitRWYCenterDrive.Add("Null");
            ErrorCodeDoubleCircuitRWYCenterDrive.Add("Command Length Error");           
        }

        private void ClearDoubleCircuitRWYCenterDriveLampsParameters()
        {
            RMS1DoubleCircuitRWYCenterDrive.Clear();
            Val2DoubleCircuitRWYCenterDrive.Clear();
            RMS2DoubleCircuitRWYCenterDrive.Clear();
            CurrentRatio1DoubleCircuitRWYCenterDrive.Clear();
            CurrentRatio2DoubleCircuitRWYCenterDrive.Clear();
            CurrentRatio3DoubleCircuitRWYCenterDrive.Clear();
            CurrentRatio4DoubleCircuitRWYCenterDrive.Clear();
            RMS1LASTDoubleCircuitRWYCenterDrive.Clear();
            RMS2LASTDoubleCircuitRWYCenterDrive.Clear();
            SNSIADoubleCircuitRWYCenterDrive.Clear();
            SNSIBDoubleCircuitRWYCenterDrive.Clear();
            SNSIIADoubleCircuitRWYCenterDrive.Clear();
            SNSIIBDoubleCircuitRWYCenterDrive.Clear();
            LEDF1DoubleCircuitRWYCenterDrive.Clear();
            LEDF2DoubleCircuitRWYCenterDrive.Clear();
            RESIADoubleCircuitRWYCenterDrive.Clear();
            RESIBDoubleCircuitRWYCenterDrive.Clear();
            RESIIADoubleCircuitRWYCenterDrive.Clear();
            RESIIBDoubleCircuitRWYCenterDrive.Clear();
            TDoubleCircuitRWYCenterDrive.Clear();
            SecondDoubleCircuitRWYCenterDrive.Clear();
            TCHECKDoubleCircuitRWYCenterDrive.Clear();
            ErrorCodeDoubleCircuitRWYCenterDrive.Clear();           
        }
        
        void DoubleCircuitRWYCenterDriveLampParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\ " + CreateExcel4 +" "+CreateExcel1+ " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel4 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "Val2"; 
                ExcelSheet.Cells[2, 4] = "RMS2";
                ExcelSheet.Cells[2, 5] = "Current_Ratio1";
                ExcelSheet.Cells[2, 6] = "Current_Ratio2";
                ExcelSheet.Cells[2, 7] = "Current_Ratio3";
                ExcelSheet.Cells[2, 8] = "Current_Ratio4";
                ExcelSheet.Cells[2, 9] = "RMS1_Last";
                ExcelSheet.Cells[2, 10] = "RMS2_Last";
                ExcelSheet.Cells[2, 11] = "SNS_IA";
                ExcelSheet.Cells[2, 12] = "SNS_IB";
                ExcelSheet.Cells[2, 13] = "SNS_IIA"; 
                ExcelSheet.Cells[2, 14] = "SNS_IIB"; 
                ExcelSheet.Cells[2, 15] = "LED_F1"; 
                ExcelSheet.Cells[2, 16] = "LED_F2"; 
                ExcelSheet.Cells[2, 17] = "RES_IA";
                ExcelSheet.Cells[2, 18] = "RES_IB";
                ExcelSheet.Cells[2, 19] = "RES_IIA";
                ExcelSheet.Cells[2, 20] = "RES_IIB";
                ExcelSheet.Cells[2, 21] = "T";
                ExcelSheet.Cells[2, 22] = "Second";
                ExcelSheet.Cells[2, 23] = "T_Check";
                ExcelSheet.Cells[2, 24] = "Error Code";                

                //输出各个参数值
                for (int i = 0; i < RMS1DoubleCircuitRWYCenterDrive.Count; i++)
                {
                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = Val2DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = RMS2DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = CurrentRatio1DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = CurrentRatio2DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = CurrentRatio3DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = CurrentRatio4DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = RMS1LASTDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = RMS2LASTDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = SNSIADoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = SNSIBDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = SNSIIADoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 14] = SNSIIBDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = LEDF1DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 16] = LEDF2DoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 17] = RESIADoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 18] = RESIBDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 19] = RESIIADoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 20] = RESIIBDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 21] = TDoubleCircuitRWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 23] = TCHECKDoubleCircuitRWYCenterDrive[i].ToString();
                    if (SecondDoubleCircuitRWYCenterDrive[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 22] = SecondDoubleCircuitRWYCenterDrive[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 22] = ((int)SecondDoubleCircuitRWYCenterDrive[i] / 3600).ToString() + ":" + (((int)SecondDoubleCircuitRWYCenterDrive[i] % 3600) / 60).ToString() + ":" + (((int)SecondDoubleCircuitRWYCenterDrive[i] % 3600) % 60).ToString();
                    }
                    ExcelSheet.Cells[3 + i, 24] = ErrorCodeDoubleCircuitRWYCenterDrive[i].ToString();
                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearDoubleCircuitRWYCenterDriveLampsParameters(); 

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 12寸双路跑中驱动灯具状态信息解析
        private void DoubleCircuitRWYCenterDrive12inchesLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x0C && CompleteData[i + 4] == 0x0D)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][5] * 500);
                        RMS2DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][6] * 500);
                        Val2DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][7] * 20);
                        Val3DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][8]);
                        RMSMID1DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][9] * 16);
                        RMSMID2DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][10] * 16);
                        RMS1LASTDoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][11] * 4);
                        CurrentRatio1DoubleCircuitRWYCenterDrive12inches.Add((float)(DataArray[i][12] / 10.0));
                        CurrentRatio2DoubleCircuitRWYCenterDrive12inches.Add((float)(DataArray[i][13] / 10.0));
                        CurrentRatio3DoubleCircuitRWYCenterDrive12inches.Add((float)(DataArray[i][14] / 10.0));
                        CurrentRatio4DoubleCircuitRWYCenterDrive12inches.Add((float)(DataArray[i][15] / 10.0));
                        RESIADoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][16] * 124);
                        RESIBDoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][17] * 124);
                        RESIIADoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][18] * 124);
                        RESIIBDoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][19] * 124);
                        SNSIADoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][20] * 16);
                        SNSIBDoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][21] * 16);
                        SNSIIADoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][22] * 16);
                        LEDF2DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][23]);
                        LEDF1DoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][24]);
                        TDoubleCircuitRWYCenterDrive12inches.Add((SByte)DataArray[i][25]);
                        RMS2LASTDoubleCircuitRWYCenterDrive12inches.Add(DataArray[i][30] * 4);                                             

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][26 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondDoubleCircuitRWYCenterDrive12inches.Add(SecondResult);
                        ErrorCodeDoubleCircuitRWYCenterDrive12inches.Add("No Error");
                    }
                    else
                    {
                        DoubleCircuitRWYCenterDrive12inchesLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    DoubleCircuitRWYCenterDrive12inchesLampCommandLengthErrorHandle();
                }
            }

        }

        private void DoubleCircuitRWYCenterDrive12inchesLampCheckValueErrorHandle()
        {
            RMS1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMS2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            Val2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            Val3DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMSMID1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMSMID2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMS1LASTDoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio3DoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio4DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIBDoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIIBDoubleCircuitRWYCenterDrive12inches.Add("Null");
            SNSIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            SNSIBDoubleCircuitRWYCenterDrive12inches.Add("Null");
            SNSIIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            LEDF2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            LEDF1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            TDoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMS2LASTDoubleCircuitRWYCenterDrive12inches.Add("Null");
            SecondDoubleCircuitRWYCenterDrive12inches.Add("Null");
            ErrorCodeDoubleCircuitRWYCenterDrive12inches.Add("Check Value Error");            
        }

        private void DoubleCircuitRWYCenterDrive12inchesLampCommandLengthErrorHandle()
        {
            RMS1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMS2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            Val2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            Val3DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMSMID1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMSMID2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMS1LASTDoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio3DoubleCircuitRWYCenterDrive12inches.Add("Null");
            CurrentRatio4DoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIBDoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            RESIIBDoubleCircuitRWYCenterDrive12inches.Add("Null");
            SNSIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            SNSIBDoubleCircuitRWYCenterDrive12inches.Add("Null");
            SNSIIADoubleCircuitRWYCenterDrive12inches.Add("Null");
            LEDF2DoubleCircuitRWYCenterDrive12inches.Add("Null");
            LEDF1DoubleCircuitRWYCenterDrive12inches.Add("Null");
            TDoubleCircuitRWYCenterDrive12inches.Add("Null");
            RMS2LASTDoubleCircuitRWYCenterDrive12inches.Add("Null");
            SecondDoubleCircuitRWYCenterDrive12inches.Add("Null");
            ErrorCodeDoubleCircuitRWYCenterDrive12inches.Add("Command Length Error");
        }

        private void ClearDoubleCircuitRWYCenterDrive12inchesLampsParameters()
        {
            RMS1DoubleCircuitRWYCenterDrive12inches.Clear();
            RMS2DoubleCircuitRWYCenterDrive12inches.Clear();
            Val2DoubleCircuitRWYCenterDrive12inches.Clear();
            Val3DoubleCircuitRWYCenterDrive12inches.Clear();
            RMSMID1DoubleCircuitRWYCenterDrive12inches.Clear();
            RMSMID2DoubleCircuitRWYCenterDrive12inches.Clear();
            RMS1LASTDoubleCircuitRWYCenterDrive12inches.Clear();
            CurrentRatio1DoubleCircuitRWYCenterDrive12inches.Clear();
            CurrentRatio2DoubleCircuitRWYCenterDrive12inches.Clear();
            CurrentRatio3DoubleCircuitRWYCenterDrive12inches.Clear();
            CurrentRatio4DoubleCircuitRWYCenterDrive12inches.Clear();
            RESIADoubleCircuitRWYCenterDrive12inches.Clear();
            RESIBDoubleCircuitRWYCenterDrive12inches.Clear();
            RESIIADoubleCircuitRWYCenterDrive12inches.Clear();
            RESIIBDoubleCircuitRWYCenterDrive12inches.Clear();
            SNSIADoubleCircuitRWYCenterDrive12inches.Clear();
            SNSIBDoubleCircuitRWYCenterDrive12inches.Clear();
            SNSIIADoubleCircuitRWYCenterDrive12inches.Clear();
            LEDF2DoubleCircuitRWYCenterDrive12inches.Clear();
            LEDF1DoubleCircuitRWYCenterDrive12inches.Clear();
            TDoubleCircuitRWYCenterDrive12inches.Clear();
            RMS2LASTDoubleCircuitRWYCenterDrive12inches.Clear();
            SecondDoubleCircuitRWYCenterDrive12inches.Clear();
            ErrorCodeDoubleCircuitRWYCenterDrive12inches.Clear();           
        }

        void DoubleCircuitRWYCenterDrive12inchesLampParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\ " + CreateExcel5 + " " + CreateExcel1 + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel5 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "RMS2"; 
                ExcelSheet.Cells[2, 4] = "Val2";
                ExcelSheet.Cells[2, 5] = "Val3";
                ExcelSheet.Cells[2, 6] = "RMSMID1";
                ExcelSheet.Cells[2, 7] = "RMSMID2";
                ExcelSheet.Cells[2, 8] = "RMS1_Last"; 
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
                ExcelSheet.Cells[2, 20] = "LED_F2";
                ExcelSheet.Cells[2, 21] = "LED_F1"; 
                ExcelSheet.Cells[2, 22] = "T"; 
                ExcelSheet.Cells[2, 23] = "Second";
                ExcelSheet.Cells[2, 24] = "RMS2_Last"; 
                ExcelSheet.Cells[2,25]= "Error Code";

                //输出各个参数值
                for (int i = 0; i < RMS1DoubleCircuitRWYCenterDrive12inches.Count; i++)
                {

                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = RMS2DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = Val2DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = Val3DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = RMSMID1DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = RMSMID2DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = RMS1LASTDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = CurrentRatio1DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = CurrentRatio2DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = CurrentRatio3DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = CurrentRatio4DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = RESIADoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 14] = RESIBDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = RESIIADoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 16] = RESIIBDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 17] = SNSIADoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 18] = SNSIBDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 19] = SNSIIADoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 20] = LEDF2DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 21] = LEDF1DoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 22] = TDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    if (SecondDoubleCircuitRWYCenterDrive12inches[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 23] = SecondDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 23] = ((int)SecondDoubleCircuitRWYCenterDrive12inches[i] / 3600).ToString() + ":" + (((int)SecondDoubleCircuitRWYCenterDrive12inches[i] % 3600) / 60).ToString() + ":" + (((int)SecondDoubleCircuitRWYCenterDrive12inches[i] % 3600) % 60).ToString();
                    }
                    ExcelSheet.Cells[3 + i, 24] = RMS2LASTDoubleCircuitRWYCenterDrive12inches[i].ToString();
                    ExcelSheet.Cells[3 + i, 25] = ErrorCodeDoubleCircuitRWYCenterDrive12inches[i].ToString();
                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearDoubleCircuitRWYCenterDrive12inchesLampsParameters();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 8寸警戒灯灯具状态信息解析
        private void RWYGuardLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x08 && CompleteData[i + 4] == 0x04)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1RWYGuardLight.Add(DataArray[i][5] * 1100);
                        Val2RWYGuardLight.Add(DataArray[i][6] * 20);
                        Val3RWYGuardLight.Add(DataArray[i][7]);
                        RMSRWYGuardLight.Add(DataArray[i][8] *4);
                        CurrentRatio1RWYGuardLight.Add((float)(DataArray[i][9] / 10.0));
                        CurrentRatio3RWYGuardLight.Add((float)(DataArray[i][10] / 10.0));
                        WaveformRWYGuardLight.Add(DataArray[i][11]);
                        ChannelRWYGuardLight.Add(DataArray[i][12]);
                        SNSIARWYGuardLight.Add(DataArray[i][13] * 16);
                        SNSIIARWYGuardLight.Add(DataArray[i][14] *16);
                        LEDF1RWYGuardLight.Add(DataArray[i][15]);
                        TRWYGuardLight.Add((SByte)DataArray[i][16]);                        
                        FlashFrequencyRWYGuardLight.Add(DataArray[i][21]);
                        ModeRWYGuardLight.Add(DataArray[i][22]);                                               

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][17 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondRWYGuardLight.Add(SecondResult);
                        ErrorCodeRWYGuardLight.Add("No Error");                        
                    }
                    else
                    {
                        RWYGuardLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    RWYGuardLampCommandLengthErrorHandle();
                }
            }

        }

        private void RWYGuardLampCheckValueErrorHandle()
        {
            RMS1RWYGuardLight.Add("Null");
            Val2RWYGuardLight.Add("Null");
            Val3RWYGuardLight.Add("Null");
            RMSRWYGuardLight.Add("Null");
            CurrentRatio1RWYGuardLight.Add("Null");
            CurrentRatio3RWYGuardLight.Add("Null");
            WaveformRWYGuardLight.Add("Null");
            ChannelRWYGuardLight.Add("Null");
            SNSIARWYGuardLight.Add("Null");
            SNSIIARWYGuardLight.Add("Null");
            LEDF1RWYGuardLight.Add("Null");
            TRWYGuardLight.Add("Null");
            FlashFrequencyRWYGuardLight.Add("Null");
            ModeRWYGuardLight.Add("Null");
            SecondRWYGuardLight.Add("Null");
            ErrorCodeRWYGuardLight.Add("Check Value Error");           
        }

        private void RWYGuardLampCommandLengthErrorHandle()
        {
            RMS1RWYGuardLight.Add("Null");
            Val2RWYGuardLight.Add("Null");
            Val3RWYGuardLight.Add("Null");
            RMSRWYGuardLight.Add("Null");
            CurrentRatio1RWYGuardLight.Add("Null");
            CurrentRatio3RWYGuardLight.Add("Null");
            WaveformRWYGuardLight.Add("Null");
            ChannelRWYGuardLight.Add("Null");
            SNSIARWYGuardLight.Add("Null");
            SNSIIARWYGuardLight.Add("Null");
            LEDF1RWYGuardLight.Add("Null");
            TRWYGuardLight.Add("Null");
            FlashFrequencyRWYGuardLight.Add("Null");
            ModeRWYGuardLight.Add("Null");
            SecondRWYGuardLight.Add("Null");
            ErrorCodeRWYGuardLight.Add("Command Length Error");            
        }

        private void ClearRWYGuardLampsParameters()
        {
            RMS1RWYGuardLight.Clear();
            Val2RWYGuardLight.Clear();
            Val3RWYGuardLight.Clear();
            RMSRWYGuardLight.Clear();
            CurrentRatio1RWYGuardLight.Clear();
            CurrentRatio3RWYGuardLight.Clear();
            WaveformRWYGuardLight.Clear();
            ChannelRWYGuardLight.Clear();
            SNSIARWYGuardLight.Clear();
            SNSIIARWYGuardLight.Clear();
            LEDF1RWYGuardLight.Clear();
            TRWYGuardLight.Clear();
            FlashFrequencyRWYGuardLight.Clear();
            ModeRWYGuardLight.Clear();
            SecondRWYGuardLight.Clear();
            ErrorCodeRWYGuardLight.Clear();           
        }

        void RWYGuardLampParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\ " + CreateExcel6 + " " + CreateExcel1 + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel6 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "Val2";
                ExcelSheet.Cells[2, 4] = "Val3";
                ExcelSheet.Cells[2, 5] = "RMS";
                ExcelSheet.Cells[2, 6] = "Current_Ratio1";
                ExcelSheet.Cells[2, 7] = "Current_Ratio3";
                ExcelSheet.Cells[2, 8] = "Waveform";
                ExcelSheet.Cells[2, 9] = "Channel";
                ExcelSheet.Cells[2, 10] = "SNS_IA";
                ExcelSheet.Cells[2, 11] = "SNS_IIA";
                ExcelSheet.Cells[2, 12] = "LED_F1";
                ExcelSheet.Cells[2, 13] = "T";
                ExcelSheet.Cells[2, 14] = "Second";
                ExcelSheet.Cells[2, 15] = "Flash Frequency";
                ExcelSheet.Cells[2, 16] = "Mode";               
                ExcelSheet.Cells[2, 17] = "Error Code";

                //输出各个参数值
                for (int i = 0; i < RMS1RWYGuardLight.Count; i++)
                {                    
                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1RWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = Val2RWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = Val3RWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = RMSRWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = CurrentRatio1RWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = CurrentRatio3RWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = WaveformRWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = ChannelRWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = SNSIARWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = SNSIIARWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = LEDF1RWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = TRWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = FlashFrequencyRWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 16] = ModeRWYGuardLight[i].ToString();
                    ExcelSheet.Cells[3 + i, 17] = ErrorCodeRWYGuardLight[i].ToString();
                   
                    if (SecondRWYGuardLight[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 14] = SecondRWYGuardLight[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 14] = ((int)SecondRWYGuardLight[i] / 3600).ToString() + ":" + (((int)SecondRWYGuardLight[i] % 3600) / 60).ToString() + ":" + (((int)SecondRWYGuardLight[i] % 3600) % 60).ToString();
                    }
                    
                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearRWYGuardLampsParameters();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 双路滑中驱动灯具状态信息解析
        private void DoubleCircuitTWYCenterDriveLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x05 && CompleteData[i + 4] == 0x06)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {                        
                        RMS1DoubleCircuitTWYCenterDrive.Add(DataArray[i][5] * 500);
                        RMS2DoubleCircuitTWYCenterDrive.Add(DataArray[i][6] * 500);
                        RMS1LASTDoubleCircuitTWYCenterDrive.Add(DataArray[i][7] * 4);
                        RMS2LASTDoubleCircuitTWYCenterDrive.Add(DataArray[i][8] * 4);
                        Val2DoubleCircuitTWYCenterDrive.Add(DataArray[i][9] *20);
                        TCHECKDoubleCircuitTWYCenterDrive.Add(DataArray[i][10] *16);
                        LEDF1DoubleCircuitTWYCenterDrive.Add(DataArray[i][11]);
                        LEDF2DoubleCircuitTWYCenterDrive.Add(DataArray[i][12]);
                        Shock1DoubleCircuitTWYCenterDrive.Add(DataArray[i][13]);
                        ShockDoubleCircuitTWYCenterDrive.Add(DataArray[i][14]);
                        ShortFlagDoubleCircuitTWYCenterDrive.Add(DataArray[i][15]);
                        TDoubleCircuitTWYCenterDrive.Add((SByte)DataArray[i][16]);
                        AMaxDoubleCircuitTWYCenterDrive.Add(DataArray[i][21]*16);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][17 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondDoubleCircuitTWYCenterDrive.Add(SecondResult);
                        ErrorCodeDoubleCircuitTWYCenterDrive.Add("No Error");
                    }
                    else
                    {
                        DoubleCircuitTWYCenterDriveLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    DoubleCircuitTWYCenterDriveLampCommandLengthErrorHandle();
                }
            }

        }

        private void DoubleCircuitTWYCenterDriveLampCheckValueErrorHandle()
        {
            RMS1DoubleCircuitTWYCenterDrive.Add("Null");
            RMS2DoubleCircuitTWYCenterDrive.Add("Null");
            RMS1LASTDoubleCircuitTWYCenterDrive.Add("Null");
            RMS2LASTDoubleCircuitTWYCenterDrive.Add("Null");
            Val2DoubleCircuitTWYCenterDrive.Add("Null");
            TCHECKDoubleCircuitTWYCenterDrive.Add("Null");
            LEDF1DoubleCircuitTWYCenterDrive.Add("Null");
            LEDF2DoubleCircuitTWYCenterDrive.Add("Null");
            Shock1DoubleCircuitTWYCenterDrive.Add("Null");
            ShockDoubleCircuitTWYCenterDrive.Add("Null");
            ShortFlagDoubleCircuitTWYCenterDrive.Add("Null");
            TDoubleCircuitTWYCenterDrive.Add("Null");
            AMaxDoubleCircuitTWYCenterDrive.Add("Null");
            SecondDoubleCircuitTWYCenterDrive.Add("Null");
            ErrorCodeDoubleCircuitTWYCenterDrive.Add("Check Value Error");  
        }

        private void DoubleCircuitTWYCenterDriveLampCommandLengthErrorHandle()
        {
            RMS1DoubleCircuitTWYCenterDrive.Add("Null");
            RMS2DoubleCircuitTWYCenterDrive.Add("Null");
            RMS1LASTDoubleCircuitTWYCenterDrive.Add("Null");
            RMS2LASTDoubleCircuitTWYCenterDrive.Add("Null");
            Val2DoubleCircuitTWYCenterDrive.Add("Null");
            TCHECKDoubleCircuitTWYCenterDrive.Add("Null");
            LEDF1DoubleCircuitTWYCenterDrive.Add("Null");
            LEDF2DoubleCircuitTWYCenterDrive.Add("Null");
            Shock1DoubleCircuitTWYCenterDrive.Add("Null");
            ShockDoubleCircuitTWYCenterDrive.Add("Null");
            ShortFlagDoubleCircuitTWYCenterDrive.Add("Null");
            TDoubleCircuitTWYCenterDrive.Add("Null");
            AMaxDoubleCircuitTWYCenterDrive.Add("Null");
            SecondDoubleCircuitTWYCenterDrive.Add("Null");
            ErrorCodeDoubleCircuitTWYCenterDrive.Add("Command Length Error");           
        }

        private void ClearDoubleCircuitTWYCenterDriveParameters()
        {
            RMS1DoubleCircuitTWYCenterDrive.Clear();
            RMS2DoubleCircuitTWYCenterDrive.Clear();
            RMS1LASTDoubleCircuitTWYCenterDrive.Clear();
            RMS2LASTDoubleCircuitTWYCenterDrive.Clear();
            Val2DoubleCircuitTWYCenterDrive.Clear();
            TCHECKDoubleCircuitTWYCenterDrive.Clear();
            LEDF1DoubleCircuitTWYCenterDrive.Clear();
            LEDF2DoubleCircuitTWYCenterDrive.Clear();
            Shock1DoubleCircuitTWYCenterDrive.Clear();
            ShockDoubleCircuitTWYCenterDrive.Clear();
            ShortFlagDoubleCircuitTWYCenterDrive.Clear();
            TDoubleCircuitTWYCenterDrive.Clear();
            AMaxDoubleCircuitTWYCenterDrive.Clear();
            SecondDoubleCircuitTWYCenterDrive.Clear();
            ErrorCodeDoubleCircuitTWYCenterDrive.Clear();           
        }

        void DoubleCircuitTWYCenterDriveParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\ " + CreateExcel7 + " " + CreateExcel1 + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel7 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "RMS2";
                ExcelSheet.Cells[2, 4] = "RMS1_LAST";
                ExcelSheet.Cells[2, 5] = "RMS2_LAST";
                ExcelSheet.Cells[2, 6] = "Val2";
                ExcelSheet.Cells[2, 7] = "T_CHECK";
                ExcelSheet.Cells[2, 8] = "LED_F1";
                ExcelSheet.Cells[2, 9] = "LED_F2";
                ExcelSheet.Cells[2, 10] = "Shock1";
                ExcelSheet.Cells[2, 11] = "Shock";
                ExcelSheet.Cells[2, 12] = "Short_Flag";
                ExcelSheet.Cells[2, 13] = "T";
                ExcelSheet.Cells[2, 14] = "Second";
                ExcelSheet.Cells[2, 15] = "A_MAX";
                ExcelSheet.Cells[2, 16] = "Error Code";

                //输出各个参数值
                for (int i = 0; i < RMS1DoubleCircuitTWYCenterDrive.Count; i++)
                {
                   
                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1DoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = RMS2DoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = RMS1LASTDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = RMS2LASTDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = Val2DoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = TCHECKDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = LEDF1DoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = LEDF2DoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = Shock1DoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = ShockDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = ShortFlagDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = TDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = AMaxDoubleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 16] = ErrorCodeDoubleCircuitTWYCenterDrive[i].ToString();

                    if (SecondDoubleCircuitTWYCenterDrive[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 14] = SecondDoubleCircuitTWYCenterDrive[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 14] = ((int)SecondDoubleCircuitTWYCenterDrive[i] / 3600).ToString() + ":" + (((int)SecondDoubleCircuitTWYCenterDrive[i] % 3600) / 60).ToString() + ":" + (((int)SecondDoubleCircuitTWYCenterDrive[i] % 3600) % 60).ToString();
                    }

                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearDoubleCircuitTWYCenterDriveParameters();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 单路滑中驱动灯具状态信息解析
        private void SingleCircuitTWYCenterDriveLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x04 && CompleteData[i + 4] == 0x05)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1SingleCircuitTWYCenterDrive.Add(DataArray[i][5] * 500);
                        RMS2SingleCircuitTWYCenterDrive.Add(DataArray[i][6] * 500);
                        RMS1LASTSingleCircuitTWYCenterDrive.Add(DataArray[i][7] * 4);
                        RMS2LASTSingleCircuitTWYCenterDrive.Add(DataArray[i][8] * 4);
                        Val2SingleCircuitTWYCenterDrive.Add(DataArray[i][9] * 20);
                        TCHECKSingleCircuitTWYCenterDrive.Add(DataArray[i][10] * 16);
                        LEDF1SingleCircuitTWYCenterDrive.Add(DataArray[i][11]);
                        LEDF2SingleCircuitTWYCenterDrive.Add(DataArray[i][12]);
                        Shock1SingleCircuitTWYCenterDrive.Add(DataArray[i][13]);
                        ShockSingleCircuitTWYCenterDrive.Add(DataArray[i][14]);
                        ShortFlagSingleCircuitTWYCenterDrive.Add(DataArray[i][15]);
                        TSingleCircuitTWYCenterDrive.Add((SByte)DataArray[i][16]);
                        AMaxSingleCircuitTWYCenterDrive.Add(DataArray[i][21] * 16);

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][17 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondSingleCircuitTWYCenterDrive.Add(SecondResult);
                        ErrorCodeSingleCircuitTWYCenterDrive.Add("No Error");
                    }
                    else
                    {
                        SingleCircuitTWYCenterDriveLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    SingleCircuitTWYCenterDriveLampCommandLengthErrorHandle();
                }
            }

        }

        private void SingleCircuitTWYCenterDriveLampCheckValueErrorHandle()
        {
            RMS1SingleCircuitTWYCenterDrive.Add("Null");
            RMS2SingleCircuitTWYCenterDrive.Add("Null");
            RMS1LASTSingleCircuitTWYCenterDrive.Add("Null");
            RMS2LASTSingleCircuitTWYCenterDrive.Add("Null");
            Val2SingleCircuitTWYCenterDrive.Add("Null");
            TCHECKSingleCircuitTWYCenterDrive.Add("Null");
            LEDF1SingleCircuitTWYCenterDrive.Add("Null");
            LEDF2SingleCircuitTWYCenterDrive.Add("Null");
            Shock1SingleCircuitTWYCenterDrive.Add("Null");
            ShockSingleCircuitTWYCenterDrive.Add("Null");
            ShortFlagSingleCircuitTWYCenterDrive.Add("Null");
            TSingleCircuitTWYCenterDrive.Add("Null");
            AMaxSingleCircuitTWYCenterDrive.Add("Null");
            SecondSingleCircuitTWYCenterDrive.Add("Null");
            ErrorCodeSingleCircuitTWYCenterDrive.Add("Check Value Error");
        }

        private void SingleCircuitTWYCenterDriveLampCommandLengthErrorHandle()
        {
            RMS1SingleCircuitTWYCenterDrive.Add("Null");
            RMS2SingleCircuitTWYCenterDrive.Add("Null");
            RMS1LASTSingleCircuitTWYCenterDrive.Add("Null");
            RMS2LASTSingleCircuitTWYCenterDrive.Add("Null");
            Val2SingleCircuitTWYCenterDrive.Add("Null");
            TCHECKSingleCircuitTWYCenterDrive.Add("Null");
            LEDF1SingleCircuitTWYCenterDrive.Add("Null");
            LEDF2SingleCircuitTWYCenterDrive.Add("Null");
            Shock1SingleCircuitTWYCenterDrive.Add("Null");
            ShockSingleCircuitTWYCenterDrive.Add("Null");
            ShortFlagSingleCircuitTWYCenterDrive.Add("Null");
            TSingleCircuitTWYCenterDrive.Add("Null");
            AMaxSingleCircuitTWYCenterDrive.Add("Null");
            SecondSingleCircuitTWYCenterDrive.Add("Null");
            ErrorCodeSingleCircuitTWYCenterDrive.Add("Command Length Error");            
        }

        private void ClearSingleCircuitTWYCenterDriveParameters()
        {
            RMS1SingleCircuitTWYCenterDrive.Clear();
            RMS2SingleCircuitTWYCenterDrive.Clear();
            RMS1LASTSingleCircuitTWYCenterDrive.Clear();
            RMS2LASTSingleCircuitTWYCenterDrive.Clear();
            Val2SingleCircuitTWYCenterDrive.Clear();
            TCHECKSingleCircuitTWYCenterDrive.Clear();
            LEDF1SingleCircuitTWYCenterDrive.Clear();
            LEDF2SingleCircuitTWYCenterDrive.Clear();
            Shock1SingleCircuitTWYCenterDrive.Clear();
            ShockSingleCircuitTWYCenterDrive.Clear();
            ShortFlagSingleCircuitTWYCenterDrive.Clear();
            TSingleCircuitTWYCenterDrive.Clear();
            AMaxSingleCircuitTWYCenterDrive.Clear();
            SecondSingleCircuitTWYCenterDrive.Clear();
            ErrorCodeSingleCircuitTWYCenterDrive.Clear();
        }

        void SingleCircuitTWYCenterDriveParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\ " + CreateExcel8 + " " + CreateExcel1 + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel8 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "RMS2";
                ExcelSheet.Cells[2, 4] = "RMS1_LAST";
                ExcelSheet.Cells[2, 5] = "RMS2_LAST";
                ExcelSheet.Cells[2, 6] = "Val2";
                ExcelSheet.Cells[2, 7] = "T_CHECK";
                ExcelSheet.Cells[2, 8] = "LED_F1";
                ExcelSheet.Cells[2, 9] = "LED_F2";
                ExcelSheet.Cells[2, 10] = "Shock1";
                ExcelSheet.Cells[2, 11] = "Shock";
                ExcelSheet.Cells[2, 12] = "Short_Flag";
                ExcelSheet.Cells[2, 13] = "T";
                ExcelSheet.Cells[2, 14] = "Second";
                ExcelSheet.Cells[2, 15] = "A_MAX";
                ExcelSheet.Cells[2, 16] = "Error Code";

                //输出各个参数值
                for (int i = 0; i < RMS1SingleCircuitTWYCenterDrive.Count; i++)
                {

                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1SingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = RMS2SingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = RMS1LASTSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = RMS2LASTSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = Val2SingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = TCHECKSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = LEDF1SingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = LEDF2SingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = Shock1SingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = ShockSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = ShortFlagSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = TSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = AMaxSingleCircuitTWYCenterDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 16] = ErrorCodeSingleCircuitTWYCenterDrive[i].ToString();

                    if (SecondSingleCircuitTWYCenterDrive[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 14] = SecondSingleCircuitTWYCenterDrive[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 14] = ((int)SecondSingleCircuitTWYCenterDrive[i] / 3600).ToString() + ":" + (((int)SecondSingleCircuitTWYCenterDrive[i] % 3600) / 60).ToString() + ":" + (((int)SecondSingleCircuitTWYCenterDrive[i] % 3600) % 60).ToString();
                    }

                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearSingleCircuitTWYCenterDriveParameters();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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

        #region 立式跑道灯具驱动灯具状态信息解析
        private void ElevatedRWYLightDriveLampDataAnalysis(byte[] CompleteData)
        {
            byte[][] DataArray;
            ArrayList commandCount = new ArrayList();

            for (int i = 0; i < CompleteData.Length; i++)
            {
                if (CompleteData[i] == 0x02 && CompleteData[i + 1] == 0xAA && CompleteData[i + 2] == 0x01 && CompleteData[i + 3] == 0x09 && CompleteData[i + 4] == 0x09)
                {
                    commandCount.Add(i);
                }
            }

            DataArray = new byte[commandCount.Count][];

            for (int i = 0; i < commandCount.Count; i++)
            {
                if (i < commandCount.Count - 1)
                {
                    DataArray[i] = new byte[(int)commandCount[i + 1] - (int)commandCount[i]];
                }
                else
                {
                    DataArray[i] = new byte[CompleteData.Length - (int)commandCount[i]];
                }

                for (int j = 0; j < DataArray[i].Length; j++)
                {
                    DataArray[i][j] = CompleteData[(int)commandCount[i] + j];
                }
            }

            for (int i = 0; i < DataArray.Length; i++)
            {
                if (DataArray[i].Length == 32)
                {
                    byte checkOutValue = CalculateCheckOutValue(DataArray[i]);
                    if (checkOutValue == DataArray[i][DataArray[i].Length - 1])
                    {
                        RMS1ElevatedRWYLightDrive.Add(DataArray[i][5] * 500);
                        RMS1LASTElevatedRWYLightDrive.Add(DataArray[i][6] * 4);
                        Val2ElevatedRWYLightDrive.Add(DataArray[i][7] * 20);
                        TCHECKElevatedRWYLightDrive.Add(DataArray[i][8] * 16);
                        LEDF1ElevatedRWYLightDrive.Add(DataArray[i][9]);
                        LEDVSNS1ElevatedRWYLightDrive.Add(DataArray[i][10] * 16);
                        LEDVSNS2ElevatedRWYLightDrive.Add(DataArray[i][11] * 16);
                        TempertureElevatedRWYLightDrive.Add(DataArray[i][12]);
                        HumidityElevatedRWYLightDrive.Add(DataArray[i][13]);
                        AMaxElevatedRWYLightDrive.Add(DataArray[i][14] * 16);
                        ShortFlagElevatedRWYLightDrive.Add(DataArray[i][15]);
                        TElevatedRWYLightDrive.Add((SByte)DataArray[i][16]);
                

                        int SecondResult = 0;
                        for (int j = 0; j < 4; j++)
                        {
                            int SecondOrigin = DataArray[i][17 + j];
                            SecondResult |= SecondOrigin;
                            if (j < 3)
                            {
                                SecondResult <<= 8;
                            }
                        }
                        SecondElevatedRWYLightDrive.Add(SecondResult);
                        ErrorCodeElevatedRWYLightDrive.Add("No Error");
                    }
                    else
                    {
                        ElevatedRWYLightDriveLampCheckValueErrorHandle();
                    }
                }
                else
                {
                    ElevatedRWYLightDriveLampCommandLengthErrorHandle();
                }
            }

        }

        private void ElevatedRWYLightDriveLampCheckValueErrorHandle()
        {
            RMS1ElevatedRWYLightDrive.Add("Null");
            RMS1LASTElevatedRWYLightDrive.Add("Null");
            Val2ElevatedRWYLightDrive.Add("Null");
            TCHECKElevatedRWYLightDrive.Add("Null");
            LEDF1ElevatedRWYLightDrive.Add("Null");
            LEDVSNS1ElevatedRWYLightDrive.Add("Null");
            LEDVSNS2ElevatedRWYLightDrive.Add("Null");
            TempertureElevatedRWYLightDrive.Add("Null");
            HumidityElevatedRWYLightDrive.Add("Null");
            AMaxElevatedRWYLightDrive.Add("Null");
            ShortFlagElevatedRWYLightDrive.Add("Null");
            TElevatedRWYLightDrive.Add("Null");
            SecondElevatedRWYLightDrive.Add("Null");
            ErrorCodeElevatedRWYLightDrive.Add("Check Value Error");            
        }

        private void ElevatedRWYLightDriveLampCommandLengthErrorHandle()
        {
            RMS1ElevatedRWYLightDrive.Add("Null");
            RMS1LASTElevatedRWYLightDrive.Add("Null");
            Val2ElevatedRWYLightDrive.Add("Null");
            TCHECKElevatedRWYLightDrive.Add("Null");
            LEDF1ElevatedRWYLightDrive.Add("Null");
            LEDVSNS1ElevatedRWYLightDrive.Add("Null");
            LEDVSNS2ElevatedRWYLightDrive.Add("Null");
            TempertureElevatedRWYLightDrive.Add("Null");
            HumidityElevatedRWYLightDrive.Add("Null");
            AMaxElevatedRWYLightDrive.Add("Null");
            ShortFlagElevatedRWYLightDrive.Add("Null");
            TElevatedRWYLightDrive.Add("Null");
            SecondElevatedRWYLightDrive.Add("Null");
            ErrorCodeElevatedRWYLightDrive.Add("Command Length Error");            
        }

        private void ClearElevatedRWYLightDriveParameters()
        {
            RMS1ElevatedRWYLightDrive.Clear();
            RMS1LASTElevatedRWYLightDrive.Clear();
            Val2ElevatedRWYLightDrive.Clear();
            TCHECKElevatedRWYLightDrive.Clear();
            LEDF1ElevatedRWYLightDrive.Clear();
            LEDVSNS1ElevatedRWYLightDrive.Clear();
            LEDVSNS2ElevatedRWYLightDrive.Clear();
            TempertureElevatedRWYLightDrive.Clear();
            HumidityElevatedRWYLightDrive.Clear();
            AMaxElevatedRWYLightDrive.Clear();
            ShortFlagElevatedRWYLightDrive.Clear();
            TElevatedRWYLightDrive.Clear();
            SecondElevatedRWYLightDrive.Clear();
            ErrorCodeElevatedRWYLightDrive.Clear();            
        }

        void ElevatedRWYLightDriveParametersCreatExcel()
        {
            try
            {
                //创建excel模板
                str_fileName = "d:\\ " + CreateExcel9 + " " + CreateExcel1 + " " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
                ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
                ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

                //设置Excel列名           
                ExcelSheet.Cells[1, 1] = CreateExcel9 + " " + CreateExcel1;
                ExcelSheet.Cells[2, 1] = CreateExcel2;
                ExcelSheet.Cells[2, 2] = "RMS1";
                ExcelSheet.Cells[2, 3] = "RMS1_LAST";
                ExcelSheet.Cells[2, 4] = "Val2";
                ExcelSheet.Cells[2, 5] = "T_CHECK";
                ExcelSheet.Cells[2, 6] = "LED_F1";
                ExcelSheet.Cells[2, 7] = "LED_VSNS1";
                ExcelSheet.Cells[2, 8] = "LED_VSNS2";
                ExcelSheet.Cells[2, 9] = "Temperture";
                ExcelSheet.Cells[2, 10] = "Humidity";
                ExcelSheet.Cells[2, 11] = "A_MAX";
                ExcelSheet.Cells[2, 12] = "Short_Flag";
                ExcelSheet.Cells[2, 13] = "T";
                ExcelSheet.Cells[2, 14] = "Second";
                ExcelSheet.Cells[2, 15] = "Error Code";
                

                //输出各个参数值
                for (int i = 0; i < RMS1ElevatedRWYLightDrive.Count; i++)
                {

                    ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                    ExcelSheet.Cells[3 + i, 2] = RMS1ElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 3] = RMS1LASTElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 4] = Val2ElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 5] = TCHECKElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 6] = LEDF1ElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 7] = LEDVSNS1ElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 8] = LEDVSNS2ElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 9] = TempertureElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 10] = HumidityElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 11] = AMaxElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 12] = ShortFlagElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 13] = TElevatedRWYLightDrive[i].ToString();
                    ExcelSheet.Cells[3 + i, 15] = ErrorCodeElevatedRWYLightDrive[i].ToString();

                    if (SecondElevatedRWYLightDrive[i].ToString() == "Null")
                    {
                        ExcelSheet.Cells[3 + i, 14] = SecondElevatedRWYLightDrive[i].ToString();
                    }
                    else
                    {
                        ExcelSheet.Cells[3 + i, 14] = ((int)SecondElevatedRWYLightDrive[i] / 3600).ToString() + ":" + (((int)SecondElevatedRWYLightDrive[i] % 3600) / 60).ToString() + ":" + (((int)SecondElevatedRWYLightDrive[i] % 3600) % 60).ToString();
                    }

                }

                ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
                ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
                ExcelApp.Quit();                                                                      //退出Excel应用程序    

                ClearElevatedRWYLightDriveParameters();

                ShowEXCELHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowEXCELHandleProcess.Visibility = Visibility.Hidden;
                }));

                this.Dispatcher.Invoke(new System.Action(() =>
                {
                    if (MessageBox.Show(MessageboxContent28, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
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
                    if (MessageBox.Show(MessageboxContent29, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            RefreshStringMessageLanguage();
            if (judgeFeedbackCommand==3)
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

                    string FileName = "d:\\" + CreateTxt1 + " "+DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                    FileStream aFile = new FileStream(FileName, FileMode.Create);
                    StreamWriter sw = new StreamWriter(aFile);

                    sw.Write(OriginalData);
                    sw.Close();

                    ShowTXTHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                    {
                        ShowTXTHandleProcess.Visibility = Visibility.Hidden;
                    }));

                    if ( MessageBox.Show(MessageboxContent30, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Information)==MessageBoxResult.OK)
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
                    ShowTXTHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                    {
                        ShowTXTHandleProcess.Visibility = Visibility.Hidden;
                    }));

                    if ( MessageBox.Show(MessageboxContent26, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                ShowTXTHandleProcess.Dispatcher.Invoke(new System.Action(() =>
                {
                    ShowTXTHandleProcess.Visibility = Visibility.Hidden;
                }));

                if ( MessageBox.Show(MessageboxContent31, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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

        #region 工厂模式参数设置
        private void SetLightParametersInFactoryMode_Click(object sender, RoutedEventArgs e)
        {
            RefreshStringMessageLanguage();

            string judgeResult = "";
            if(hardwareVersion1==12&&softwareNumber==3)
            {
                judgeResult = MessageboxContent43;
            }
            else if(hardwareVersion1 == 5 && softwareNumber == 6)
            {
                judgeResult = MessageboxContent45;
            }
            else if (hardwareVersion1 == 12 && softwareNumber == 0)
            {
                judgeResult = MessageboxContent46;
            }
            else
            {
                judgeResult = MessageboxContent32;
            }

            if (MessageBox.Show(judgeResult, MessageboxHeader1, MessageBoxButton.OK,MessageBoxImage.Information)==MessageBoxResult.OK)
            {
                ConfigurationWindow.IsEnabled = true;            

                if(ConfirmLampName.Text!=""&&ConfirmLampModel.Text!=""&&ConfirmSettingOpenCircuitParameter.Text!="")
                {
                    if(hardwareVersion1 == 8 && softwareNumber == 4)
                    {
                        settingChannel = Convert.ToByte(ChannelSelect.SelectedIndex);
                        settingFlashFrequency = Convert.ToByte(FlashFrequencySelect.SelectedItem);
                        settingWaveform = Convert.ToByte(WaveformSelect.SelectedIndex);

                        ConfigureRWYGuardLightSettingParametersCommand();
                    }
                    else if(hardwareVersion1==9&&softwareNumber==9)
                    {
                        settingIICFLAG = Convert.ToByte(IICFlagSelect.SelectedIndex);

                        ConfigureElevatedRWYLightSettingParametersCommand();
                    }
                    else
                    {
                        ConfigureSettingParametersCommand();
                    }                    

                    MessageBoxResult result = MessageBox.Show(MessageboxContent33, MessageboxHeader2, MessageBoxButton.YesNo, MessageBoxImage.Question);
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
                            if( MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                    if( MessageBox.Show(MessageboxContent34, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
            else
            {
                ConfigurationWindow.IsEnabled = false;
            }
        }

        #region 灯具名称
        private void SelectApproachChenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupApproachChenterlineLight.Visibility = Visibility.Visible;            

            SelectAPPS12SLEDC.IsChecked = false;

            OtherConfigurationAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectApproachChenterlineLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";             
            }));
        }

        private void SelectApproachCrossbarLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupApproachCrossbarLight.Visibility = Visibility.Visible;            

            SelectAPPS12LLEDC.IsChecked = false;
            SelectAPPS12RLEDC.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectApproachCrossbarLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectApproachSideRowLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupApproachSideRowLight.Visibility = Visibility.Visible;

            SelectAPSS12LLEDR.IsChecked = false;
            SelectAPSS12RLEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectApproachSideRowLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRWYThresholdWingBarLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYThresholdWingBarLight.Visibility = Visibility.Visible;            

            SelectTHWS12LLEDG.IsChecked = false;
            SelectTHWS12RLEDG.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYThresholdWingBarLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRWYThresholdLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYThresholdLight.Visibility = Visibility.Visible;

            SelectTHRS12LLEDG.IsChecked = false;
            SelectTHRS12RLEDG.IsChecked = false;
            SelectTHRS12SLEDG.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYThresholdLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYEdgeLight.Visibility = Visibility.Visible;           

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
            SelectRELC12LEDRRB1P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYEdgeLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void Select12inchesRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            Group12inchesRWYEndLight.Visibility = Visibility.Visible;
           
            SelectENDS12LEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelect12inchesRWYEndLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRWYThresholdEndLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYThresholdEndLight.Visibility = Visibility.Visible;
            
            SelectTAES12LLEDGR1P.IsChecked = false;
            SelectTAES12RLEDGR1P.IsChecked = false;
            SelectTAES12SLEDGR1P.IsChecked = false;
            SelectTAES12LLEDGRMR2P.IsChecked = false;
            SelectTAES12RLEDGRMR2P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYThresholdEndLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRWYCenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYCenterlineLight.Visibility = Visibility.Visible;
           
            SelectRCLS08LEDCB1P.IsChecked = false;
            SelectRCLS08LEDRB1P.IsChecked = false;
            SelectRCLS08LEDCC1P.IsChecked = false;
            SelectRCLS08LEDRC1P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYCenterlineLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void Select12inchesRWYCenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            Group12inchesRWYCenterlineLight.Visibility = Visibility.Visible;

            SelectRCLS12LEDCCMR2P.IsChecked = false;
            SelectRCLS12LEDRCMR2P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelect12inchesRWYCenterlineLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRWYTouchdownZoneLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Visible;
            
            SelectTDZS08LLEDC.IsChecked = false;
            SelectTDZS08RLEDC.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYTouchdownZoneLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void Select8inchesRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            Group8inchesRWYEndLight.Visibility = Visibility.Visible;
            
            SelectENDS08LEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelect8inchesRWYEndLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectRapidExitTWYIndicatorLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Visible;
           
            SelectRAPS08LEDY.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRapidExitTWYIndicatorLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectCombinedRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Visible;

            SelectRELC12LEDCYC1P.IsChecked = false;
            SelectRELC12LEDCCC1P.IsChecked = false;
            SelectRELC12LEDCRC1P.IsChecked = false;
            SelectRELC12LEDRYC1P.IsChecked = false;
            SelectRELC12LEDCYB1P.IsChecked = false;
            SelectRELC12LEDCCB1P.IsChecked = false;
            SelectRELC12LEDCRB1P.IsChecked = false;
            SelectRELC12LEDRYB1P.IsChecked = false;
            SelectRELC12LEDCBC1P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectCombinedRWYEdgeLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }
       
        private void SelectRWYGuardLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupRWYGuardLight.Visibility = Visibility.Visible;
            
            SelectHRGS08LEDY.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectRWYGuardLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectTWYCenterLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupTWYCenterlineLight.Visibility = Visibility.Visible;
            
            SelectTCLMS08SLEDGG1P.IsChecked = false;
            SelectTCLMS08SLEDGY1P.IsChecked = false;
            SelectTCLMS08SLEDYY1P.IsChecked = false;
            SelectTCLMS08SLEDYB1P.IsChecked = false;
            SelectTCLMS08SLEDGB1P.IsChecked = false;
            SelectTCLMS08CLEDGG1P.IsChecked = false;
            SelectTCLMS08CLEDGY1P.IsChecked = false;
            SelectTCLMS08CLEDYY1P.IsChecked = false;
            SelectTCLMS08CLEDYB1P.IsChecked = false;
            SelectTCLMS08CLEDGB1P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectTWYCenterLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectTWYCenterLight2P_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupTWYCenterlineLight2P.Visibility = Visibility.Visible;

            SelectTCLMS08SLEDGG2P.IsChecked = false;
            SelectTCLMS08SLEDGY2P.IsChecked = false;
            SelectTCLMS08SLEDYY2P.IsChecked = false;
            SelectTCLMS08SLEDYB2P.IsChecked = false;
            SelectTCLMS08SLEDGB2P.IsChecked = false;
            SelectTCLMS08CLEDGG2P.IsChecked = false;
            SelectTCLMS08CLEDGY2P.IsChecked = false;
            SelectTCLMS08CLEDYY2P.IsChecked = false;
            SelectTCLMS08CLEDYB2P.IsChecked = false;
            SelectTCLMS08CLEDGB2P.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectTWYCenterLight2P.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectTWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupTWYEdgeLight.Visibility = Visibility.Visible;
            
            SelectTOEL08LEDB.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectTWYEdgeLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectTWYStopBarLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupTWYStopBarLight.Visibility = Visibility.Visible;
            
            SelectSBLMS08SLEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectTWYStopBarLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectIntermediateHoldingPositionLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupIntermediateHoldingPositionLight.Visibility = Visibility.Visible;
            
            SelectTPLMS08SLEDY.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectIntermediateHoldingPositionLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectTWYIntersectionsLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupTWYIntersectionsLight.Visibility = Visibility.Visible;
            
            SelectTOIL08LLEDY.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectTWYIntersectionsLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedApproachCenterlineLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedApproachCenterlineLight.Visibility = Visibility.Visible;
            
            SelectEULAPLEDCCenterline.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedApproachCenterlineLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedApproachCrossbarLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedApproachCrossbarLight.Visibility = Visibility.Visible;
            
            SelectEULAPLEDCCrossbar.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedApproachCrossbarLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedApproachSideRowLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedApproachSideRowLight.Visibility = Visibility.Visible;
            
            SelectEULSRLEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedApproachSideRowLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedRWYEdgeLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedRWYEdgeLight.Visibility = Visibility.Visible;
            
            SelectEBLRELEDYC.IsChecked = false;
            SelectEBLRELEDCY.IsChecked = false;
            SelectEBLRELEDCC.IsChecked = false;
            SelectEBLRELEDCR.IsChecked = false;
            SelectEBLRELEDRC.IsChecked = false;


            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedRWYEdgeLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedRWYThresholdLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedRWYThresholdLight.Visibility = Visibility.Visible;
            
            SelectEULTHLEDG.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedRWYThresholdLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedRWYThresholdWingbarLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedRWYThresholdWingbarLight.Visibility = Visibility.Visible;
            
            SelectEULTHWLEDG.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedRWYThresholdWingbarLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedRWYEndLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedRWYEndLight.Visibility = Visibility.Visible;
            
            SelectEULEDLEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedRWYEndLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectElevatedTWYStopBarLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupElevatedTWYStopBarLight.Visibility = Visibility.Visible;

            SelectEULSBLEDR.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectElevatedTWYStopBarLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        private void SelectAircraftStandManoeuvringGuidanceLight_Checked(object sender, RoutedEventArgs e)
        {
            ALLGroupCollapsed();
            GroupAircraftStandManoeuvringGuidanceLight.Visibility = Visibility.Visible;

            SelectASMG08LEDYM.IsChecked = false;

            OtherConfigurationAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampName.Text = TBSelectAircraftStandManoeuvringGuidanceLight.Text.ToString();
                ConfirmLampModel.Text = "";
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));
        }

        public void ALLGroupCollapsed()
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
            Group12inchesRWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupRWYTouchdownZoneLight.Visibility = Visibility.Collapsed;
            Group8inchesRWYEndLight.Visibility = Visibility.Collapsed;
            GroupRapidExitTWYIndicatorLight.Visibility = Visibility.Collapsed;
            GroupCombinedRWYEdgeLight.Visibility = Visibility.Collapsed;
            GroupRWYGuardLight.Visibility = Visibility.Collapsed;
            GroupTWYCenterlineLight.Visibility = Visibility.Collapsed;
            GroupTWYCenterlineLight2P.Visibility = Visibility.Collapsed;
            GroupTWYStopBarLight.Visibility = Visibility.Collapsed;
            GroupIntermediateHoldingPositionLight.Visibility = Visibility.Collapsed;
            GroupTWYIntersectionsLight.Visibility = Visibility.Collapsed;
            GroupTWYEdgeLight.Visibility = Visibility.Collapsed;
            GroupElevatedApproachCenterlineLight.Visibility = Visibility.Collapsed;
            GroupElevatedApproachCrossbarLight.Visibility = Visibility.Collapsed;
            GroupElevatedApproachSideRowLight.Visibility = Visibility.Collapsed;
            GroupElevatedRWYEdgeLight.Visibility = Visibility.Collapsed;
            GroupElevatedRWYEndLight.Visibility = Visibility.Collapsed;
            GroupElevatedRWYThresholdLight.Visibility = Visibility.Collapsed;
            GroupElevatedRWYThresholdWingbarLight.Visibility = Visibility.Collapsed;
            GroupElevatedTWYStopBarLight.Visibility = Visibility.Collapsed;
            GroupAircraftStandManoeuvringGuidanceLight.Visibility = Visibility.Collapsed;
        }

        public void OtherConfigurationAllCollapsed()
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Collapsed;
            SelectOpenCircuitFalse.Visibility = Visibility.Collapsed;
            FlashFrequencySelectLabel.Visibility = Visibility.Collapsed;
            FlashFrequencySelectBorder.Visibility = Visibility.Collapsed;
            FlashFrequencySelect.Visibility = Visibility.Collapsed;
            ChannelSelectLabel.Visibility = Visibility.Collapsed;
            ChannelSelectBorder.Visibility = Visibility.Collapsed;
            ChannelSelect.Visibility = Visibility.Collapsed;
            WaveformSelectLabel.Visibility = Visibility.Collapsed;
            WaveformSelectBorder.Visibility = Visibility.Collapsed;
            WaveformSelect.Visibility = Visibility.Collapsed;
            IICFlagSelectLabel.Visibility = Visibility.Collapsed;
            IICFlagSelectBorder.Visibility = Visibility.Collapsed;
            IICFlagSelect.Visibility = Visibility.Collapsed;
        }

       
        #endregion

        #region 灯具型号
        private void SelectAPPS12SLEDC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();            

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPPS12SLEDC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureAPPS12SLEDCParameters();
        }

        private void SelectAPPS12LLEDC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();
           
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPPS12LLEDC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureAPPS12LLEDCParameters();
        }

        private void SelectAPPS12RLEDC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPPS12RLEDC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureAPPS12RLEDCParameters();
        }

        private void SelectAPSS12LLEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPSS12LLEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureAPSS12LLEDRParameters();
        }

        private void SelectAPSS12RLEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectAPSS12RLEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureAPSS12RLEDRParameters();
        }

        private void SelectTHWS12LLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHWS12LLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTHWS12LLEDGParameters();
        }

        private void SelectTHWS12RLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHWS12RLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTHWS12RLEDGParameters();
        }

        private void SelectTHRS12LLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHRS12LLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTHRS12LLEDGParameters();
        }

        private void SelectTHRS12RLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHRS12RLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTHRS12RLEDGParameters();
        }

        private void SelectTHRS12SLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTHRS12SLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTHRS12SLEDGParameters();
        }

        private void SelectRELS12LLEDYC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDYC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12LLEDYCParameters();
        }

        private void SelectRELS12RLEDYC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDYC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12RLEDYCParameters();
        }

        private void SelectRELS12LLEDCY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDCY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12LLEDCYParameters();
        }

        private void SelectRELS12RLEDCY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDCY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12RLEDCYParameters();
        }

        private void SelectRELS12LLEDCC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDCC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12LLEDCCParameters();
        }

        private void SelectRELS12RLEDCC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDCC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12RLEDCCParameters();
        }

        private void SelectRELS12LLEDCR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDCR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12LLEDCRParameters();
        }

        private void SelectRELS12RLEDCR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDCR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12RLEDCRParameters();
        }

        private void SelectRELS12LLEDRC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12LLEDRC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12LLEDRCParameters();
        }

        private void SelectRELS12RLEDRC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELS12RLEDRC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELS12RLEDRCParameters();
        }

        private void SelectRELC12LEDRRB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDRRB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDRRB1PParameters();
        }

        private void SelectENDS12LEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectENDS12LEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureENDS12LEDRParameters();
        }

        private void SelectTAES12LLEDGR1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTAES12LLEDGR1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTAES12LLEDGR1PParameters();
        }

        private void SelectTAES12RLEDGR1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTAES12RLEDGR1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTAES12RLEDGR1PParameters();
        }

        private void SelectTAES12SLEDGR1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTAES12SLEDGR1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTAES12SLEDGR1PParameters();
        }

        private void SelectTAES12LLEDGRMR2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = TBSelectTAES12LLEDGRMR2P.Text.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTAES12LLEDGRMR2PParameters();
        }

        private void SelectTAES12RLEDGRMR2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = TBSelectTAES12RLEDGRMR2P.Text.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTAES12RLEDGRMR2PParameters();
        }

        private void SelectRCLS08LEDCB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDCB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRCLS08LEDCB1PParameters();
        }

        private void SelectRCLS08LEDRB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDRB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRCLS08LEDRB1PParameters();
        }

        private void SelectRCLS08LEDCC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDCC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRCLS08LEDCC1PParameters();
        }

        private void SelectRCLS08LEDRC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRCLS08LEDRC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRCLS08LEDRC1PParameters();
        }

        private void SelectRCLS12LEDCCMR2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = TBSelectRCLS12LEDCCMR2P.Text.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRCLS12LEDCCMR2PParameters();
        }

        private void SelectRCLS12LEDRCMR2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = TBSelectRCLS12LEDRCMR2P.Text.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRCLS12LEDRCMR2PParameters();
        }

        private void SelectTDZS08LLEDC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTDZS08LLEDC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTDZS08LLEDCParameters();
        }

        private void SelectTDZS08RLEDC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTDZS08RLEDC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTDZS08RLEDCParameters();
        }

        private void SelectENDS08LEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectENDS08LEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureENDS08LEDRParameters();
        }

        private void SelectRAPS08LEDY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRAPS08LEDY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRAPS08LEDYParameters();
        }

        private void SelectRELC12LEDCYC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCYC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCYC1PParameters();
        }

        private void SelectRELC12LEDCCC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCCC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCCC1PParameters();
        }

        private void SelectRELC12LEDCRC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCRC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCRC1PParameters();
        }

        private void SelectRELC12LEDRYC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDRYC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDRYC1PParameters();
        }

        private void SelectRELC12LEDCYB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCYB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCYB1PParameters();
        }

        private void SelectRELC12LEDCCB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCCB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCCB1PParameters();
        }

        private void SelectRELC12LEDCRB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCRB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCRB1PParameters();
        }

        private void SelectRELC12LEDRYB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDRYB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDRYB1PParameters();
        }

        private void SelectRELC12LEDCBC1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectRELC12LEDCBC1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureRELC12LEDCBC1PParameters();
        }

        private void SelectHRGS08LEDY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();
            FlashFrequencySelectLabel.Visibility = Visibility.Visible;
            FlashFrequencySelectBorder.Visibility = Visibility.Visible;
            FlashFrequencySelect.Visibility = Visibility.Visible;
            ChannelSelectLabel.Visibility = Visibility.Visible;
            ChannelSelectBorder.Visibility = Visibility.Visible;
            ChannelSelect.Visibility = Visibility.Visible;
            WaveformSelectLabel.Visibility = Visibility.Visible;
            WaveformSelectBorder.Visibility = Visibility.Visible;
            WaveformSelect.Visibility = Visibility.Visible;
            

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectHRGS08LEDY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";
            }));

            ConfigureHRGS08LEDYParameters();
        }

        private void SelectTCLMS08SLEDGG1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDGG1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDGG1PParameters();
        }

        private void SelectTCLMS08SLEDGY1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDGY1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDGY1PParameters();
        }

        private void SelectTCLMS08SLEDYY1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDYY1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDYY1PParameters();
        }

        private void SelectTCLMS08SLEDYB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDYB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDYB1PParameters();
        }

        private void SelectTCLMS08SLEDGB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDGB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDGB1PParameters();
        }

        private void SelectTCLMS08CLEDGG1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDGG1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDGG1PParameters();
        }

        private void SelectTCLMS08CLEDGY1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDGY1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDGY1PParameters();
        }

        private void SelectTCLMS08CLEDYY1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDYY1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDYY1PParameters();
        }

        private void SelectTCLMS08CLEDYB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDYB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDYB1PParameters();
        }

        private void SelectTCLMS08CLEDGB1P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDGB1P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDGB1PParameters();
        }

        private void SelectTCLMS08SLEDGG2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDGG2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDGG2PParameters();
        }

        private void SelectTCLMS08SLEDGY2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDGY2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDGY2PParameters();
        }

        private void SelectTCLMS08SLEDYY2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDYY2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDYY2PParameters();
        }

        private void SelectTCLMS08SLEDYB2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDYB2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDYB2PParameters();
        }

        private void SelectTCLMS08SLEDGB2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08SLEDGB2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08SLEDGB2PParameters();
        }

        private void SelectTCLMS08CLEDGG2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDGG2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDGG2PParameters();
        }

        private void SelectTCLMS08CLEDGY2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDGY2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDGY2PParameters();
        }

        private void SelectTCLMS08CLEDYY2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDYY2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDYY2PParameters();
        }

        private void SelectTCLMS08CLEDYB2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDYB2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDYB2PParameters();
        }

        private void SelectTCLMS08CLEDGB2P_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTCLMS08CLEDGB2P.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTCLMS08CLEDGB2PParameters();
        }

        private void SelectSBLMS08SLEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();


            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectSBLMS08SLEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureSBLMS08SLEDRParameters();
        }

        private void SelectTPLMS08SLEDY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTPLMS08SLEDY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTPLMS08SLEDYParameters();
        }

        private void SelectTOIL08LLEDY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTOIL08LLEDY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTOIL08LLEDYParameters();
        }

        private void SelectTOEL08LEDB_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectTOEL08LEDB.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureTOEL08LEDBParameters();
        }

        //coding now...
        private void SelectEULAPLEDCCenterline_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULAPLEDCCenterline.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULAPLEDCCenterlineParameters();
        }

        private void SelectEULAPLEDCCrossbar_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULAPLEDCCrossbar.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULAPLEDCCrossbarParameters();
        }

        private void SelectEULSRLEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULSRLEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULSRLEDRParameters();
        }

        private void SelectEBLRELEDYC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEBLRELEDYC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEBLRELEDYCParameters();
        }

        private void SelectEBLRELEDCY_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEBLRELEDCY.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEBLRELEDCYParameters();
        }

        private void SelectEBLRELEDCC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEBLRELEDCC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEBLRELEDCCParameters();
        }

        private void SelectEBLRELEDCR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEBLRELEDCR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEBLRELEDCRParameters();
        }

        private void SelectEBLRELEDRC_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEBLRELEDRC.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEBLRELEDRCParameters();
        }

        private void SelectEULTHLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULTHLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULTHLEDGParameters();
        }

        private void SelectEULTHWLEDG_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULTHWLEDG.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULTHWLEDGParameters();
        }

        private void SelectEULEDLEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULEDLEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULEDLEDRParameters();
        }

        private void SelectEULSBLEDR_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = false;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();

            IICFlagSelectLabel.Visibility = Visibility.Visible;
            IICFlagSelectBorder.Visibility = Visibility.Visible;
            IICFlagSelect.Visibility = Visibility.Visible;

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectEULSBLEDR.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureEULSBLEDRParameters();
        }

        private void SelectASMG08LEDYM_Checked(object sender, RoutedEventArgs e)
        {
            SelectOpenCircuitTrue.Visibility = Visibility.Visible;
            SelectOpenCircuitFalse.Visibility = Visibility.Visible;

            SelectOpenCircuitTrue.IsEnabled = true;
            SelectOpenCircuitFalse.IsEnabled = true;
            SelectOpenCircuitTrue.IsChecked = false;
            SelectOpenCircuitFalse.IsChecked = false;

            OtherConfigurationVisibilityAllCollapsed();            

            this.Dispatcher.Invoke(new System.Action(() =>
            {
                ConfirmLampModel.Text = SelectASMG08LEDYM.Content.ToString();
                ConfirmSettingOpenCircuitParameter.Text = "";

            }));

            ConfigureASMG08LEDYMParameters();
        }

        public void OtherConfigurationVisibilityAllCollapsed()
        {
            FlashFrequencySelectLabel.Visibility = Visibility.Collapsed;
            FlashFrequencySelectBorder.Visibility = Visibility.Collapsed;
            FlashFrequencySelect.Visibility = Visibility.Collapsed;
            ChannelSelectLabel.Visibility = Visibility.Collapsed;
            ChannelSelectBorder.Visibility = Visibility.Collapsed;
            ChannelSelect.Visibility = Visibility.Collapsed;
            WaveformSelectLabel.Visibility = Visibility.Collapsed;
            WaveformSelectBorder.Visibility = Visibility.Collapsed;
            WaveformSelect.Visibility = Visibility.Collapsed;
            IICFlagSelectLabel.Visibility = Visibility.Collapsed;
            IICFlagSelectBorder.Visibility = Visibility.Collapsed;
            IICFlagSelect.Visibility = Visibility.Collapsed;
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

        private void ConfigureRELC12LEDRRB1PParameters()
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
            settingLampsNumber = 0x2E;
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

        private void ConfigureTAES12LLEDGRMR2PParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x07;
            settingIA[2] = 0x06;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x07;
            settingIB[2] = 0x06;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x08;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x2A;
        }

        private void ConfigureTAES12RLEDGRMR2PParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x07;
            settingIA[2] = 0x06;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x07;
            settingIB[2] = 0x06;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x08;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x2B;
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

        private void ConfigureRCLS12LEDCCMR2PParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x07;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x07;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x2C;
        }

        private void ConfigureRCLS12LEDRCMR2PParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x03;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x07;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x2D;
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

        private void ConfigureRELC12LEDCBC1PParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x09;
            settingIA[2] = 0x03;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
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
            settingLampsNumber = 0x2F;
        }
        

        private void ConfigureHRGS08LEDYParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x08;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIB[0] = 0x00;
            settingIB[1] = 0x00;
            settingIB[2] = 0x00;
            settingIB[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x08;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingIIB[0] = 0x00;
            settingIIB[1] = 0x00;
            settingIIB[2] = 0x00;
            settingIIB[3] = 0x00;
            settingReadRFlag = 0x01;            
            settingLampsNumber = 0x29;
        }

        private void ConfigureTCLMS08SLEDGG1PParameters()
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
            settingLampsNumber = 0x30;
        }

        private void ConfigureTCLMS08SLEDGY1PParameters()
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
            settingLampsNumber = 0x31;
        }

        private void ConfigureTCLMS08SLEDYY1PParameters()
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
            settingLampsNumber = 0x32;
        }

        private void ConfigureTCLMS08SLEDYB1PParameters()
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
            settingLampsNumber = 0x33;
        }

        private void ConfigureTCLMS08SLEDGB1PParameters()
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
            settingLampsNumber = 0x34;
        }

        private void ConfigureTCLMS08CLEDGG1PParameters()
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
            settingLampsNumber = 0x35;
        }

        private void ConfigureTCLMS08CLEDGY1PParameters()
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
            settingLampsNumber = 0x36;
        }

        private void ConfigureTCLMS08CLEDYY1PParameters()
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
            settingLampsNumber = 0x37;
        }

        private void ConfigureTCLMS08CLEDYB1PParameters()
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
            settingLampsNumber = 0x38;
        }

        private void ConfigureTCLMS08CLEDGB1PParameters()
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
            settingLampsNumber = 0x39;
        }

        private void ConfigureTCLMS08SLEDGG2PParameters()
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
            settingLampsNumber = 0x3A;
        }

        private void ConfigureTCLMS08SLEDGY2PParameters()
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
            settingLampsNumber = 0x3B;
        }

        private void ConfigureTCLMS08SLEDYY2PParameters()
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
            settingLampsNumber = 0x3C;
        }

        private void ConfigureTCLMS08SLEDYB2PParameters()
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
            settingLampsNumber = 0x3D;
        }

        private void ConfigureTCLMS08SLEDGB2PParameters()
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
            settingLampsNumber = 0x3E;
        }

        private void ConfigureTCLMS08CLEDGG2PParameters()
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
            settingLampsNumber = 0x3F;
        }

        private void ConfigureTCLMS08CLEDGY2PParameters()
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
            settingLampsNumber = 0x40;
        }

        private void ConfigureTCLMS08CLEDYY2PParameters()
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
            settingLampsNumber = 0x41;
        }

        private void ConfigureTCLMS08CLEDYB2PParameters()
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
            settingLampsNumber = 0x42;
        }

        private void ConfigureTCLMS08CLEDGB2PParameters()
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
            settingLampsNumber = 0x43;
        }

        private void ConfigureSBLMS08SLEDRParameters()
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
            settingLampsNumber = 0x45;
        }

        private void ConfigureTPLMS08SLEDYParameters()
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
            settingLampsNumber = 0x46;
        }

        private void ConfigureTOIL08LLEDYParameters()
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
            settingLampsNumber = 0x47;
        }

        private void ConfigureTOEL08LEDBParameters()
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
            settingLampsNumber = 0x44;
        }

        public void ConfigureEULAPLEDCCenterlineParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x48;
        }

        public void ConfigureEULAPLEDCCrossbarParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x49;
        }

        public void ConfigureEULSRLEDRParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x4A;
        }

        public void ConfigureEBLRELEDYCParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x4B;
        }

        public void ConfigureEBLRELEDCYParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x4C;
        }

        public void ConfigureEBLRELEDCCParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x4D;
        }

        public void ConfigureEBLRELEDCRParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x4E;
        }

        public void ConfigureEBLRELEDRCParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x4F;
        }

        public void ConfigureEULTHLEDGParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x50;
        }

        public void ConfigureEULTHWLEDGParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x51;
        }

        public void ConfigureEULEDLEDRParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x52;
        }

        public void ConfigureEULSBLEDRParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x00;
            settingIA[2] = 0x00;
            settingIA[3] = 0x00;
            settingIIA[0] = 0x00;
            settingIIA[1] = 0x00;
            settingIIA[2] = 0x00;
            settingIIA[3] = 0x00;
            settingBreakVal1 = 0x00;
            settingBreakVal2 = 0x00;
            settingORIVOLT = 0x00;
            settingRMSSET[0] = 0x00;
            settingRMSSET[1] = 0x00;
            settingReadRFlag = 0x00;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x53;
        }

        public void ConfigureASMG08LEDYMParameters()
        {
            settingIA[0] = 0x00;
            settingIA[1] = 0x07;
            settingIA[2] = 0x06;
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
            settingReadRFlag = 0x01;
            settingMosFlag = 0x00;
            settingLampsNumber = 0x54;
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
            settingParameterCommand[0] = 0x02;
            settingParameterCommand[1] = 0x55;
            settingParameterCommand[2] = 0x11;
            settingParameterCommand[3] = 0x58;
            settingParameterCommand[4] = 0x12;

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

        //跑道警戒灯，生成设置参数指令
        public void ConfigureRWYGuardLightSettingParametersCommand()
        {
            settingParameterCommand[0] = 0x55;
            settingParameterCommand[1] = 0x02;
            settingParameterCommand[2] = 0x11;
            settingParameterCommand[3] = 0x58;
            settingParameterCommand[4] = 0x12;

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
            settingParameterCommand[22] = settingChannel;
            settingParameterCommand[23] = settingBreakFlag;
            settingParameterCommand[24] = settingLampsNumber;
            settingParameterCommand[25] = settingFlashFrequency;
            settingParameterCommand[26] = settingWaveform;
            settingParameterCommand[27] = CalculateCheckOutValue(settingParameterCommand);
        }

        //立式跑道灯具，生成设置参数指令
        public void ConfigureElevatedRWYLightSettingParametersCommand()
        {
            settingParameterCommand[0] = 0x03;
            settingParameterCommand[1] = 0x66;
            settingParameterCommand[2] = 0x11;
            settingParameterCommand[3] = 0x58;
            settingParameterCommand[4] = 0x12;

            for (int i = 0; i < settingIA.Length; i++)
            {
                settingParameterCommand[5 + i] = settingIA[i];
            }
            
            for (int i = 0; i < settingIIA.Length; i++)
            {
                settingParameterCommand[9 + i] = settingIIA[i];
            }

            settingParameterCommand[13] = settingBreakVal1;
            settingParameterCommand[14] = settingBreakVal2;
            settingParameterCommand[15] = settingORIVOLT;

            for (int i = 0; i < settingRMSSET.Length; i++)
            {
                settingParameterCommand[16 + i] = settingRMSSET[i];
            }

            settingParameterCommand[18] = settingIICFLAG;
            settingParameterCommand[21] = settingReadRFlag;
            settingParameterCommand[22] = settingMosFlag;
            settingParameterCommand[23] = settingBreakFlag;
            settingParameterCommand[24] = settingLampsNumber;
            settingParameterCommand[27] = CalculateCheckOutValue(settingParameterCommand);
        }

        #endregion

        #region 工厂模式下，恢复初始设置
        private void RestoreOriginalStatus_Click(object sender, RoutedEventArgs e)
        {
            RefreshStringMessageLanguage();
            MessageBoxResult result = MessageBox.Show(MessageboxContent35, MessageboxHeader2, MessageBoxButton.YesNo, MessageBoxImage.Question);

            if(result==MessageBoxResult.Yes)
            {
                ConfigurationWindow.IsEnabled = true;

                if (lampsPort.IsOpen)
                {
                    judgeFeedbackCommand = 1;

                    if((hardwareVersion1 == 8 && softwareNumber == 4))
                    {
                        //lampsPort.Write(InFactoryModeRWYGuardLightRestoreOriginalCommand, 0, 28);
                    }
                    else
                    {
                        lampsPort.Write(InFactoryModeCommonLightRestoreOriginalCommand, 0, 28);
                    }                    
                }
                else
                {
                    if (MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        ConfigurationWindow.IsEnabled = true;
                    }
                    else
                    {
                        ConfigurationWindow.IsEnabled = false;
                    }
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
        #endregion

        #endregion

        #region 开发者模式设置参数
        private void SetLightParametersInDeveloperMode_Click(object sender, RoutedEventArgs e)
        {
            RefreshStringMessageLanguage();
            ShowSetParameterCommand.Text = "";
            if (SetParameterIA.Text != "" && SetParameterIB.Text != "" && SetParameterIIA.Text != "" && SetParameterIIB.Text != "")
            {
                InDeveloperModeConfigureSettingParametersCommand();

                for (int i = 0; i < InDeveloperModeSettingParameterCommand.Length; i++)
                {
                    ShowSetParameterCommand.Text += Convert.ToString(InDeveloperModeSettingParameterCommand[i], 16).PadLeft(2, '0').ToUpper() + " ";
                }

                MessageBoxResult result = MessageBox.Show(MessageboxContent33, MessageboxHeader2, MessageBoxButton.YesNo, MessageBoxImage.Question);

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
                        if( MessageBox.Show(MessageboxContent9, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                if( MessageBox.Show(MessageboxContent36, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
            InDeveloperModeSettingLampsNumber = Convert.ToByte(SetLightNumber.SelectedIndex);
        }
        #endregion

        #region 生成电流数组
        private byte[] CalculateCurrentBuffer(string stringCurrentValue, int textboxNumber)
        {
            RefreshStringMessageLanguage();
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
                    if( MessageBox.Show(MessageboxContent37, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                                    if( MessageBox.Show(MessageboxContent38, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                                        if( MessageBox.Show(MessageboxContent39, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                                    if( MessageBox.Show(MessageboxContent38, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                                        if( MessageBox.Show(MessageboxContent39, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                                    if( MessageBox.Show(MessageboxContent38, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                                        if( MessageBox.Show(MessageboxContent39, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                            if( MessageBox.Show(MessageboxContent38, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                        if( MessageBox.Show(MessageboxContent40, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                    if( MessageBox.Show(MessageboxContent38, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                    if( MessageBox.Show(MessageboxContent39, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
            if(hardwareVersion1 == 8 && softwareNumber == 4)
            {
                InDeveloperModeChannelSelectContent = Convert.ToByte(InDeveloperModeChannelSelect.SelectedIndex);
                InDeveloperModeFlashFrequencyContent = Convert.ToByte(InDeveloperModeFlashFrequencySelect.SelectedItem);
                InDeveloperModeWaveformSelectContent = Convert.ToByte(InDeveloperModeWaveformSelect.SelectedIndex);

                InDeveloperModeSettingParameterCommand[0] = 0x55;
                InDeveloperModeSettingParameterCommand[1] = 0x02;
                InDeveloperModeSettingParameterCommand[2] = 0x11;
                InDeveloperModeSettingParameterCommand[3] = 0x58;
                InDeveloperModeSettingParameterCommand[4] = 0x12;

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
                InDeveloperModeSettingParameterCommand[22] = InDeveloperModeChannelSelectContent;
                InDeveloperModeSettingParameterCommand[23] = InDeveloperModeSettingBreakFlag;
                InDeveloperModeSettingParameterCommand[24] = InDeveloperModeSettingLampsNumber;
                InDeveloperModeSettingParameterCommand[25] = InDeveloperModeFlashFrequencyContent;
                InDeveloperModeSettingParameterCommand[26] = InDeveloperModeWaveformSelectContent;
                InDeveloperModeSettingParameterCommand[27] = CalculateCheckOutValue(InDeveloperModeSettingParameterCommand);
            }
            else
            {
                InDeveloperModeSettingParameterCommand[0] = 0x02;
                InDeveloperModeSettingParameterCommand[1] = 0x55;
                InDeveloperModeSettingParameterCommand[2] = 0x11;
                InDeveloperModeSettingParameterCommand[3] = 0x58;
                InDeveloperModeSettingParameterCommand[4] = 0x12;
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
            
        }
        #endregion

        #endregion
      
        #region 用户登录
        string userName;
        string password;
        private void LogIn_Click(object sender, RoutedEventArgs e)
        {
            RefreshStringMessageLanguage();
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
                    if(MessageBox.Show(MessageboxContent41, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error)==MessageBoxResult.OK)
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
                if( MessageBox.Show(MessageboxContent42, MessageboxHeader1, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
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
                AnswerSoftwareVersion.Text = "";
                AnswerHardwareVersion.Text = "";
                AnswerChannelSelect.Text = "";
                AnswerFlashFrequency.Text = "";
                AnswerWaveformSelect.Text = "";

                SelectApproachChenterlineLight.IsChecked = false;
                SelectApproachCrossbarLight.IsChecked = false;
                SelectApproachSideRowLight.IsChecked = false;
                SelectRWYThresholdWingBarLight.IsChecked = false;
                SelectRWYThresholdLight.IsChecked = false;
                SelectRWYEdgeLight.IsChecked = false;
                Select12inchesRWYEndLight.IsChecked = false;
                SelectRWYThresholdEndLight.IsChecked = false;
                SelectRWYCenterlineLight.IsChecked = false;
                Select12inchesRWYCenterlineLight.IsChecked = false;
                SelectRWYTouchdownZoneLight.IsChecked = false;
                Select8inchesRWYEndLight.IsChecked = false;
                SelectRapidExitTWYIndicatorLight.IsChecked = false;
                SelectCombinedRWYEdgeLight.IsChecked = false;
                SelectRWYGuardLight.IsChecked = false;
                SelectTWYCenterLight.IsChecked = false;
                SelectTWYCenterLight2P.IsChecked = false;
                SelectTWYStopBarLight.IsChecked = false;
                SelectIntermediateHoldingPositionLight.IsChecked = false;
                SelectTWYIntersectionsLight.IsChecked = false;
                SelectTWYEdgeLight.IsChecked = false;
                SelectElevatedApproachCenterlineLight.IsChecked = false;
                SelectElevatedApproachCrossbarLight.IsChecked = false;
                SelectElevatedApproachSideRowLight.IsChecked = false;
                SelectElevatedRWYEdgeLight.IsChecked = false;
                SelectElevatedRWYEndLight.IsChecked = false;
                SelectElevatedRWYThresholdLight.IsChecked = false;
                SelectElevatedRWYThresholdWingbarLight.IsChecked = false;
                SelectElevatedTWYStopBarLight.IsChecked = false;
                SelectAircraftStandManoeuvringGuidanceLight.IsChecked = false;


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
                SelectRELC12LEDRRB1P.IsChecked = false;
                SelectENDS12LEDR.IsChecked = false;
                SelectTAES12LLEDGR1P.IsChecked = false;
                SelectTAES12RLEDGR1P.IsChecked = false;
                SelectTAES12SLEDGR1P.IsChecked = false;
                SelectTAES12LLEDGRMR2P.IsChecked = false;
                SelectTAES12RLEDGRMR2P.IsChecked = false;
                SelectRCLS08LEDCB1P.IsChecked = false;
                SelectRCLS08LEDRB1P.IsChecked = false;
                SelectRCLS08LEDCC1P.IsChecked = false;
                SelectRCLS08LEDRC1P.IsChecked = false;
                SelectRCLS12LEDCCMR2P.IsChecked = false;
                SelectRCLS12LEDRCMR2P.IsChecked = false;
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
                SelectRELC12LEDCBC1P.IsChecked = false;
                SelectHRGS08LEDY.IsChecked = false;
                SelectTCLMS08SLEDGG1P.IsChecked = false;
                SelectTCLMS08SLEDGY1P.IsChecked = false;
                SelectTCLMS08SLEDYY1P.IsChecked = false;
                SelectTCLMS08SLEDYB1P.IsChecked = false;
                SelectTCLMS08SLEDGB1P.IsChecked = false;
                SelectTCLMS08CLEDGG1P.IsChecked = false;
                SelectTCLMS08CLEDGY1P.IsChecked = false;
                SelectTCLMS08CLEDYY1P.IsChecked = false;
                SelectTCLMS08CLEDYB1P.IsChecked = false;
                SelectTCLMS08CLEDGB1P.IsChecked = false;
                SelectTCLMS08SLEDGG2P.IsChecked = false;
                SelectTCLMS08SLEDGY2P.IsChecked = false;
                SelectTCLMS08SLEDYY2P.IsChecked = false;
                SelectTCLMS08SLEDYB2P.IsChecked = false;
                SelectTCLMS08SLEDGB2P.IsChecked = false;
                SelectTCLMS08CLEDGG2P.IsChecked = false;
                SelectTCLMS08CLEDGY2P.IsChecked = false;
                SelectTCLMS08CLEDYY2P.IsChecked = false;
                SelectTCLMS08CLEDYB2P.IsChecked = false;
                SelectTCLMS08CLEDGB2P.IsChecked = false;
                SelectSBLMS08SLEDR.IsChecked = false;
                SelectTPLMS08SLEDY.IsChecked = false;
                SelectTOIL08LLEDY.IsChecked = false;
                SelectTOEL08LEDB.IsChecked = false;
                SelectEULAPLEDCCenterline.IsChecked = false;
                SelectEULAPLEDCCrossbar.IsChecked = false;
                SelectEULSRLEDR.IsChecked = false;
                SelectEBLRELEDYC.IsChecked = false;
                SelectEBLRELEDCY.IsChecked = false;
                SelectEBLRELEDCC.IsChecked = false;
                SelectEBLRELEDCR.IsChecked = false;
                SelectEBLRELEDCR.IsChecked = false;
                SelectEBLRELEDRC.IsChecked = false;
                SelectEULTHLEDG.IsChecked = false;
                SelectEULTHWLEDG.IsChecked = false;
                SelectEULTHWLEDG.IsChecked = false;
                SelectEULEDLEDR.IsChecked = false;
                SelectEULSBLEDR.IsChecked = false;
                SelectASMG08LEDYM.IsChecked = false;


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
                Select12inchesRWYCenterlineLight.IsEnabled = true;
                SelectRWYTouchdownZoneLight.IsEnabled = true;
                Select8inchesRWYEndLight.IsEnabled = true;
                SelectRapidExitTWYIndicatorLight.IsEnabled = true;
                SelectCombinedRWYEdgeLight.IsEnabled = true;
                SelectRWYGuardLight.IsEnabled = true;
                SelectTWYCenterLight.IsEnabled = true;
                SelectTWYCenterLight2P.IsEnabled = true;
                SelectTWYStopBarLight.IsEnabled = true;
                SelectIntermediateHoldingPositionLight.IsEnabled = true;
                SelectTWYIntersectionsLight.IsEnabled = true;
                SelectTWYEdgeLight.IsEnabled = true;
                SelectElevatedApproachCenterlineLight.IsEnabled = true;
                SelectElevatedApproachCrossbarLight.IsEnabled = true;
                SelectElevatedApproachSideRowLight.IsEnabled = true;
                SelectElevatedRWYEdgeLight.IsEnabled = true;
                SelectElevatedRWYEndLight.IsEnabled = true;
                SelectElevatedRWYThresholdLight.IsEnabled = true;
                SelectElevatedRWYThresholdWingbarLight.IsEnabled = true;
                SelectElevatedTWYStopBarLight.IsEnabled = true;
                SelectAircraftStandManoeuvringGuidanceLight.IsEnabled = true;

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

        #region 设置软件语言
        LanguageHelper myLanguageHelper = new LanguageHelper();
        private void ConfirmLanguageSetting_Click(object sender, RoutedEventArgs e)
        {
            string languageFileName="";
            if (LanguageSelect.SelectedIndex==0)
            {
                languageFileName = "/Resources/Langs/zh-CN.xaml";
            }

            if (LanguageSelect.SelectedIndex == 1)
            {
                languageFileName = "/Resources/Langs/en-US.xaml";
            }

            myLanguageHelper.LoadLanguageFile(languageFileName);
        }

        public void RefreshStringMessageLanguage()
        {
            #region 后台代码，串口设置页面，中英文切换字符串
            LampInchesLabel1 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel1");
            LampInchesLabel2 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel2");
            LampInchesLabel3 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel3");
            LampInchesLabel4 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel4");
            LampInchesLabel5 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel5");
            LampInchesLabel6 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel6");
            LampInchesLabel7 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel7");
            LampInchesLabel8 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel8");
            LampInchesLabel9 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel9");
            LampInchesLabel10 = (string)System.Windows.Application.Current.FindResource("LangsLampInchesLabel10");


            #endregion

            #region 后台代码，工厂模式页面，中英文切换字符串
            AnswerHardwareVersion0 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion0");
            AnswerHardwareVersion1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion1");
            AnswerHardwareVersion2 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion2");
            AnswerHardwareVersion3 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion3");
            AnswerHardwareVersion4 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion4");
            AnswerHardwareVersion5 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion5");
            AnswerHardwareVersion6 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion6");
            AnswerHardwareVersion7 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion7");
            AnswerHardwareVersion8 = (string)System.Windows.Application.Current.FindResource("LangsAnswerHardwareVersion8");


            AnswerLampModel0 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel0");
            AnswerLampModel1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel1");
            AnswerLampModel2 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel2");
            AnswerLampModel3 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel3");
            AnswerLampModel4 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel4");
            AnswerLampModel5 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel5");
            AnswerLampModel6 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel6");
            AnswerLampModel7 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel7");
            AnswerLampModel8 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel8");
            AnswerLampModel9 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel9");
            AnswerLampModel10 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel10");
            AnswerLampModel11 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel11");
            AnswerLampModel12 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel12");
            AnswerLampModel13 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel13");
            AnswerLampModel14 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel14");
            AnswerLampModel15 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel15");
            AnswerLampModel16 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel16");
            AnswerLampModel17 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel17");
            AnswerLampModel18 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel18");
            AnswerLampModel19 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel19");
            AnswerLampModel20 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel20");
            AnswerLampModel21 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel21");
            AnswerLampModel22 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel22");
            AnswerLampModel23 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel23");
            AnswerLampModel24 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel24");
            AnswerLampModel25 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel25");
            AnswerLampModel26 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel26");
            AnswerLampModel27 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel27");
            AnswerLampModel28 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel28");
            AnswerLampModel29 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel29");
            AnswerLampModel30 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel30");
            AnswerLampModel31 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel31");
            AnswerLampModel32 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel32");
            AnswerLampModel33 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel33");
            AnswerLampModel34 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel34");
            AnswerLampModel35 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel35");
            AnswerLampModel36 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel36");
            AnswerLampModel37 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel37");
            AnswerLampModel38 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel38");
            AnswerLampModel39 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel39");
            AnswerLampModel40 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel40");
            AnswerLampModel41 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel41");
            AnswerLampModel42 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel42");
            AnswerLampModel43 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel43");
            AnswerLampModel44 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel44");
            AnswerLampModel45 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel45");
            AnswerLampModel46 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel46");
            AnswerLampModel47 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel47");
            AnswerLampModel48 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel48");
            AnswerLampModel49 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel49");
            AnswerLampModel50 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel50");
            AnswerLampModel51 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel51");
            AnswerLampModel52 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel52");
            AnswerLampModel53 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel53");
            AnswerLampModel54 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel54");
            AnswerLampModel55 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel55");
            AnswerLampModel56 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel56");
            AnswerLampModel57 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel57");
            AnswerLampModel58 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel58");
            AnswerLampModel59 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel59");
            AnswerLampModel60 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel60");
            AnswerLampModel61 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel61");
            AnswerLampModel62 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel62");
            AnswerLampModel63 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel63");
            AnswerLampModel64 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel64");
            AnswerLampModel65 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel65");
            AnswerLampModel66 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel66");
            AnswerLampModel67 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel67");
            AnswerLampModel69 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel69");
            AnswerLampModel70 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel70");
            AnswerLampModel71 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel71");
            AnswerLampModel68 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel68");
            AnswerLampModel72 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel72");
            AnswerLampModel73 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel73");
            AnswerLampModel74 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel74");
            AnswerLampModel75 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel75");
            AnswerLampModel76 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel76");
            AnswerLampModel77 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel77");
            AnswerLampModel78 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel78");
            AnswerLampModel79 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel79");
            AnswerLampModel80 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel80");
            AnswerLampModel81 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel81");
            AnswerLampModel82 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel82");
            AnswerLampModel83 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel83");
            AnswerLampModel84 = (string)System.Windows.Application.Current.FindResource("LangsAnswerLampModel84");


            AnswerOpenCircuit1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerOpenCircuit1");
            AnswerOpenCircuit2 = (string)System.Windows.Application.Current.FindResource("LangsAnswerOpenCircuit2");
            #endregion

            #region 后台代码，开发者模式页面，中英文切换字符串
            AnswerStatus1 = (string)System.Windows.Application.Current.FindResource("LangsAnswerStatus1");
            CreateExcel1 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel1");
            CreateExcel2 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel2");
            CreateExcel3 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel3");
            CreateExcel4 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel4");
            CreateExcel5 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel5");
            CreateExcel6 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel6");
            CreateExcel7 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel7");
            CreateExcel8 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel8");
            CreateExcel9 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel9");
            CreateExcel10 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel10");
            CreateExcel11 = (string)System.Windows.Application.Current.FindResource("LangsCreateExcel11");


            CreateTxt1 = (string)System.Windows.Application.Current.FindResource("LangsCreateTxt1");
            #endregion

            #region 后台代码，Messagebox，中英文切换字符串
            MessageboxHeader1 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxHeader1");
            MessageboxHeader2 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxHeader2");

            MessageboxContent1 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent1");
            MessageboxContent2 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent2");
            MessageboxContent3 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent3");
            MessageboxContent4 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent4");
            MessageboxContent5 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent5");
            MessageboxContent6 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent6");
            MessageboxContent7 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent7");
            MessageboxContent8 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent8");
            MessageboxContent9 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent9");
            MessageboxContent10 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent10");
            MessageboxContent11 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent11");
            MessageboxContent12 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent12");
            MessageboxContent13 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent13");
            MessageboxContent14 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent14");
            MessageboxContent15 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent15");
            MessageboxContent16 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent16");
            MessageboxContent17 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent17");
            MessageboxContent18 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent18");
            MessageboxContent19 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent19");
            MessageboxContent20 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent20");
            MessageboxContent21 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent21");
            MessageboxContent22 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent22");
            MessageboxContent23 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent23");
            MessageboxContent24 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent24");
            MessageboxContent25 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent25");
            MessageboxContent26 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent26");
            MessageboxContent27 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent27");
            MessageboxContent28 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent28");
            MessageboxContent29 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent29");
            MessageboxContent30 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent30");
            MessageboxContent31 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent31");
            MessageboxContent32 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent32");
            MessageboxContent33 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent33");
            MessageboxContent34 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent34");
            MessageboxContent35 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent35");
            MessageboxContent36 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent36");
            MessageboxContent37 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent37");
            MessageboxContent38 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent38");
            MessageboxContent39 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent39");
            MessageboxContent40 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent40");
            MessageboxContent41 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent41");
            MessageboxContent42 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent42");
            MessageboxContent43 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent43");
            MessageboxContent44 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent44");
            MessageboxContent45 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent45");
            MessageboxContent46 = (string)System.Windows.Application.Current.FindResource("LangsMessageboxContent46");

            #endregion
        }








        #endregion
        

        

        
    }
}
