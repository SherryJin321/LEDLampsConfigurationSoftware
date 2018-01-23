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
using System.Windows.Shapes;

namespace LEDLampsConfigurationSoftware
{
    /// <summary>
    /// TotalTimeEnquiryWindows.xaml 的交互逻辑
    /// </summary>
    public partial class TotalTimeEnquiryWindows : Window
    {
        public uint breakDownCount { get; set; }
        public uint totalTime { get; set; }


        public TotalTimeEnquiryWindows()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if(breakDownCount==0xFFFFFFFF)
            {
                AnswerBreakDownCount.Text = 0.ToString();
            }
            else
            {
                AnswerBreakDownCount.Text = breakDownCount.ToString();
            }


            if (totalTime==0||totalTime==0xFFFFFFFF)
            {
                AnswerTotalTime.Text = 0.ToString();
            }
            else
            {
                AnswerTotalTime.Text = Convert.ToString(totalTime / 3600).PadLeft(2,'0') + ":" + Convert.ToString((totalTime % 3600) / 60).PadLeft(2, '0') + ":" + Convert.ToString((totalTime % 3600) % 60).PadLeft(2, '0');
            }
           
        }

        # region 关闭当前窗口
        private void CloseTotalTimeEnquiryWindows_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        #endregion
    }
}
