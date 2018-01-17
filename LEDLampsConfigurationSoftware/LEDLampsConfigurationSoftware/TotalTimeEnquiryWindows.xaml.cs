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
        public int breakDownCount { get; set; }
        public int totalTime { get; set; }


        public TotalTimeEnquiryWindows()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            AnswerBreakDownCount.Text = breakDownCount.ToString();

            if(totalTime==0)
            {
                AnswerTotalTime.Text = totalTime.ToString();
            }
            else
            {
                AnswerTotalTime.Text = (totalTime / 3600).ToString() + ":" + ((totalTime % 3600) / 60).ToString() + ":" + ((totalTime % 3600) % 60).ToString();
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
