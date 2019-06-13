using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using 降水侵蚀力计算.Handler;

namespace 降水侵蚀力计算
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 打开文件
        /// </summary>
        private async void Open_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel 文件|*.xlsx|Excel 2003|*.xls";
            var result = dialog.ShowDialog(this);
            if(result == true)
            {
                var fileName = dialog.FileName;
                var handler = new DefaultHandler(fileName);
                try
                {
                    this.message.Text = "正在计算, 请耐心等待...";
                    await Task.Factory.StartNew(handler.Handle);
                    this.message.Text = "计算完成, 选择文件重新计算.";
                    MessageBox.Show("计算完成", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    Process.Start(fileName);
                }
                catch(Exception ex)
                {
                    Debug.WriteLine(ex);
                    MessageBox.Show("发生异常", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
    }
}
