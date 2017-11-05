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
using Google.Cloud.Speech.V1;
using System.Threading;

namespace VCC_wpf
{
    /// <summary>
    /// Announcment_Window.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Announcment_Window : Window
    {
        static string tmp = "";
        public Announcment_Window()
        {
            InitializeComponent();
        }

        public void setLabel(String text)
        {
            TextLabel.Content += text;
        }

       
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Announcment ann = new Announcment();
            //Task<string> getResult = ann.Start(5);//5초동안

            //setLabel(getResult.ToString());

            return;
        }
    }
}
