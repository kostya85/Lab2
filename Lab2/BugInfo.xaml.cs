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

namespace Lab2
{
    /// <summary>
    /// Логика взаимодействия для BugInfo.xaml
    /// </summary>
    public partial class BugInfo : Window
    {
        public BugInfo()
        {
            InitializeComponent();
        }
        public BugInfo(string commonInfo, string effects, string extraInfo)
        {
            InitializeComponent();
            CommonInfo.Text = commonInfo;
            Effects.Content = effects;
            ExtraInfo.Content = extraInfo;
        }

        private void GoBack_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
