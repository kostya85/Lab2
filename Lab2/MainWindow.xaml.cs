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
using System.IO;
using Microsoft.Win32;

namespace Lab2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public void StartWindow()
        {
            if (!System.IO.File.Exists("data.xlsx"))
            {
                UserFilePath.Visibility = Visibility.Visible;
                SearchButton.Visibility = Visibility.Visible;
                DownloadButton.Visibility = Visibility.Visible;
                UpgradeData.Visibility = Visibility.Collapsed;
                MessageBox.Show("При загрузке программы необходимый файл с базой данных не был найден!\nПожалуйста, загрузите файл из сети Интернет,\nлибо выберите уже существующий на Вашем компьютере!");
            }
            else
            {
                UserFilePath.Visibility = Visibility.Collapsed;
                SearchButton.Visibility = Visibility.Collapsed;
                DownloadButton.Visibility = Visibility.Collapsed;
                UpgradeData.Visibility = Visibility.Visible;
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            StartWindow();
            
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            // Открываем окно диалога с пользователем.
            if (openFileDialog.ShowDialog() == true)
            {
                // Получаем расширение файла, выбранного пользователем.
                var extension = System.IO.Path.GetExtension(openFileDialog.FileName);
                UserFilePath.Text = openFileDialog.FileName;
                if (extension.ToString() != ".xlsx")
                {
                    MessageBox.Show($"Вы выбрали файл с расширением {extension}\nНеобходим файл с расширением .xlsx!");
                }
                else
                {
                    var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                }
                
                
            }
        }
    }
}
