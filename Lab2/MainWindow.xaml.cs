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
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Net;

namespace Lab2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public void StartWindow()
        {
            if (!System.IO.File.Exists(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
            {
                
                SearchButton.Visibility = Visibility.Visible;
                DownloadButton.Visibility = Visibility.Visible;
                UpdateData.Visibility = Visibility.Collapsed;
                
                MessageBox.Show("При загрузке программы необходимый файл с базой данных не был найден!\nПожалуйста, загрузите файл из сети Интернет,\nлибо выберите уже существующий на Вашем компьютере!", "Ошибка - Нет файла");
            }
            else
            {
                
                SearchButton.Visibility = Visibility.Collapsed;
                DownloadButton.Visibility = Visibility.Collapsed;
                UpdateData.Visibility = Visibility.Visible;
                if(CorrectFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
                ParsingFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                else
                {
                    MessageBox.Show("Выбранный Вами документ не является корректным!", "Ошибка распознавания");
                }
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
            openFileDialog.DefaultExt = "*.xlsx";
            openFileDialog.Title = "Выберите документ для загрузки данных";

            // Открываем окно диалога с пользователем.
            if (openFileDialog.ShowDialog() == true)
            {
                // Получаем расширение файла, выбранного пользователем.
                var extension = System.IO.Path.GetExtension(openFileDialog.FileName);
                
                if (extension.ToString() != ".xlsx")
                {
                    MessageBox.Show($"Вы выбрали файл с расширением {extension}\nНеобходим файл с расширением .xlsx!", "Ошибка - Неверный тип файла");
                }
                else
                {
                    if (CorrectFile(openFileDialog.FileName))
                    {
                        File.Copy(openFileDialog.FileName, AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                        StartWindow();
                        
                    }
                    else
                    {
                        MessageBox.Show("Выбранный Вами документ не является корректным!", "Ошибка распознавания");
                    }
                }
                
                
            }
        }
      private bool CorrectFile(string fileName)//Проверяет файл на корректность (кол-во колонок, в дальнейшем - названия колонок)
        {
            
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {

                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
                List<Bug> l = new List<Bug>();
                string str;
                int rCnt;
                int cCnt;
                int rw = 0;
                int cl = 0;


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;
                xlWorkBook.Close(true, fileName, null);
                xlApp.Quit();
                if (cl != 10)
                {
                    return false;
                }
                else
                {
                    return true;
                }
                            
                
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка распознавания документа: " + e.Message, "Ошибка распознавания");
                return false;
            }
           
                
                
                
            
        }
        private void ParsingFile(string fileName)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
                List<Bug> l = new List<Bug>();
                string str;
                int rCnt;
                int cCnt;
                int rw = 0;
                int cl = 0;

               
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                for (rCnt = 3; rCnt <= rw; rCnt++)
                {

                    //for (cCnt = 1; cCnt <= 3; cCnt++)
                    //{
                    //    str = ((range.Cells[rCnt, cCnt] as Excel.Range).Value2).ToString();
                    //    l1.Add(str);
                    //}
                    string id = ((range.Cells[rCnt, 1] as Excel.Range).Value2).ToString();
                    string des = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                    string source = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                    l.Add(new Bug(id, des, source));
                }

               
                WholeData.ItemsSource = l;
            }
            catch(Exception e)
            {
                MessageBox.Show("Ошибка парсинга: "+e.Message);
            }
            finally
            {
                xlWorkBook.Close(true, fileName, null);
                xlApp.Quit();
            }
        }

        private void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            WebClient wc = new WebClient();
            wc.DownloadFile(url, AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
            if (CorrectFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx")) ParsingFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
            StartWindow();
        }
    }
}
