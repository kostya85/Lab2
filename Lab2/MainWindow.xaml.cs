﻿using System;
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
                    ParsingFile(openFileDialog.FileName);
                }
                
                
            }
        }
        private void ParsingFile(string fileName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            List<Bug> l = new List<Bug>();
            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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
                string source= ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                l.Add(new Bug(id, des, source));
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            WholeData.ItemsSource = l;
        }
    }
}
