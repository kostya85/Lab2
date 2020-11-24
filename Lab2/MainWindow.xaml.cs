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
        public static int CurrentStartNumber { get; set; } = 0;
        public static List<Bug> l = new List<Bug>();
        public static int PaginationCountValue { get; set; } = 15;
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
                if (CorrectFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
                {
                    ParsingFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                    Pagination(PaginationCountValue);
                }
                else
                {
                    MessageBox.Show("Выбранный Вами документ не является корректным!", "Ошибка распознавания");
                }
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            WholeData.AutoGenerateColumns = false;
            WholeData.Columns.Add(new DataGridTextColumn
            {
                Header = "Идентификатор угрозы",
                Binding = new Binding("Id")
            });

            WholeData.Columns.Add(new DataGridTextColumn
            {
                Header = "Наименование угрозы",
                Binding = new Binding("Description")
            });
            WholeData.Columns[0].Width = 135;
            WholeData.IsReadOnly = true;
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
                    string id = ((range.Cells[rCnt, 1] as Excel.Range).Value).ToString();
                    string des = ((range.Cells[rCnt, 2] as Excel.Range).Value).ToString();
                    string fulldes = ((range.Cells[rCnt, 3] as Excel.Range).Value).ToString();
                    string source = ((range.Cells[rCnt, 4] as Excel.Range).Value).ToString();
                    string objdanger = ((range.Cells[rCnt, 5] as Excel.Range).Value).ToString();
                    string confdanger = ((range.Cells[rCnt, 6] as Excel.Range).Value).ToString();
                    string fulldanger = ((range.Cells[rCnt, 7] as Excel.Range).Value).ToString();
                    string accessdanger = ((range.Cells[rCnt, 8] as Excel.Range).Value).ToString();
                    DateTime datestart = ((range.Cells[rCnt, 9] as Excel.Range).Value);
                    DateTime dateupdate = ((range.Cells[rCnt, 10] as Excel.Range).Value);
                    
                    
                    l.Add(new Bug(id, des, fulldes, source, objdanger, confdanger, fulldanger, accessdanger, datestart,dateupdate));
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

        private void WholeData_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Bug path = WholeData.SelectedItem as Bug;

                string commonInfo = $"Идентификатор угрозы: {path.Id}\n\nНаименование угрозы: {path.Description}" +
                    $"\n\nОписание угрозы: {path.FullDescription}\nОбъект воздействия: {path.ObjectDanger}\n";
                string effects = $"Нарушение конфиденциальности: {path.ConfDanger}\n" +
                    $"Нарушение целостности: {path.FullDanger}\n" +
                    $"Нарушение досупности: {path.AccessDanger}";
                string extraInfo = $"Дата включения угрозы: {path.DateStart.ToString("dd.MM.yyyy")}\n\n" +
                    $"Дата последнего изменения данных: {path.DateUpdate.ToString("dd.MM.yyyy")}";
                BugInfo b = new BugInfo(commonInfo, effects, extraInfo);

                b.Show();
            }
            catch(Exception f)
            {
                MessageBox.Show($"Ошибка при открытии подробной информации об угрозе: \n{f.Message}\nПопробуйте еще раз...","Ошибка открытия");
            }
        }

        //private void ShowTypes(object sender, RoutedEventArgs e)
        //{
        //    string fileName = AppDomain.CurrentDomain.BaseDirectory + "data.xlsx";
        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            

        //        Excel.Worksheet xlWorkSheet;
        //        Excel.Range range;
        //        List<string> Types = new List<string>();
        //        Type str;
        //        int rCnt;
        //        int cCnt;
        //        int rw = 0;
        //        int cl = 0;


        //        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

        //        range = xlWorkSheet.UsedRange;
        //        rw = range.Rows.Count;
        //        cl = range.Columns.Count;




        //        for (cCnt = 1; cCnt <= cl; cCnt++)
        //        {
        //        if ((range.Cells[3, cCnt] as Excel.Range).Value is double) {

        //            Types.Add($"column {cCnt} is double\n");
        //        }
        //        else if((range.Cells[3, cCnt] as Excel.Range).Value is string)
        //        {
        //            Types.Add($"column {cCnt} is string\n");
        //        }
        //        else if ((range.Cells[3, cCnt] as Excel.Range).Value is bool)
        //        {
        //            Types.Add($"column {cCnt} is bool\n");
        //        }
        //        else if ((range.Cells[3, cCnt] as Excel.Range).Value is DateTime)
        //        {
        //            Types.Add($"column {cCnt} is DateTime\n");
        //        }
        //        else
        //        {
        //            Types.Add($"column {cCnt} is Type\n");
        //        }

        //    }




        //    xlWorkBook.Close(true, fileName, null);
        //    xlApp.Quit();

        //    string res = "";
        //    foreach (var s in Types) res += s;
        //    MessageBox.Show(res);
                
            
        //}

        private void SaveData_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel File|*.xlsx";
            saveFileDialog.Title = "Выберите путь для сохранения базы";
            if (saveFileDialog.ShowDialog() == true)
            {
                if (File.Exists(saveFileDialog.FileName))
                {
                    try
                    {
                        File.Delete(saveFileDialog.FileName);
                        File.Copy(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx", saveFileDialog.FileName);
                    }
                    catch(Exception k)
                    {
                        MessageBox.Show($"Возникла ошибка при сохранении файла: \n{k.Message}");
                    }
                }
                else File.Copy(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx", saveFileDialog.FileName);
            }

               
        }
        private void Pagination(int count)
        {
            List<Bug> show = new List<Bug>();
            if (CurrentStartNumber == 0 && LeftButton != null && RightButton != null) { LeftButton.IsEnabled = false; RightButton.IsEnabled = true; }
            else if ((CurrentStartNumber + count >= l.Count) && RightButton != null && LeftButton != null) { RightButton.IsEnabled = false; LeftButton.IsEnabled = true; }
            else if (RightButton != null && LeftButton != null) { RightButton.IsEnabled = true; LeftButton.IsEnabled = true; }
            if (l.Count - (CurrentStartNumber + count) >= 1)
            {
                for(int i = CurrentStartNumber; i < CurrentStartNumber+count; i++)
                {
                    show.Add(l[i]);
                }
                if (Diapazon != null)
                {
                    Diapazon.Content = $"{CurrentStartNumber + 1}-{CurrentStartNumber + count}";
                }
            }
            else
            {
                for(int i = CurrentStartNumber; i < l.Count; i++)
                {
                    show.Add(l[i]);
                }
                if (Diapazon != null)
                {
                    Diapazon.Content = $"{CurrentStartNumber + 1}-{l.Count}";
                }
            }
            WholeData.ItemsSource = show;
            
        }
        

      

        private void PaginationChoose_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (PaginationChoose.SelectedIndex == 0)
            {
                
                Pagination(15);
                
            }
            else
            {
              
                Pagination(20);
            }
        }

        private void LeftButton_Click(object sender, RoutedEventArgs e)
        {
            int count;
            if (PaginationChoose != null)
            {
                if (PaginationChoose.SelectedIndex == 0) count = 15;
                else count = 20;
                if (CurrentStartNumber - count > 0)
                {
                    CurrentStartNumber -= count;
                }
                else CurrentStartNumber = 0;
                Pagination(count);
            }
            
        }

        private void RightButton_Click(object sender, RoutedEventArgs e)
        {
            int count;
            if (PaginationChoose != null)
            {
                if (PaginationChoose.SelectedIndex == 0) count = 15;
                else count = 20;
                if (CurrentStartNumber + count >= l.Count)
                {
                    CurrentStartNumber -= count;
                }
                else CurrentStartNumber += count;
                Pagination(count);
            }
        }
    }
}
