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
using System.Net.NetworkInformation;

namespace Lab2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// При открытии программы считывается файл update.txt. Если его нет, то он создается, в него записывается текущая дата.
    /// Если разница дат >=30 дней, то происходит автообновление базы, файл update.txt перезаписывается.
    /// При ручном обновлении файл update.txt также перезаписывается.
    /// </summary>
    public partial class MainWindow : Window
    {
        public static int CurrentStartNumber { get; set; } = 0;
        public static List<Bug> l = new List<Bug>();
        public static int PaginationCountValue { get; set; } = 15;
        public void StartWindow()//Данный метод отвечает за UI стартового окна
        {
            if (!System.IO.File.Exists(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
            {
                
                SearchButton.Visibility = Visibility.Visible;
                DownloadButton.Visibility = Visibility.Visible;
                UpdateData.Visibility = Visibility.Collapsed;
                ErrorText.Visibility = Visibility.Visible;
                SaveData.Visibility = Visibility.Collapsed;
                LeftButton.IsEnabled = false;
                RightButton.IsEnabled = false;
                //MessageBox.Show("При загрузке программы необходимый файл с базой данных не был найден!\nПожалуйста, загрузите файл из сети Интернет,\nлибо выберите уже существующий на Вашем компьютере!", "Ошибка - Нет файла");
            }
            else
            {
                
                
                if (CorrectFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
                {
                    ErrorText.Visibility = Visibility.Collapsed;
                    SearchButton.Visibility = Visibility.Collapsed;
                    DownloadButton.Visibility = Visibility.Collapsed;
                    UpdateData.Visibility = Visibility.Visible;
                    SaveData.Visibility = Visibility.Visible;
                    AutoUpdate();
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

        private void SearchButton_Click(object sender, RoutedEventArgs e)//Данный метод отвечает за обработку нажатия на кнопку "Выбрать файл" при отсутствии файла в корневой папке программы
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
                    
                }
                
                
            }
        }
      private bool CorrectFile(string fileName)//Проверяет файл на корректность (кол-во колонок, тип колонок)
        {
            
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {

                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
                List<Bug> l = new List<Bug>();
                
                
                
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
                    throw new Exception("Кол-во колонок неверное!");
                    
                }
                else
                {
                    int cCnt;
                    List<string> Types = new List<string>();
                    for (cCnt = 1; cCnt <= cl; cCnt++)
                    {
                        if ((range.Cells[3, cCnt] as Excel.Range).Value is double)
                        {

                            Types.Add($"double");
                        }
                        else if ((range.Cells[3, cCnt] as Excel.Range).Value is string)
                        {
                            Types.Add($"string");
                        }
                        else if ((range.Cells[3, cCnt] as Excel.Range).Value is bool)
                        {
                            Types.Add($"bool");
                        }
                        else if ((range.Cells[3, cCnt] as Excel.Range).Value is DateTime)
                        {
                            Types.Add($"DateTime");
                        }
                        else
                        {
                            Types.Add($"Type");
                        }
                    }
                    List<string> Correct = new List<string>() { "double", "string", "string", "string", "string", "double", "double", "double", "DateTime", "DateTime" };
                    if(Correct.Equals(Types))return true;
                    else throw new Exception("Типы колонок неверные!");
                }
                            
                
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка распознавания документа: " + e.Message, "Ошибка распознавания");
                return false;
            }
           
                
                
                
            
        }
        private void ParsingFile(string fileName)//Данный метод отвечает за парсинг файла путем создания экземпляров класса Bug и их сессионного хранения в статическом списке l
        {
            if (l.Count > 0) l = new List<Bug>();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
                
                
                int rCnt;
                
                int rw = 0;
                int cl = 0;

               
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                for (rCnt = 3; rCnt <= rw; rCnt++)
                {

                    
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

               
               // WholeData.ItemsSource = l;
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

        private void DownloadButton_Click(object sender, RoutedEventArgs e)//Данный метод отвечает за обработку нажатия на кнопку "Загрузить из Интернет" при отсутствии файла в корневой папке программы
        {
            try
            {
                if (!CheckForInternetConnection()) { throw new Exception(); }
                else
                {
                    string url = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
                    WebClient wc = new WebClient();
                    wc.DownloadFile(url, AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                    if (CorrectFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
                    {
                        ParsingFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                        StartWindow();
                    }
                }
            }
            catch (Exception exep)
            {
                MessageBox.Show("Ошибка загрузки файла, возможно Вы не подключены к Интернет!", "Ошибка загрузки");
            }
        }

        private void WholeData_MouseUp(object sender, MouseButtonEventArgs e)//Данный метод отвечает за отображение полной информации об угрозе при нажатии на нее в DataGrid
        {
            try
            {
                Bug path = WholeData.SelectedItem as Bug;
                if (path != null)
                {
                    string commonInfo = $"Идентификатор угрозы: {path.Id}\n\nНаименование угрозы: {path.Description}" +
                        $"\n\nОписание угрозы: {path.FullDescription}\n\nОбъект воздействия: {path.ObjectDanger}\n";
                    string effects = $"Нарушение конфиденциальности: {path.ConfDanger}\n" +
                        $"Нарушение целостности: {path.FullDanger}\n" +
                        $"Нарушение досупности: {path.AccessDanger}";
                    string extraInfo = $"Дата включения угрозы: {path.DateStartToString}\n\n" +
                        $"Дата последнего изменения данных: {path.DateUpdateToString}";
                    BugInfo b = new BugInfo(commonInfo, effects, extraInfo);

                    b.Show();
                }
            }
            catch(Exception f)
            {
                MessageBox.Show($"Ошибка при открытии подробной информации об угрозе: \n{f.Message}\nПопробуйте еще раз...","Ошибка открытия");
            }
        }
        //Это я для себя определял, какой тип имеет каждая колонка
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

        private void SaveData_Click(object sender, RoutedEventArgs e)//Метод отвечает за сохранения файла по запросу пользователя
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel File|*.xlsx";
            saveFileDialog.Title = "Выберите путь для сохранения базы";
            if (saveFileDialog.ShowDialog() == true)
            {
                if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx")) MessageBox.Show("Файла не существует в корневом каталоге!", "Ошибка сохранения");
                else { 
                if (File.Exists(saveFileDialog.FileName))
                {
                    try
                    {
                        File.Delete(saveFileDialog.FileName);
                        File.Copy(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx", saveFileDialog.FileName);
                    }
                    catch (Exception k)
                    {
                        MessageBox.Show($"Возникла ошибка при сохранении файла: \n{k.Message}");
                    }
                }
                else File.Copy(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx", saveFileDialog.FileName);
                }
            }

               
        }
        private void Pagination(int count)//Метод отвечает за пагинацию
        {
            List<Bug> show = new List<Bug>();
            if (CurrentStartNumber == 0 && LeftButton != null && RightButton != null && CurrentStartNumber + count >= l.Count) { LeftButton.IsEnabled = false; RightButton.IsEnabled = false; }
            else if (CurrentStartNumber == 0 && LeftButton != null && RightButton != null) { LeftButton.IsEnabled = false; RightButton.IsEnabled = true; }
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
                PaginationCountValue = 15;
                Pagination(15);
                
            }
            else
            {
                PaginationCountValue = 20;
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
        private void Update()
        {
            List<Bug> result = new List<Bug>();
            List<Bug> before = l;
            try
            {
                if (!CheckForInternetConnection()) { throw new Exception(); }
                else
                {
                    string url = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
                    WebClient wc = new WebClient();

                    if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx")) File.Delete(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                    wc.DownloadFile(url, AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                    if (CorrectFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx"))
                    {
                        ParsingFile(AppDomain.CurrentDomain.BaseDirectory + "data.xlsx");
                        List<Bug> after = l;

                        for (int i = 0; i < before.Count; i++)
                        {
                            bool isExist = false;
                            bool hasChanges = false;
                            for (int j = 0; j < after.Count; j++)
                            {
                                if (after[j].Id == before[i].Id)
                                {
                                    isExist = true;
                                    if (before[i].Description != after[j].Description) { before[i].Description = "[БЫЛО]\n" + before[i].Description + "\n[СТАЛО]\n" + after[j].Description; hasChanges = true; }
                                    if (before[i].FullDescription != after[j].FullDescription) { before[i].FullDescription = "[БЫЛО]\n" + before[i].FullDescription + "\n[СТАЛО]\n" + after[j].FullDescription; hasChanges = true; }
                                    if (before[i].Source != after[j].Source) { before[i].Source = "[БЫЛО]\n" + before[i].Source + "\n[СТАЛО]\n" + after[j].Source; hasChanges = true; }
                                    if (before[i].ObjectDanger != after[j].ObjectDanger) { before[i].ObjectDanger = "[БЫЛО]\n" + before[i].ObjectDanger + "\n[СТАЛО]\n" + after[j].ObjectDanger; hasChanges = true; }
                                    if (before[i].AccessDanger != after[j].AccessDanger) { before[i].AccessDanger = "[БЫЛО]\n" + before[i].AccessDanger + "\n[СТАЛО]\n" + after[j].AccessDanger; hasChanges = true; }
                                    if (before[i].FullDanger != after[j].FullDanger) { before[i].FullDanger = "[БЫЛО]\n" + before[i].FullDanger + "\n[СТАЛО]\n" + after[j].FullDanger; hasChanges = true; }
                                    if (before[i].ConfDanger != after[j].ConfDanger) { before[i].ConfDanger = "[БЫЛО]\n" + before[i].ConfDanger + "\n[СТАЛО]\n" + after[j].ConfDanger; hasChanges = true; }
                                    //if (before[i].DateStart != after[j].DateStart) before[i].ConfDanger += "[БЫЛО]\n" + before[i].ConfDanger + "\n[СТАЛО]\n" + after[j].ConfDanger;
                                    before[i].DateUpdate = after[j].DateUpdate;
                                    if (hasChanges) result.Add(before[i]);
                                }
                            }
                            if (!isExist)
                            {
                                before[i].Id = "[УДАЛЕНА ЗАПИСЬ]\n" + before[i].Id;
                                result.Add(before[i]);
                            }
                        }
                        for (int i = 0; i < after.Count; i++)
                        {
                            bool isExist = false;
                            for (int j = 0; j < before.Count; j++)
                            {
                                if (after[i].Id == before[j].Id)
                                {
                                    isExist = true;
                                }
                            }
                            if (!isExist)
                            {
                                after[i].Id = "[ДОБАВЛЕНА ЗАПИСЬ]\n" + after[i].Id;
                                result.Add(after[i]);
                            }
                        }


                        UpdateWindow update = new UpdateWindow(result);
                       
                        File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "update.txt", DateTime.Now.ToString());
                        
                        update.Show();
                        StartWindow();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки файла, возможно Вы не подключены к Интернет!", "Ошибка загрузки");
            }
        }
        private void UpdateData_Click(object sender, RoutedEventArgs e)
        {

            Update();
            
        }
        public bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }
        private void AutoUpdate()
        {
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "update.txt"))
            {
                DateTime lastDate = DateTime.Parse(File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "update.txt"));
                DateTime now = DateTime.Now;
                if ((now - lastDate).TotalDays >= 30)
                {
                    File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "update.txt", DateTime.Now.ToString());
                    Update();
                }
            }
            else
            {
                //File.Create(AppDomain.CurrentDomain.BaseDirectory + "update.txt");
                File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "update.txt",DateTime.Now.ToString());
            }
        }
    }
}
