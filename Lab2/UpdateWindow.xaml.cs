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
    /// Логика взаимодействия для UpdateWindow.xaml
    /// </summary>
    public partial class UpdateWindow : Window
    {
        public static int CurrentStartNumber { get; set; } = 0;
        public static List<Bug> l = new List<Bug>();
        public UpdateWindow()
        {
            InitializeComponent();
            
        }
        public UpdateWindow(List<Bug> list)
        {
            l = list;
            InitializeComponent();
            Info.Text += l.Count;
            UpdateData.AutoGenerateColumns = false;
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Идентификатор угрозы",
                Binding = new Binding("Id")
            });

            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Наименование угрозы",
                Binding = new Binding("Description")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Описание угрозы",
                Binding = new Binding("FullDescription")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Источник угрозы",
                Binding = new Binding("Source")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Объект воздействия",
                Binding = new Binding("ObjectDanger")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Объект воздействия",
                Binding = new Binding("ObjectDanger")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Нарушение конфиденциальности",
                Binding = new Binding("ConfDanger")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Нарушение целостности",
                Binding = new Binding("FullDanger")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Нарушение доступности",
                Binding = new Binding("AccessDanger")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Дата включения",
                Binding = new Binding("DateStartToString")
            });
            UpdateData.Columns.Add(new DataGridTextColumn
            {
                Header = "Дата последнего изменения",
                Binding = new Binding("DateUpdateToString")
            });
            UpdateData.Columns[0].Width = 150;
            UpdateData.IsReadOnly = true;
            
            Pagination(15);
        }

        private void GoBack_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Pagination(int count)
        {
            List<Bug> show = new List<Bug>();
            if(CurrentStartNumber == 0 && LeftButton != null && RightButton != null&&CurrentStartNumber+count>=l.Count) { LeftButton.IsEnabled = false; RightButton.IsEnabled = false; }
            else if (CurrentStartNumber == 0 && LeftButton != null && RightButton != null) { LeftButton.IsEnabled = false; RightButton.IsEnabled = true; }
            else if ((CurrentStartNumber + count >= l.Count) && RightButton != null && LeftButton != null) { RightButton.IsEnabled = false; LeftButton.IsEnabled = true; }
            else if (RightButton != null && LeftButton != null) { RightButton.IsEnabled = true; LeftButton.IsEnabled = true; }
            if (l.Count - (CurrentStartNumber + count) >= 1)
            {
                for (int i = CurrentStartNumber; i < CurrentStartNumber + count; i++)
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
                for (int i = CurrentStartNumber; i < l.Count; i++)
                {
                    show.Add(l[i]);
                }
                if (Diapazon != null)
                {
                    Diapazon.Content = $"{CurrentStartNumber + 1}-{l.Count}";
                }
            }
            UpdateData.ItemsSource = show;

        }
        private void LeftButton_Click(object sender, RoutedEventArgs e)
        {
            
                if (CurrentStartNumber - 15 > 0)
                {
                    CurrentStartNumber -= 15;
                }
                else CurrentStartNumber = 0;
                Pagination(15);
            

        }

        private void RightButton_Click(object sender, RoutedEventArgs e)
        {
          
           
              
                
                if (CurrentStartNumber + 15 >= l.Count)
                {
                    CurrentStartNumber -= 15;
                }
                else CurrentStartNumber += 15;
                Pagination(15);
            
        }
    }
}
