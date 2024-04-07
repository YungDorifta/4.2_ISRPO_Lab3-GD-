using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace Lab3WPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection <ProductInfo> Data { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            Data = new ObservableCollection<ProductInfo>();
            TableBox.ItemsSource = Data;
            Data.Add(new ProductInfo() { Id = 1, Product = "Каша", Count = 3, Price = 30 });
        }

        private void FormDocument(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog();
            bool? result = dialog.ShowDialog();
            if (result == true)
            {
                string filename = dialog.FileName;
                int ZakazNum = 0;
                Int32.TryParse(NumBox.Text, out ZakazNum);
                DateTime ZakazDate = DateBox.DisplayDate;
                if (!String.IsNullOrEmpty(DateBox.Text)) DateTime.TryParse(DateBox.Text, out ZakazDate);

                foreach (ProductInfo item in Data)
                {
                    if (item.Id == 0) item.Id = Data.IndexOf(item) + 1;
                }

                Docwork.SetWord(filename, PostavBox.Text, PokupBox.Text, Docwork.findSummary(Data), ZakazNum, ZakazDate, Data);
            }
        }
        
        private void TableBox_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            var AllBinded = TableBox.Items.OfType<Object>().ToList();
            SumBox.Content = "Итого: " + Docwork.findSummary(Data).ToString();
        }
    }
}
