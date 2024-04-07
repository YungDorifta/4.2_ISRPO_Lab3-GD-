using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;

//Сделано: все параметры накладной передаются в метод
//Сделано: выбор папки и имени файла
//Сделано: дата при отсутствии выставляет сегодняшнюю
//Сделано: ввод данных о продуктах вручную



namespace Lab3WPF
{
    public class Docwork
    {
        public static ObservableCollection<ProductInfo> defaultShop = new ObservableCollection<ProductInfo>(){
            new ProductInfo()
        {
            Id = 1,
                Product = "Апельсины",
                Count = 50,
                Price = 120.5,
              },
              new ProductInfo()
        {
            Id = 1,
                Product = "Бананы",
                Count = 130,
                Price = 100.5,
              },
              new ProductInfo()
        {
            Id = 1,
                Product = "Помидоры",
                Count = 120,
                Price = 150.5,
              },
              new ProductInfo()
        {
            Id = 1,
                Product = "Огурцы",
                Count = 150,
                Price = 140.5,
              }
    };

        public static double findSummary(ObservableCollection<ProductInfo> Shop)
        {
            double sum = 0;
            foreach (ProductInfo item in Shop)
            {
                sum += item.Price * item.Count;
            }
            return sum;
        }

        public static void SetWord(string argFilePath, string argPostavshic, string argPokupatel, double argSummary, int argZakazNum, DateTime argData, ObservableCollection<ProductInfo> argProducts)
        {
            // создаем приложение ворд
            Word.Application winword = new Word.Application();
            //winword.Visible = true;

            // добавляем документ
            Word.Document document = winword.Documents.Add();

            // добавляем параграф с номером накладной и выбранной датой
            Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
            //DateTime? selectDate = DateTime.Now;
            DateTime? selectDate = argData;
            string invoiceNumber = argZakazNum.ToString();

            // добавление даты при необходимости
            if (selectDate != null)
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber, " от ", selectDate.Value.ToString("dd.MM.yyyy"));
            else
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber);

            invoicePar.Range.Font.Name = "Times new roman";
            invoicePar.Range.Font.Size = 14;
            invoicePar.Range.InsertParagraphAfter();

            // добавляем параграф с поставщиком
            string PurchasertxtBox = argPostavshic;
            Word.Paragraph providerPar = document.Content.Paragraphs.Add();
            providerPar.Range.Text = string.Concat("Поставщик: ", PurchasertxtBox);
            providerPar.Range.Font.Name = "Times new roman";
            providerPar.Range.Font.Size = 14;
            providerPar.Range.InsertParagraphAfter();

            // добавляем параграф с потребителем
            Word.Paragraph customerPar = document.Content.Paragraphs.Add();
            string ProvidertxtBox = argPokupatel;
            customerPar.Range.Text = "Покупатель: " + ProvidertxtBox;
            customerPar.Range.Font.Name = "Times new roman";
            customerPar.Range.Font.Size = 14;
            customerPar.Range.InsertParagraphAfter();
            
            // кол-во строк
            int nRows = argProducts.Count;

            // создание таблицы вордфайла
            Word.Table myTable = document.Tables.Add(customerPar.Range, nRows, 5);
            myTable.Borders.Enable = 1;

            // добавляем данные из таблицы в ворд
            double summary = 0;
            for (int i = 1; i < argProducts.Count + 1; i++)
            {
                var dataRow = myTable.Rows[i].Cells;
                dataRow[1].Range.Text = argProducts[i - 1].Id.ToString();
                dataRow[2].Range.Text = argProducts[i - 1].Product;
                dataRow[3].Range.Text = argProducts[i - 1].Count.ToString();
                dataRow[4].Range.Text = argProducts[i - 1].Price.ToString();
                dataRow[5].Range.Text = (argProducts[i - 1].Price * argProducts[i - 1].Count).ToString();
                summary += argProducts[i - 1].Price * argProducts[i - 1].Count;
            }

            Word.Paragraph summaryPar = document.Content.Paragraphs.Add();
            summaryPar.Range.Text = string.Concat("Итого: ", summary.ToString());
            summaryPar.Range.Font.Name = "Times new roman";
            summaryPar.Range.Font.Size = 14;
            summaryPar.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            summaryPar.Range.InsertParagraphAfter();


            // указываем в какой файл сохранить
            // и места где его нужно сохранить
            object filename = @"H:\ИСРПО\Лабы\2 семестр\!-Лаб3(GD)\WordFiles\wordDocument.doc";
            if (!String.IsNullOrEmpty(argFilePath)) filename = argFilePath;
            if (!filename.ToString().Contains(".doc")) filename += ".doc";
            document.SaveAs(filename);
            document.Close();
            winword.Quit();
        }


    }
}
