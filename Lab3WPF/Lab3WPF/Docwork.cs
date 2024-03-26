using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Lab3WPF
{
    public class Docwork
    {
        public static void SetWord(string argFileName)
        {
            // создаем приложение ворд
            Word.Application winword = new Word.Application();
            //winword.Visible = true;

            // добавляем документ
            Word.Document document = winword.Documents.Add();

            // добавляем параграф с номером накладной и выбранной датой
            Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
            DateTime? selectDate = DateTime.Now;
            string invoiceNumber = "12";

            // добавление даты при необходимости
            if (selectDate != null)
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber, " от ", selectDate.Value.ToString("dd.MM.yyyy"));
            else
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber);

            invoicePar.Range.Font.Name = "Times new roman";
            invoicePar.Range.Font.Size = 14;
            invoicePar.Range.InsertParagraphAfter();

            // добавляем параграф с поставщиком
            string PurchasertxtBox = "ООО Фирма №1";
            Word.Paragraph providerPar = document.Content.Paragraphs.Add();
            providerPar.Range.Text = string.Concat("Поставщик: ", PurchasertxtBox);
            providerPar.Range.Font.Name = "Times new roman";
            providerPar.Range.Font.Size = 14;
            providerPar.Range.InsertParagraphAfter();

            // добавляем параграф с потребителем
            Word.Paragraph customerPar = document.Content.Paragraphs.Add();
            string ProvidertxtBox = "ООО Фирма №2";
            customerPar.Range.Text = "Покупатель: " + ProvidertxtBox;
            customerPar.Range.Font.Name = "Times new roman";
            customerPar.Range.Font.Size = 14;
            customerPar.Range.InsertParagraphAfter();

            // формируем таблицу
            // количество колонок - 4
            // количество строк - nRows
            List<Commodity> Shop = new List<Commodity>()
            {
              new Commodity()
              {
                Id = 1,
                Product = "Апельсины",
                Count = 50,
                Price = 120.5,
              },
              new Commodity()
              {
                Id = 1,
                Product = "Бананы",
                Count = 130,
                Price = 100.5,
              },
              new Commodity()
              {
                Id = 1,
                Product = "Помидоры",
                Count = 120,
                Price = 150.5,
              },
              new Commodity()
              {
                Id = 1,
                Product = "Огурцы",
                Count = 150,
                Price = 140.5,
              }
            };

            // кол-во строк
            int nRows = Shop.Count;

            // создание таблицы вордфайла
            Word.Table myTable = document.Tables.Add(customerPar.Range, nRows, 5);
            myTable.Borders.Enable = 1;

            // добавляем данные из таблицы в ворд
            for (int i = 1; i < Shop.Count + 1; i++)
            {
                var dataRow = myTable.Rows[i].Cells;
                dataRow[1].Range.Text = Shop[i - 1].Id.ToString();
                dataRow[2].Range.Text = Shop[i - 1].Product;
                dataRow[3].Range.Text = Shop[i - 1].Count.ToString();
                dataRow[4].Range.Text = Shop[i - 1].Price.ToString();
                dataRow[5].Range.Text = (Shop[i - 1].Price * Shop[i - 1].Count).ToString();
            }

            // указываем в какой файл сохранить
            // TODO - добавьте возможность выбора названия файла
            // и места где его нужно сохранить
            object filename = @"H:\ИСРПО\Лабы\2 семестр\!-Лаб3(GD)\WordFiles\";
            if (!String.IsNullOrEmpty(argFileName)) filename += argFileName;
            else filename = @"H:\ИСРПО\Лабы\2 семестр\!-Лаб3(GD)\WordFiles\wordDocument.doc";
            document.SaveAs(filename);
            document.Close();
            winword.Quit();
        }
    }
}
