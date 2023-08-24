using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Net.Http;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.IO;


namespace Parser__rubilnik.ru_
{
    public partial class Form1 : Form
    {
        string excelFilePath = null;

        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Вставьте ссылку на страницу с товаром");
                return;
            }

            await ProcessParsingAsync(textBox1.Text, excelFilePath);
        }

        private async Task ProcessParsingAsync(string url, string excelFilePath)
        {
            try
            {
                Disable();
                await ParseAndSaveAsync(url, excelFilePath);
                Enable();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}");
            }
        }

        private void Enable()
        {
            textBox1.Enabled = true;
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void Disable()
        {
            textBox1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
        }

        private async Task ParseAndSaveAsync(string url, string excelFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var web = new HtmlWeb();
                    var htmlDocument = await web.LoadFromWebAsync(url);

                    var blockNodes = htmlDocument.DocumentNode.SelectNodes("//div[@class='inner_wrap TYPE_1']");
                    if (blockNodes != null)
                    {
                        foreach (var blockNode in blockNodes)
                        {
                            var titleNode = blockNode.SelectSingleNode(".//div[@class='item-title']");
                            var priceNode = blockNode.SelectSingleNode(".//span[@class='price_value']");

                            if (titleNode != null && priceNode != null)
                            {
                                string title = titleNode.InnerText.Trim();
                                string price = priceNode.InnerText.Trim();

                                var row = worksheet.Cells["A:A"].FirstOrDefault(cell => cell.Value?.ToString() == title)?.Start.Row;

                                if (row.HasValue)
                                {
                                    worksheet.Cells[row.Value, 3].Value = price;
                                }
                            }
                        }

                        await package.SaveAsync();
                    }
                }
                MessageBox.Show("Парсинг завершен успешно.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " +  ex.Message + "\n\n\n\n P.s. если у вас открыт файл Excel, который вы пытаетесь обновить, то необходимо его закрыть, чтобы избежать неправильной работы программы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = openFileDialog.FileName;
            }

            if (excelFilePath != null)
            {
                button1.Enabled = true;
            }
        }


    }
}


//private void button1_Click(object sender, EventArgs e)
//{
//    string url = textBox1.Text;
//    parse(url);
//}
//
//async private void parse(string url)
//{
//    // Создание экземпляра HttpClient для выполнения HTTP-запросов
//    using (var httpClient = new HttpClient())
//    {
//        // Выполнение GET-запроса и получение содержимого страницы
//        string html = await httpClient.GetStringAsync(url);
//
//        // Создание экземпляра HtmlDocument для обработки HTML
//        var htmlDocument = new HtmlAgilityPack.HtmlDocument();
//        htmlDocument.LoadHtml(html);
//
//        // Используем XPath для нахождения блоков класса "inner_wrap TYPE_1"
//        var blockNodes = htmlDocument.DocumentNode.SelectNodes("//div[@class='inner_wrap TYPE_1']");
//        if (blockNodes != null)
//        {
//            foreach (var blockNode in blockNodes)
//            {
//                // Находим элементы внутри блока
//                var titleNode = blockNode.SelectSingleNode(".//div[@class='item-title']");
//                var valueNode = blockNode.SelectSingleNode(".//span[@class='value font_sxs']");
//                var priceNode = blockNode.SelectSingleNode(".//span[@class='price_value']");
//
//                Console.WriteLine();
//
//                // Извлекаем текст из элементов и выводим на экран
//                if (titleNode != null)
//                {
//                    Console.WriteLine("Title: " + titleNode.InnerText.Trim());
//                }
//                if (valueNode != null)
//                {
//                    Console.WriteLine("Value: " + valueNode.InnerText.Trim());
//                }
//                if (priceNode != null)
//                {
//                    Console.WriteLine("Price: " + priceNode.InnerText.Trim());
//                }
//
//
//            }
//        }
//    }
//}