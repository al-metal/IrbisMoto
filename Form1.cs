using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using web;

namespace IrbisMoto
{
    public partial class Form1 : Form
    {
        web.WebRequest webRequest = new web.WebRequest();
        string otv = null;
        int deleteTovar = 0;
        int editPrice = 0;
        WebClient webClient = new WebClient();

        public Form1()
        {
            InitializeComponent();
            if (!Directory.Exists("files"))
            {
                Directory.CreateDirectory("files");
            }
            if (!Directory.Exists("pic"))
            {
                Directory.CreateDirectory("pic");
            }

            if (!File.Exists("files\\miniText.txt"))
            {
                File.Create("files\\miniText.txt");
            }

            if (!File.Exists("files\\fullText.txt"))
            {
                File.Create("files\\fullText.txt");
            }

            if (!File.Exists("files\\title.txt"))
            {
                File.Create("files\\title.txt");
            }

            if (!File.Exists("files\\description.txt"))
            {
                File.Create("files\\description.txt");
            }

            if (!File.Exists("files\\keywords.txt"))
            {
                File.Create("files\\keywords.txt");
            }
            StreamReader altText = new StreamReader("files\\miniText.txt", Encoding.GetEncoding("windows-1251"));
            while (!altText.EndOfStream)
            {
                string str = altText.ReadLine();
                rtbMiniText.AppendText(str + "\n");
            }
            altText.Close();

            altText = new StreamReader("files\\fullText.txt", Encoding.GetEncoding("windows-1251"));
            while (!altText.EndOfStream)
            {
                string str = altText.ReadLine();
                rtbFullText.AppendText(str + "\n");
            }
            altText.Close();

            altText = new StreamReader("files\\title.txt", Encoding.GetEncoding("windows-1251"));
            while (!altText.EndOfStream)
            {
                string str = altText.ReadLine();
                tbTitle.AppendText(str + "\n");
            }
            altText.Close();

            altText = new StreamReader("files\\description.txt", Encoding.GetEncoding("windows-1251"));
            while (!altText.EndOfStream)
            {
                string str = altText.ReadLine();
                tbDescription.AppendText(str + "\n");
            }
            altText.Close();

            altText = new StreamReader("files\\keywords.txt", Encoding.GetEncoding("windows-1251"));
            while (!altText.EndOfStream)
            {
                string str = altText.ReadLine();
                tbKeywords.AppendText(str + "\n");
            }
            altText.Close();
        }

        private void btnSaveTemplates_Click(object sender, EventArgs e)
        {
            int count = 0;
            StreamWriter writers = new StreamWriter("files\\miniText.txt", false, Encoding.GetEncoding(1251));
            count = rtbMiniText.Lines.Length;
            for (int i = 0; rtbMiniText.Lines.Length > i; i++)
            {
                if (count - 1 == i)
                {
                    if (rtbFullText.Lines[i] == "")
                        break;
                }
                writers.WriteLine(rtbMiniText.Lines[i].ToString());
            }
            writers.Close();

            writers = new StreamWriter("files\\fullText.txt", false, Encoding.GetEncoding(1251));
            count = rtbFullText.Lines.Length;
            for (int i = 0; count > i; i++)
            {
                if (count - 1 == i)
                {
                    if (rtbFullText.Lines[i] == "")
                        break;
                }
                writers.WriteLine(rtbFullText.Lines[i].ToString());
            }
            writers.Close();

            writers = new StreamWriter("files\\title.txt", false, Encoding.GetEncoding(1251));
            writers.WriteLine(tbTitle.Lines[0]);
            writers.Close();

            writers = new StreamWriter("files\\description.txt", false, Encoding.GetEncoding(1251));
            writers.WriteLine(tbDescription.Lines[0]);
            writers.Close();

            writers = new StreamWriter("files\\keywords.txt", false, Encoding.GetEncoding(1251));
            writers.WriteLine(tbKeywords.Lines[0]);
            writers.Close();

            MessageBox.Show("Сохранено");
        }

        private void btnActual_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo("Прайс.xlsx");
            ExcelPackage p = new ExcelPackage(file);
            ExcelWorksheet w = p.Workbook.Worksheets[3];
            int q = w.Dimension.Rows;

            //for (int i = 8; q > i; i++)
            //{
            //    if (w.Cells[i, 1].Value == null)
            //        break;

            //    double articl = (double)w.Cells[i, 1].Value;
            //    double quantity = (double)w.Cells[i, 9].Value;
            //    double priceIrbisDiler = (double)w.Cells[i, 6].Value;
            //    double actualPrice = Price(priceIrbisDiler);
            //    string action = (string)w.Cells[i, 14].Value;

            //    ExcelRange er = w.Cells[i, 2];
            //    if (er.Hyperlink != null)
            //    {
            //        string urlImg = er.Hyperlink.ToString();
            //        try
            //        {
            //            webClient.DownloadFile(urlImg, "pic\\" + articl + ".jpg");
            //        }
            //        catch
            //        {

            //        }

            //    }

            //    if (action != "")
            //    {
            //        switch (action)
            //        {
            //            case "ЛУЧШАЯ ЦЕНА!":
            //                action = "&markers[3]=1";
            //                break;
            //            case "Новое поступление":
            //                action = "&markers[1]=1";
            //                break;
            //            case "Новое постуление":
            //                action = "&markers[1]=1";
            //                break;
            //            default:
            //                action = "";
            //                break;
            //        }
            //    }

            //    if (quantity == 0)
            //    {
            //        otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
            //        string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
            //        if (urlTovar != "")
            //        {
            //            urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
            //            List<string> tovarList = webRequest.arraySaveimage(urlTovar);
            //            if (action != "")
            //            {
            //                tovarList[39] = action;
            //                tovarList[9] = actualPrice.ToString();
            //                webRequest.saveImage(tovarList);
            //                editPrice++;
            //            }
            //            else
            //            {
            //                webRequest.deleteProduct(tovarList);
            //                deleteTovar++;
            //            }
            //        }
            //    }
            //    else
            //    {
            //        otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
            //        string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
            //        if (urlTovar != "")
            //        {
            //            urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
            //            List<string> tovarList = webRequest.arraySaveimage(urlTovar);
            //            double priceBike18 = Convert.ToDouble(tovarList[9].ToString());
            //            if(actualPrice != priceBike18)
            //            {
            //                tovarList[39] = action;
            //                tovarList[9] = actualPrice.ToString();
            //                webRequest.saveImage(tovarList);
            //                editPrice++;
            //            }else if(tovarList[39].ToString() != action)
            //            {
            //                tovarList[39] = action;
            //                webRequest.saveImage(tovarList);
            //                editPrice++;
            //            }
            //        }
            //        else
            //        {
            //            string name = (string)w.Cells[i, 3].Value;
            //            string stock = (string)w.Cells[i, 14].Value;
            //        }
            //    }
            //}
            //MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);

            //Запчасти для снегоходов
            //w = p.Workbook.Worksheets[4];
            //q = w.Dimension.Rows;

            //for (int i = 8; q > i; i++)
            //{
            //    if (w.Cells[i, 1].Value != null)
            //    {
            //        double articl = (double)w.Cells[i, 1].Value;
            //        double quantity = (double)w.Cells[i, 9].Value;
            //        double priceIrbisDiler = (double)w.Cells[i, 6].Value;
            //        double actualPrice = Price(priceIrbisDiler);
            //        string action = (string)w.Cells[i, 14].Value;

            //        ExcelRange er = w.Cells[i, 2];
            //        if (er.Hyperlink != null)
            //        {
            //            string urlImg = er.Hyperlink.ToString();
            //            try
            //            {
            //                webClient.DownloadFile(urlImg, "pic\\" + articl + ".jpg");
            //            }
            //            catch
            //            {

            //            }

            //        }

            //        if (action != "")
            //        {
            //            switch (action)
            //            {
            //                case "ЛУЧШАЯ ЦЕНА!":
            //                    action = "&markers[3]=1";
            //                    break;
            //                case "Новое поступление":
            //                    action = "&markers[1]=1";
            //                    break;
            //                case "Новое постуление":
            //                    action = "&markers[1]=1";
            //                    break;
            //                default:
            //                    action = "";
            //                    break;
            //            }
            //        }

            //        if (quantity == 0)
            //        {
            //            otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
            //            string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
            //            if (urlTovar != "")
            //            {
            //                urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
            //                List<string> tovarList = webRequest.arraySaveimage(urlTovar);
            //                if (action != "")
            //                {
            //                    tovarList[39] = action;
            //                    tovarList[9] = actualPrice.ToString();
            //                    webRequest.saveImage(tovarList);
            //                    editPrice++;
            //                }
            //                else
            //                {
            //                    webRequest.deleteProduct(tovarList);
            //                    deleteTovar++;
            //                }
            //            }
            //        }
            //        else
            //        {
            //            otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
            //            string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
            //            if (urlTovar != "")
            //            {
            //                urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
            //                List<string> tovarList = webRequest.arraySaveimage(urlTovar);
            //                double priceBike18 = Convert.ToDouble(tovarList[9].ToString());
            //                if (actualPrice != priceBike18)
            //                {
            //                    tovarList[39] = action;
            //                    tovarList[9] = actualPrice.ToString();
            //                    webRequest.saveImage(tovarList);
            //                    editPrice++;
            //                }
            //                else if (tovarList[39].ToString() != action)
            //                {
            //                    tovarList[39] = action;
            //                    webRequest.saveImage(tovarList);
            //                    editPrice++;
            //                }
            //            }
            //            else
            //            {
            //                string name = (string)w.Cells[i, 3].Value;
            //                string stock = (string)w.Cells[i, 14].Value;
            //            }
            //        }
            //    }


            //}

            w = p.Workbook.Worksheets[5];
            q = w.Dimension.Rows;

            for (int i = 8; q > i; i++)
            {
                if (w.Cells[i, 2].Value != null)
                {

                    double articl = (double)w.Cells[i, 2].Value;
                    double quantity = (double)w.Cells[i, 10].Value;
                    double priceIrbisDiler = (double)w.Cells[i, 7].Value;
                    double actualPrice = Price(priceIrbisDiler);
                    string action = (string)w.Cells[i, 14].Value;

                    ExcelRange er = w.Cells[i, 1];
                    if (er.Hyperlink != null)
                    {
                        string urlImg = er.Hyperlink.ToString();
                        try
                        {
                            webClient.DownloadFile(urlImg, "pic\\" + articl + ".jpg");
                        }
                        catch
                        {

                        }

                    }

                    if (action != "")
                    {
                        switch (action)
                        {
                            case "ЛУЧШАЯ ЦЕНА!":
                                action = "&markers[3]=1";
                                break;
                            case "Новое поступление":
                                action = "&markers[1]=1";
                                break;
                            case "Новое постуление":
                                action = "&markers[1]=1";
                                break;
                            default:
                                action = "";
                                break;
                        }
                    }

                    if (quantity == 0)
                    {
                        otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
                        string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
                        if (urlTovar != "")
                        {
                            urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
                            List<string> tovarList = webRequest.arraySaveimage(urlTovar);
                            if (action != "")
                            {
                                tovarList[39] = action;
                                tovarList[9] = actualPrice.ToString();
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                            else
                            {
                                webRequest.deleteProduct(tovarList);
                                deleteTovar++;
                            }
                        }
                    }
                    else
                    {
                        otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
                        string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
                        if (urlTovar != "")
                        {
                            urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
                            List<string> tovarList = webRequest.arraySaveimage(urlTovar);
                            double priceBike18 = Convert.ToDouble(tovarList[9].ToString());
                            if (actualPrice != priceBike18)
                            {
                                tovarList[39] = action;
                                tovarList[9] = actualPrice.ToString();
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                            else if (tovarList[39].ToString() != action)
                            {
                                tovarList[39] = action;
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                        }
                        else
                        {
                            string name = (string)w.Cells[i, 3].Value;
                            string stock = (string)w.Cells[i, 14].Value;
                        }
                    }
                }


            }

            MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);
        }

        public double Price(double priceDiler)
        {
            double discount = 0;
            double actualPrice = 0;
            if (priceDiler <= 15)
                discount = 2.7;
            else
            if (priceDiler <= 199)
                discount = 2.5;
            else
            if (priceDiler <= 2000)
                discount = 1.7;
            else
            if (priceDiler <= 7999)
                discount = 1.4;
            else
            if (priceDiler >= 8000)
                discount = 1.3;

            actualPrice = priceDiler * discount;
            actualPrice = Math.Round(actualPrice, 0);
            int price = Convert.ToInt32(actualPrice);
            price = (price / 10) * 10;
            actualPrice = Convert.ToDouble(price);

            return actualPrice;
        }
    }
}
