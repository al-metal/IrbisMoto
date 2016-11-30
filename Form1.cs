using Bike18;
using OfficeOpenXml;
using RacerMotors;
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
using Формирование_ЧПУ;

namespace IrbisMoto
{
    public partial class Form1 : Form
    {
        web.WebRequest webRequest = new web.WebRequest();
        CHPU chpu = new CHPU();
        WebClient webClient = new WebClient();
        FileEdit files = new FileEdit();
        List<string> allTovar = new List<string>();
        nethouse nethouse = new nethouse();
        httpRequest httpRequest = new httpRequest();

        string boldOpen = "<span style=\"\"font-weight: bold; font-weight: bold;\"\">";
        string boldClose = "</span>";
        string otv = null;
        int deleteTovar = 0;
        int editPrice = 0;

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
            Properties.Settings.Default.login = tbLogin.Text;
            Properties.Settings.Default.password = tbPassword.Text;
            Properties.Settings.Default.Save();

            CookieContainer cookie = nethouse.CookieNethouse(tbLogin.Text, tbPassword.Text);

            File.Delete("naSite.csv");
            File.Delete("allTovars");
            nethouse.NewListUploadinBike18("naSite");
            List<string> newProduct = new List<string>();

            FileInfo file = new FileInfo("Прайс.xlsx");
            ExcelPackage p = new ExcelPackage(file);


            ExcelWorksheet w = p.Workbook.Worksheets[3];
            int q = w.Dimension.Rows;

            #region Раздел запчасти
            for (int i = 8; q > i; i++)
            {
                if (w.Cells[i, 1].Value == null)
                    break;

                double articl = (double)w.Cells[i, 1].Value;
                allTovarInFile(articl);
                double quantity = (double)w.Cells[i, 9].Value;
                double priceIrbisDiler = (double)w.Cells[i, 6].Value;
                double actualPrice = Price(priceIrbisDiler);
                string action = (string)w.Cells[i, 14].Value;
                string name = (string)w.Cells[i, 3].Value;
                name = name.Replace("\"", "");

                ExcelRange er = w.Cells[i, 2];
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

                if (action != null)
                    action = actionText(action);
                else
                    action = "";

                otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
                string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();

                string slug = chpu.vozvr(name);
                int space = name.IndexOf(" ");
                string strRazdel = name.Remove(space, name.Length - space);

                string razdel = irbisZapchastiRazdel(strRazdel);

                string miniText = null;
                string titleText = null;
                string descriptionText = null;
                string keywordsText = null;
                string fullText = null;
                string discount = null;
                string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";

                string nameBold = boldOpen + name + boldClose;

                miniText = miniTextTemplate();
                fullText = fullTextTemplate();
                titleText = tbTitle.Lines[0].ToString();
                descriptionText = tbDescription.Lines[0].ToString();
                keywordsText = tbKeywords.Lines[0].ToString();
                discount = discountTemplate();

                miniText = miniText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString()).Replace("<p><br /></p><p><br /></p><p><br /></p><p>", "<p><br /></p>");
                miniText = miniText.Remove(miniText.LastIndexOf("<p>"));

                fullText = fullText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString());
                fullText = fullText.Remove(fullText.LastIndexOf("<p>"));

                titleText = titleText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                descriptionText = descriptionText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                keywordsText = keywordsText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                titleText = textRemove(titleText, 255);
                descriptionText = textRemove(descriptionText, 200);
                keywordsText = textRemove(keywordsText, 100);
                slug = textRemove(slug, 64);

                if (urlTovar == "")
                {
                    string stock = (string)w.Cells[i, 14].Value;
                    bool kioshi = name.Contains("KIYOSHI");
                    if (!kioshi)
                    {
                        newProduct = new List<string>();
                        newProduct.Add("");                                 //id
                        newProduct.Add("\"" + articl + "\"");               //артикул
                        newProduct.Add("\"" + name + "\"");                 //название
                        newProduct.Add("\"" + actualPrice + "\"");          //стоимость
                        newProduct.Add("\"" + "" + "\"");                   //со скидкой
                        newProduct.Add("\"" + razdel + "\"");               //раздел товара
                        newProduct.Add("\"" + "100" + "\"");                //в наличии
                        newProduct.Add("\"" + "0" + "\"");                  //поставка
                        newProduct.Add("\"" + "1" + "\"");                  //срок поставки
                        newProduct.Add("\"" + miniText + "\"");             //краткий текст
                        newProduct.Add("\"" + fullText + "\"");             //полностью текст
                        newProduct.Add("\"" + titleText + "\"");            //заголовок страницы
                        newProduct.Add("\"" + descriptionText + "\"");      //описание
                        newProduct.Add("\"" + keywordsText + "\"");         //ключевые слова
                        newProduct.Add("\"" + slug + "\"");                 //ЧПУ
                        newProduct.Add("");                                 //с этим товаром покупают
                        newProduct.Add("");                                 //рекламные метки
                        newProduct.Add("\"" + "1" + "\"");                  //показывать
                        newProduct.Add("\"" + "0" + "\"");                  //удалить

                        files.fileWriterCSV(newProduct, "naSite");
                    }
                }
                else
                {
                    List<string> tovarList = new List<string>();
                    bool izmen = false;
                    bool del = false;
                    tovarList = nethouse.GetProductList(cookie, urlTovar);
                    if(tovarList.Count == 0)
                    {
                        continue;
                    }
                    string alsoBy = nethouse.alsoBuyTovars(tovarList);
                    if (tovarList[43] != "100")
                    {
                        tovarList[43] = "100";
                        izmen = true;
                    }

                    if (quantity == 0)
                    {
                        if (action == "")
                        {
                            nethouse.DeleteProduct(cookie, tovarList);
                            del = true;
                        }
                        else
                            tovarList[43] = "100";
                        izmen = true;
                    }
                    else
                    {
                        double priceBike18 = Convert.ToDouble(tovarList[9].ToString());

                        if (actualPrice != priceBike18)
                        {
                            tovarList[9] = actualPrice.ToString();
                            editPrice++;
                        }
                    }

                    if (tovarList[39] != action)
                    {
                        tovarList[39] = action;
                        izmen = true;
                    }

                    tovarList[42] = alsoBy;
                    tovarList[1] = slug;
                    tovarList[7] = miniText;
                    tovarList[8] = fullText;
                    tovarList[11] = descriptionText;
                    tovarList[12] = keywordsText;
                    tovarList[13] = titleText;
                    tovarList[3] = "10833347";

                    if (izmen & !del)
                        nethouse.SaveTovar(cookie, tovarList);
                }
            }
            #endregion

            #region Запчасти для снегоходов
            w = p.Workbook.Worksheets[4];
            q = w.Dimension.Rows;
            string razdelSnegohod = null;
            for (int i = 7; q > i; i++)
            {
                if (w.Cells[i, 1].Value == null)
                {
                    razdelSnegohod = (string)w.Cells[i, 2].Value;
                }
                else
                {
                    double articl = (double)w.Cells[i, 1].Value;
                    allTovarInFile(articl);
                    double quantity = (double)w.Cells[i, 9].Value;
                    double priceIrbisDiler = (double)w.Cells[i, 6].Value;
                    double actualPrice = Price(priceIrbisDiler);
                    string action = (string)w.Cells[i, 14].Value;
                    string name = (string)w.Cells[i, 3].Value;
                    name = name.Replace("\"", "");

                    ExcelRange er = w.Cells[i, 2];
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
                        action = actionText(action);

                    otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
                    string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();

                    string slug = chpu.vozvr(name);
                    string razdel = irbisSnegohod(razdelSnegohod);
                    string miniText = null;
                    string titleText = null;
                    string descriptionText = null;
                    string keywordsText = null;
                    string fullText = null;
                    string discount = null;
                    string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";

                    string nameBold = boldOpen + name + boldClose;

                    miniText = miniTextTemplate();
                    fullText = fullTextTemplate();
                    titleText = tbTitle.Lines[0].ToString();
                    descriptionText = tbDescription.Lines[0].ToString();
                    keywordsText = tbKeywords.Lines[0].ToString();
                    discount = discountTemplate();

                    miniText = miniText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString()).Replace("<p><br /></p><p><br /></p><p><br /></p><p>", "<p><br /></p>");

                    miniText = miniText.Remove(miniText.LastIndexOf("<p>"));

                    fullText = fullText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString());

                    fullText = fullText.Remove(fullText.LastIndexOf("<p>"));

                    titleText = titleText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                    descriptionText = descriptionText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                    keywordsText = keywordsText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                    titleText = textRemove(titleText, 255);
                    descriptionText = textRemove(descriptionText, 200);
                    keywordsText = textRemove(keywordsText, 100);
                    slug = textRemove(slug, 64);

                    if (urlTovar == "")
                    {
                        string stock = (string)w.Cells[i, 14].Value;

                        newProduct = new List<string>();
                        newProduct.Add("");                                 //id
                        newProduct.Add("\"" + articl + "\"");               //артикул
                        newProduct.Add("\"" + name + "\"");                 //название
                        newProduct.Add("\"" + actualPrice + "\"");          //стоимость
                        newProduct.Add("\"" + "" + "\"");                   //со скидкой
                        newProduct.Add("\"" + razdel + "\"");               //раздел товара
                        newProduct.Add("\"" + "100" + "\"");                //в наличии
                        newProduct.Add("\"" + "0" + "\"");                  //поставка
                        newProduct.Add("\"" + "1" + "\"");                  //срок поставки
                        newProduct.Add("\"" + miniText + "\"");             //краткий текст
                        newProduct.Add("\"" + fullText + "\"");             //полностью текст
                        newProduct.Add("\"" + titleText + "\"");            //заголовок страницы
                        newProduct.Add("\"" + descriptionText + "\"");      //описание
                        newProduct.Add("\"" + keywordsText + "\"");         //ключевые слова
                        newProduct.Add("\"" + slug + "\"");                 //ЧПУ
                        newProduct.Add("");                                 //с этим товаром покупают
                        newProduct.Add("");                                 //рекламные метки
                        newProduct.Add("\"" + "1" + "\"");                  //показывать
                        newProduct.Add("\"" + "0" + "\"");                  //удалить

                        files.fileWriterCSV(newProduct, "naSite");
                    }
                    else
                    {
                        bool izmen = false;
                        bool del = false;
                        List<string> tovarList = nethouse.GetProductList(cookie, urlTovar);
                        string alsoBy = nethouse.alsoBuyTovars(tovarList);
                        if (tovarList[43] != "100")
                        {
                            tovarList[43] = "100";
                            izmen = true;
                        }

                        if (quantity == 0)
                        {
                            if (action == "")
                            {
                                nethouse.DeleteProduct(cookie, tovarList);
                                del = true;
                            }
                            else
                                tovarList[43] = "100";
                            izmen = true;
                        }
                        else
                        {
                            double priceBike18 = Convert.ToDouble(tovarList[9].ToString());

                            if (actualPrice != priceBike18)
                            {
                                tovarList[9] = actualPrice.ToString();
                                editPrice++;
                            }
                        }

                        if (tovarList[39] != action)
                        {
                            tovarList[39] = action;
                            izmen = true;
                        }

                        tovarList[42] = alsoBy;
                        tovarList[1] = slug;
                        tovarList[7] = miniText;
                        tovarList[8] = fullText;
                        tovarList[11] = descriptionText;
                        tovarList[12] = keywordsText;
                        tovarList[13] = titleText;
                        tovarList[3] = "10833347";

                        if (izmen & !del)
                            nethouse.SaveTovar(cookie, tovarList);
                    }
                }
            }
            #endregion

            #region Запчасти KIYOSHI
            string razdelkiyoshi = null;
            w = p.Workbook.Worksheets[5];
            q = w.Dimension.Rows;

            for (int i = 7; q > i; i++)
            {
                if (w.Cells[i, 2].Value == null)
                {
                    razdelkiyoshi = (string)w.Cells[i, 1].Value;
                }
                else
                {

                    double articl = (double)w.Cells[i, 2].Value;
                    allTovarInFile(articl);
                    double quantity = (double)w.Cells[i, 10].Value;
                    double priceIrbisDiler = (double)w.Cells[i, 7].Value;
                    double actualPrice = Price(priceIrbisDiler);
                    string action = (string)w.Cells[i, 14].Value;
                    string name = (string)w.Cells[i, 3].Value;
                    name = name.Replace("\"", "").Replace("\n", "");

                    string slug = chpu.vozvr(name);
                    string razdel = irbisKiyoshiRazdel(razdelkiyoshi);
                    string miniText = null;
                    string titleText = null;
                    string descriptionText = null;
                    string keywordsText = null;
                    string fullText = null;
                    string discount = null;
                    string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";

                    string nameBold = boldOpen + name + boldClose;

                    miniText = miniTextTemplate();
                    fullText = fullTextTemplate();
                    titleText = tbTitle.Lines[0].ToString();
                    descriptionText = tbDescription.Lines[0].ToString();
                    keywordsText = tbKeywords.Lines[0].ToString();
                    discount = discountTemplate();

                    miniText = miniText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString()).Replace("<p><br /></p><p><br /></p><p><br /></p><p>", "<p><br /></p>");

                    miniText = miniText.Remove(miniText.LastIndexOf("<p>"));

                    fullText = fullText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString());

                    fullText = fullText.Remove(fullText.LastIndexOf("<p>"));

                    titleText = titleText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                    descriptionText = descriptionText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                    keywordsText = keywordsText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                    titleText = textRemove(titleText, 255);
                    descriptionText = textRemove(descriptionText, 200);
                    keywordsText = textRemove(keywordsText, 100);
                    slug = textRemove(slug, 64);

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
                        action = actionText(action);

                    otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
                    string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();

                    if (urlTovar == "")
                    {
                        string stock = (string)w.Cells[i, 14].Value;

                        newProduct = new List<string>();
                        newProduct.Add("");                                 //id
                        newProduct.Add("\"" + articl + "\"");               //артикул
                        newProduct.Add("\"" + name + "\"");                 //название
                        newProduct.Add("\"" + actualPrice + "\"");          //стоимость
                        newProduct.Add("\"" + "" + "\"");                   //со скидкой
                        newProduct.Add("\"" + razdel + "\"");               //раздел товара
                        newProduct.Add("\"" + "100" + "\"");                //в наличии
                        newProduct.Add("\"" + "0" + "\"");                  //поставка
                        newProduct.Add("\"" + "1" + "\"");                  //срок поставки
                        newProduct.Add("\"" + miniText + "\"");             //краткий текст
                        newProduct.Add("\"" + fullText + "\"");             //полностью текст
                        newProduct.Add("\"" + titleText + "\"");            //заголовок страницы
                        newProduct.Add("\"" + descriptionText + "\"");      //описание
                        newProduct.Add("\"" + keywordsText + "\"");         //ключевые слова
                        newProduct.Add("\"" + slug + "\"");                 //ЧПУ
                        newProduct.Add("");                                 //с этим товаром покупают
                        newProduct.Add("");                                 //рекламные метки
                        newProduct.Add("\"" + "1" + "\"");                  //показывать
                        newProduct.Add("\"" + "0" + "\"");                  //удалить

                        files.fileWriterCSV(newProduct, "naSite");
                    }
                    else
                    {
                        bool izmen = false;
                        bool del = false;
                        List<string> tovarList = nethouse.GetProductList(cookie, urlTovar);
                        string alsoBy = nethouse.alsoBuyTovars(tovarList);
                        if (tovarList[43] != "100")
                        {
                            tovarList[43] = "100";
                            izmen = true;
                        }

                        if (quantity == 0)
                        {
                            if (action == "")
                            {
                                nethouse.DeleteProduct(cookie, tovarList);
                                del = true;
                            }
                            else
                                tovarList[43] = "100";
                            izmen = true;
                        }
                        else
                        {
                            double priceBike18 = Convert.ToDouble(tovarList[9].ToString());

                            if (actualPrice != priceBike18)
                            {
                                tovarList[9] = actualPrice.ToString();
                                editPrice++;
                            }
                        }

                        if (tovarList[39] != action)
                        {
                            tovarList[39] = action;
                            izmen = true;
                        }

                        tovarList[42] = alsoBy;
                        tovarList[1] = slug;
                        tovarList[7] = miniText;
                        tovarList[8] = fullText;
                        tovarList[11] = descriptionText;
                        tovarList[12] = keywordsText;
                        tovarList[13] = titleText;
                        tovarList[3] = "10833347";

                        if (izmen & !del)
                            nethouse.SaveTovar(cookie, tovarList);
                    }
                }
            }
            #endregion

            #region uploadInSIte
            System.Threading.Thread.Sleep(20000);
            string[] naSite1 = File.ReadAllLines("naSite.csv", Encoding.GetEncoding(1251));
            if (naSite1.Length > 1)
            {
                nethouse.UploadCSVNethouse(cookie, "naSite.csv");
            }
            #endregion

            otv = httpRequest.getRequest("http://bike18.ru/products/category/1183836");
            MatchCollection razdelSite = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            string[] allprod = File.ReadAllLines("allTovars");
            for (int i = 0; razdelSite.Count > i; i++)
            {
                otv = httpRequest.getRequest("http://bike18.ru" + razdelSite[i].ToString() + "/page/all");
                MatchCollection product = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Matches(otv);
                for (int n = 0; product.Count > n; n++)
                {
                    string urlTovar = product[n].ToString();
                    otv = httpRequest.getRequest(urlTovar);
                    string artProd = new Regex("(?<=Артикул:)[\\w\\W]*?(?=</div)").Match(otv).ToString().Trim();
                    bool b = false;

                    foreach (string str in allprod)
                    {
                        if (artProd == str)
                        {
                            b = true;
                            break;
                        }
                    }

                    if (!b)
                    {
                        nethouse.DeleteProduct(cookie, urlTovar);
                        deleteTovar++;
                    }
                }
            }
            MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);
        }

        private string irbisKiyoshiRazdel(string razdelkiyoshi)
        {
            string podrazdel = "";
            switch (razdelkiyoshi)
            {
                case "Амортизаторы":
                    podrazdel = "Амортизаторы KIYOSHI";
                    break;
                case "Воздушные фильтры нулевого сопротивления":
                    podrazdel = "Воздушные фильтры KIYOSHI";
                    break;
                case "Глушители спортивные":
                    podrazdel = "Спортивные глушители KIYOSHI";
                    break;
                case "Карбюраторы, жиклеры карбюраторов":
                    podrazdel = "Карбюраторы, жиклеры карбюраторов KIYOSHI";
                    break;
                case "Электрооборудование":
                    podrazdel = "Электрооборудование KIYOSHI";
                    break;
                case "Валы коленчатые":
                    podrazdel = "Сцепления, барабаны, пружины сцепления KIYOSHI";
                    break;
                case "Подшипники":
                    podrazdel = "Подшипники KIYOSHI";
                    break;
                case "Вариаторы, грузики вариатора":
                    podrazdel = "Вариаторы, грузики вариатора KIYOSHI";
                    break;
                case "Ремни вариатора":
                    podrazdel = "Ремни вариатора KIYOSHI";
                    break;
                case "Сцепления, барабаны, пружины сцепления":
                    podrazdel = "Сцепления, барабаны, пружины сцепления KIYOSHI";
                    break;
                case "Цилиндро-поршневые группы":
                    podrazdel = "Цилиндро- поршневые группы KIYOSHI";
                    break;
                case "Лепестковые клапаны":
                    podrazdel = "Лепестковые клапаны KIYOSHI";
                    break;
                case "Газораспределительный механизм":
                    podrazdel = "Газораспределительный механизм KIYOSHI";
                    break;
                case "Стайлинг":
                    podrazdel = "Стайлинг KIYOSHI";
                    break;
                case "Наклейки":
                    podrazdel = "Стайлинг KIYOSHI";
                    break;
                default:
                    break;
            }

            string razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => " + podrazdel;
            return razdel;
        }

        private string irbisSnegohod(string razdelSnegohod)
        {
            string podrazdel = "";
            switch (razdelSnegohod)
            {
                case "гусеницы":
                    podrazdel = "Гусеницы";
                    break;
                case "снегоходы Буран":
                    podrazdel = "Снегоходы Буран";
                    break;
                case "снегоходы Тайга":
                    podrazdel = "Снегоходы Тайга";
                    break;
                case "снегоходы Dingo":
                    podrazdel = "Снегоходы Dingo";
                    break;
                case "мотобуксировщики Мухтар":
                    podrazdel = "Мотобуксировщики Мухтар";
                    break;
                default:
                    podrazdel = "Разное";
                    break;
            }
            string razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => " + podrazdel;
            return razdel;
        }

        private void allTovarInFile(double articl)
        {
            string article = articl.ToString();
            StreamWriter sw = new StreamWriter("allTovars", true);
            sw.WriteLine(article);
            sw.Close();
        }

        private string actionText(string action)
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
                case "Новинка":
                    action = "&markers[1]=1";
                    break;
                default:
                    action = "";
                    break;
            }
            return action;
        }

        private string textRemove(string text, int count)
        {
            if (text.Length > count)
            {
                text = text.Remove(count);
                text = text.Remove(text.LastIndexOf(" "));
            }
            return text;
        }

        private string discountTemplate()
        {
            string discount = "<p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> Сделай ТРОЙНОЙ удар по нашим ценам! </span></p><p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> 1. <a target=\"\"_blank\"\" href =\"\"http://bike18.ru/stock\"\"> Скидки за отзывы о товарах!</a> </span></p><p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> 2. <a target=\"\"_blank\"\" href =\"\"http://bike18.ru/stock\"\"> Друзьям скидки и подарки!</a> </span></p><p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> 3. <a target=\"\"_blank\"\" href =\"\"http://bike18.ru/stock\"\"> Нашли дешевле!? 110% разницы Ваши!</a></span></p>";
            return discount;
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

        private void btnUpdateImage_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.login = tbLogin.Text;
            Properties.Settings.Default.password = tbPassword.Text;
            Properties.Settings.Default.Save();

            CookieContainer cookie = nethouse.CookieNethouse(tbLogin.Text, tbPassword.Text);

            int countUpdateImage = 0;
            otv = webRequest.getRequest("http://bike18.ru/products/category/katalog-zapchastey-irbis");
            MatchCollection razdel = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            for (int i = 0; razdel.Count > i; i++)
            {
                otv = webRequest.getRequest("http://bike18.ru" + razdel[i].ToString() + "/page/all");
                MatchCollection tovar = new Regex("(?<=<div class=\"product-link -text-center\"><a href=\").*?(?=\" >)").Matches(otv);
                for (int n = 0; tovar.Count > n; n++)
                {
                    otv = webRequest.getRequest(tovar[n].ToString());
                    string urlImageTovar = new Regex("(?<=class=\"avatar-view \"><link rel=\"image_src\" href=\").*?(?=\">)").Match(otv).ToString();
                    if (urlImageTovar == "")
                    {
                        string articl = new Regex("(?<= Артикул:)[\\w\\W]*?(?=</div>)").Match(otv).ToString();
                        articl = articl.Trim();
                        if (File.Exists("pic\\" + articl + ".jpg"))
                        {
                            nethouse.UploadImage(cookie, tovar[n].ToString());
                            countUpdateImage++;
                        }
                    }
                }
            }

            otv = webRequest.getRequest("http://bike18.ru/products/category/katalog-zapchastey-kiyoshi");
            razdel = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            for (int i = 0; razdel.Count > i; i++)
            {
                otv = webRequest.getRequest("http://bike18.ru" + razdel[i].ToString() + "/page/all");
                MatchCollection tovar = new Regex("(?<=<div class=\"product-link -text-center\"><a href=\").*?(?=\" >)").Matches(otv);
                for (int n = 0; tovar.Count > n; n++)
                {
                    otv = webRequest.getRequest(tovar[n].ToString());
                    string urlImageTovar = new Regex("(?<=class=\"avatar-view \"><link rel=\"image_src\" href=\").*?(?=\">)").Match(otv).ToString();
                    if (urlImageTovar == "")
                    {
                        string articl = new Regex("(?<= Артикул:)[\\w\\W]*?(?=</div>)").Match(otv).ToString();
                        articl = articl.Trim();
                        if (File.Exists("pic\\" + articl + ".jpg"))
                        {
                            nethouse.UploadImage(cookie, tovar[n].ToString());
                            countUpdateImage++;
                        }
                    }
                }
            }

            otv = webRequest.getRequest("http://bike18.ru/products/category/zapchasti-dlya-snegohodov-i-motobuksirovshchikov");
            razdel = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            for (int i = 0; razdel.Count > i; i++)
            {
                otv = webRequest.getRequest("http://bike18.ru" + razdel[i].ToString() + "/page/all");
                MatchCollection tovar = new Regex("(?<=<div class=\"product-link -text-center\"><a href=\").*?(?=\" >)").Matches(otv);
                for (int n = 0; tovar.Count > n; n++)
                {
                    otv = webRequest.getRequest(tovar[n].ToString());
                    string urlImageTovar = new Regex("(?<=class=\"avatar-view \"><link rel=\"image_src\" href=\").*?(?=\">)").Match(otv).ToString();
                    if (urlImageTovar == "")
                    {
                        string articl = new Regex("(?<= Артикул:)[\\w\\W]*?(?=</div>)").Match(otv).ToString();
                        articl = articl.Trim();
                        if (File.Exists("pic\\" + articl + ".jpg"))
                        {
                            nethouse.UploadImage(cookie, tovar[n].ToString());
                            countUpdateImage++;
                        }
                    }
                }
            }

            MessageBox.Show("Обновлено картинок: " + countUpdateImage);
        }

        public string irbisZapchastiRazdel(string strRazdel)
        {
            string razdel = "Запчасти и расходники => Каталог запчастей IRBIS => ";
            switch (strRazdel)
            {
                case "Аккумуляторная":
                    razdel = razdel + "Аккумуляторы";
                    break;
                case "Амортизатор":
                    razdel = razdel + "Амортизаторы";
                    break;
                case "Амортизаторы":
                    razdel = razdel + "Амортизаторы";
                    break;
                case "Багажник":
                    razdel = razdel + "Багажники";
                    break;
                case "Бак":
                    razdel = razdel + "Баки масляные, топливные, системы охлаждения";
                    break;
                case "Барабан":
                    razdel = razdel + "Блоки переключения, бендиксы, барабаны";
                    break;
                case "Бачок":
                    razdel = razdel + "Баки масляные, топливные, системы охлаждения";
                    break;
                case "Бендикс":
                    razdel = razdel + "Блоки переключения, бендиксы, барабаны";
                    break;
                case "Блок":
                    razdel = razdel + "Блоки переключения, бендиксы, барабаны";
                    break;
                case "Блоки":
                    razdel = razdel + "Блоки переключения, бендиксы, барабаны";
                    break;
                case "Болты":
                    razdel = razdel + "Болты, буксы";
                    break;
                case "Вал":
                    razdel = razdel + "Валы";
                    break;
                case "Валы":
                    razdel = razdel + "Валы";
                    break;
                case "Вариатор":
                    razdel = razdel + "Вариаторы";
                    break;
                case "Вентиль":
                    razdel = razdel + "Вентили";
                    break;
                case "Вилка":
                    razdel = razdel + "Вилки переключения передач";
                    break;
                case "Втулка":
                    razdel = razdel + "Втулки";
                    break;
                case "Генератор":
                    razdel = razdel + "Генераторы в сборе";
                    break;
                case "Глушитель":
                    razdel = razdel + "Глушители";
                    break;
                case "Головка":
                    razdel = razdel + "Головки цилиндра";
                    break;
                case "грузики":
                    razdel = razdel + "Грузики вариатора";
                    break;
                case "Датчик":
                    razdel = razdel + "Датчики";
                    break;
                case "Двигатель":
                    razdel = razdel + "Двигатели в сборе";
                    break;
                case "Демпфер":
                    razdel = razdel + "Демпферы";
                    break;
                case "Демпферные":
                    razdel = razdel + "Демпферы";
                    break;
                case "Диск":
                    razdel = razdel + "Колесные диски";
                    break;
                case "Диски":
                    razdel = razdel + "Диски сцепления";
                    break;
                case "Жгут":
                    razdel = razdel + "Жгуты проводов";
                    break;
                case "Замков":
                    razdel = razdel + "Замки";
                    break;
                case "Замок":
                    razdel = razdel + "Замки";
                    break;
                case "Защита":
                    razdel = razdel + "Защита двигателя";
                    break;
                case "Звезда":
                    razdel = razdel + "Звезды";
                    break;
                case "Зубчатый":
                    razdel = razdel + "Зубчатые сектора";
                    break;
                case "Камера":
                    razdel = razdel + "Камеры";
                    break;
                case "Индикатор":
                    razdel = razdel + "Индикаторы";
                    break;
                case "Карбюратор":
                    razdel = razdel + "Карбюраторы";
                    break;
                case "Катушка":
                    razdel = razdel + "Катушки зажигания";
                    break;
                case "Клапан":
                    razdel = razdel + "Клапаны";
                    break;
                case "Клапаны":
                    razdel = razdel + "Клапаны";
                    break;
                case "Клипса":
                    razdel = razdel + "Клипсы";
                    break;
                case "Кнопка":
                    razdel = razdel + "Кнопки";
                    break;
                case "Кнопки":
                    razdel = razdel + "Кнопки";
                    break;
                case "Кожух":
                    razdel = razdel + "Кожухи";
                    break;
                case "Коллектор":
                    razdel = razdel + "Коллекторы";
                    break;
                case "Колодки":
                    razdel = razdel + "Тормозные колодки";
                    break;
                case "Колпачок":
                    razdel = razdel + "Свечные колпачки";
                    break;
                case "Кольца":
                    razdel = razdel + "Кольца";
                    break;
                case "Кольцо":
                    razdel = razdel + "Кольца";
                    break;
                case "Коммутатор":
                    razdel = razdel + "Коммутаторы";
                    break;
                case "Коромысла":
                    razdel = razdel + "Коромысла";
                    break;
                case "Корпус":
                    razdel = razdel + "Корпуса картеров, предохранителей";
                    break;
                case "Кран":
                    razdel = razdel + "Топливные краны";
                    break;
                case "Крепление":
                    razdel = razdel + "Крепления и кронштейны";
                    break;
                case "Кронштейн":
                    razdel = razdel + "Крепления и кронштейны";
                    break;
                case "Крыло":
                    razdel = razdel + "Крылья";
                    break;
                case "Крыльчатка":
                    razdel = razdel + "Крыльчатки";
                    break;
                case "Крышка":
                    razdel = razdel + "Крышки";
                    break;
                case "Лампа":
                    razdel = razdel + "Лампы";
                    break;
                case "Машинка":
                    razdel = razdel + "Тормозные машинки";
                    break;
                case "Мембрана":
                    razdel = razdel + "Мембраны карбюратора";
                    break;
                case "Муфта":
                    razdel = razdel + "Обгонные муфты";
                    break;
                case "Наконечник":
                    razdel = razdel + "Наконечники рулевой тяги";
                    break;
                case "Направляющие":
                    razdel = razdel + "Направляющие цепи";
                    break;
                case "Насос":
                    razdel = razdel + "Насосы";
                    break;
                case "Натяжитель":
                    razdel = razdel + "Натяжители цепи";
                    break;
                case "Обтекатели":
                    razdel = razdel + "Обтекатели";
                    break;
                case "Обтекатель":
                    razdel = razdel + "Обтекатели";
                    break;
                case "Опора":
                    razdel = razdel + "Опоры";
                    break;
                case "Ось":
                    razdel = razdel + "Оси";
                    break;
                case "Палец":
                    razdel = razdel + "Поршневые пальцы";
                    break;
                case "Панель":
                    razdel = razdel + "Панели приборов";
                    break;
                case "Патрубок":
                    razdel = razdel + "Патрубки";
                    break;
                case "Педаль":
                    razdel = razdel + "Педали тормоза";
                    break;
                case "Пластик":
                    razdel = razdel + "Пластик";
                    break;
                case "Подножка":
                    razdel = razdel + "Подножки";
                    break;
                case "Подножки":
                    razdel = razdel + "Подножки";
                    break;
                case "Подшипник":
                    razdel = razdel + "Подшипники";
                    break;
                case "Подшипники":
                    razdel = razdel + "Подшипники";
                    break;
                case "Поршневой":
                    razdel = razdel + "Поршни";
                    break;
                case "Привод":
                    razdel = razdel + "Приводы спидометра";
                    break;
                case "Прокладка":
                    razdel = razdel + "Прокладки";
                    break;
                case "Прокладки":
                    razdel = razdel + "Прокладки";
                    break;
                case "Пружина":
                    razdel = razdel + "Пружины";
                    break;
                case "Пружины":
                    razdel = razdel + "Пружины";
                    break;
                case "Радиатор":
                    razdel = razdel + "Радиаторы";
                    break;
                case "Рама":
                    razdel = razdel + "Рамы";
                    break;
                case "Реле":
                    razdel = razdel + "Реле";
                    break;
                case "Реле-регулятор":
                    razdel = razdel + "Реле";
                    break;
                case "Ремень":
                    razdel = razdel + "Ремни вариатора";
                    break;
                case "Ремкомплект":
                    razdel = razdel + "Ремкомплекты карбюраторов";
                    break;
                case "Решетка":
                    razdel = razdel + "Решетки";
                    break;
                case "Ролик":
                    razdel = razdel + "Ролики натяжителя цепи";
                    break;
                case "Ротор":
                    razdel = razdel + "Роторы";
                    break;
                case "Руль":
                    razdel = razdel + "Рули";
                    break;
                case "Ручка":
                    razdel = razdel + "Ручки, рычаги";
                    break;
                case "Ручки":
                    razdel = razdel + "Ручки, рычаги";
                    break;
                case "Рычаг":
                    razdel = razdel + "Ручки, рычаги";
                    break;
                case "Рычаги":
                    razdel = razdel + "Ручки, рычаги";
                    break;
                case "Сайлентблок":
                    razdel = razdel + "Сайлентблоки, сальники";
                    break;
                case "Сальник":
                    razdel = razdel + "Сайлентблоки, сальники";
                    break;
                case "Сальники":
                    razdel = razdel + "Сайлентблоки, сальники";
                    break;
                case "Свеча":
                    razdel = razdel + "Свечи зажигания";
                    break;
                case "Сигнал":
                    razdel = razdel + "Звуковые сигналы";
                    break;
                case "Сиденье":
                    razdel = razdel + "Сиденья";
                    break;
                case "Спица":
                    razdel = razdel + "Спицы";
                    break;
                case "Стартер":
                    razdel = razdel + "Статоры генератора";
                    break;
                case "Статор":
                    razdel = razdel + "Статоры генератора";
                    break;
                case "Ступица":
                    razdel = razdel + "Ступицы";
                    break;
                case "Суппорт":
                    razdel = razdel + "Суппорты";
                    break;
                case "Сцепление":
                    razdel = razdel + "Сцепление в сборе";
                    break;
                case "Толкатель":
                    razdel = razdel + "Толкатели";
                    break;
                case "Тормоз":
                    razdel = razdel + "Тормоза";
                    break;
                case "Траверса":
                    razdel = razdel + "Траверсы";
                    break;
                case "Трос":
                    razdel = razdel + "Тросы";
                    break;
                case "Турбина":
                    razdel = razdel + "Трубки, турбины";
                    break;
                case "Тяга":
                    razdel = razdel + "Тяги";
                    break;
                case "Указатели":
                    razdel = razdel + "Указатели поворотов";
                    break;
                case "Успокоитель":
                    razdel = razdel + "Успокоители цепи";
                    break;
                case "Фара":
                    razdel = razdel + "Фары";
                    break;
                case "Фильтр":
                    razdel = razdel + "Фильтры";
                    break;
                case "Фильтрующий":
                    razdel = razdel + "Фильтры";
                    break;
                case "Фонарь":
                    razdel = razdel + "Фары";
                    break;
                case "Цапфа":
                    razdel = razdel + "Цапфы";
                    break;
                case "Цепь":
                    razdel = razdel + "Цепи";
                    break;
                case "Цилиндро-поршневая":
                    razdel = razdel + "ЦПГ";
                    break;
                case "Шестерни":
                    razdel = razdel + "Шестерни и шайбы";
                    break;
                case "Шестерня":
                    razdel = razdel + "Шестерни и шайбы";
                    break;
                case "Шина":
                    razdel = razdel + "Шины";
                    break;
                case "Шланг":
                    razdel = razdel + "Шланги";
                    break;
                case "Электроклапан":
                    razdel = razdel + "Электроклапана карбюраторов";
                    break;
                case "Электростартер":
                    razdel = razdel + "Электростартеры";
                    break;
                default:
                    razdel = razdel + "Разное";
                    break;
            }
            return razdel;
        }

        public string miniTextTemplate()
        {
            string miniText = null;
            for (int z = 0; rtbMiniText.Lines.Length > z; z++)
            {
                if (rtbMiniText.Lines[z].ToString() == "")
                {
                    miniText += "<p><br /></p>";
                }
                else
                {
                    miniText += "<p>" + rtbMiniText.Lines[z].ToString() + "</p>";
                }
            }
            return miniText;
        }

        private string fullTextTemplate()
        {
            string fullText = null;
            for (int z = 0; rtbFullText.Lines.Length > z; z++)
            {
                if (rtbFullText.Lines[z].ToString() == "")
                {
                    fullText += "<p><br /></p>";
                }
                else
                {
                    fullText += "<p>" + rtbFullText.Lines[z].ToString() + "</p>";
                }
            }
            return fullText;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tbLogin.Text = Properties.Settings.Default.login;
            tbPassword.Text = Properties.Settings.Default.password;
        }
    }
}
//проект на 1600 строк
