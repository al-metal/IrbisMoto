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
using xNet.Net;
using web;
using Формирование_ЧПУ;
using System.Threading;

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

        Thread forms;

        string boldOpenSite = "<span style=\"font-weight: bold; font-weight: bold;\">";
        string boldOpenCSV = "<span style=\"\"font-weight: bold; font-weight: bold;\"\">";
        string boldOpen;
        string boldClose = "</span>";
        string otv = null;
        int deleteTovar = 0;
        int editPrice = 0;
        string minitextTemplate;
        string fullTextTemplate;
        string keywordsTextTemplate;
        string titleTextTemplate;
        string descriptionTextTemplate;
        string discountTemplate;
        List<string> newProduct = new List<string>();

        bool chekedEditMiniText;

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

        private void Form1_Load(object sender, EventArgs e)
        {
            tbLogin.Text = Properties.Settings.Default.login;
            tbPassword.Text = Properties.Settings.Default.password;
        }

        #region Обработка нажатия кнопок

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

            minitextTemplate = MiniTextTemplateStr();
            fullTextTemplate = FullTextTemplateStr();
            keywordsTextTemplate = tbKeywords.Lines[0].ToString();
            titleTextTemplate = tbTitle.Lines[0].ToString();
            descriptionTextTemplate = tbDescription.Lines[0].ToString();
            discountTemplate = discountTemplateStr();

            Thread tabl = new Thread(() => UpdateTovar());
            forms = tabl;
            forms.IsBackground = true;
            forms.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.login = tbLogin.Text;
            Properties.Settings.Default.password = tbPassword.Text;
            Properties.Settings.Default.Save();

            minitextTemplate = MiniTextTemplateStr();
            fullTextTemplate = FullTextTemplateStr();
            keywordsTextTemplate = tbKeywords.Lines[0].ToString();
            titleTextTemplate = tbTitle.Lines[0].ToString();
            descriptionTextTemplate = tbDescription.Lines[0].ToString();
            discountTemplate = discountTemplateStr();

            Thread tabl = new Thread(() => UpdateTovarSnegohod());
            forms = tabl;
            forms.IsBackground = true;
            forms.Start();
        }

        private void btnAccessory_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.login = tbLogin.Text;
            Properties.Settings.Default.password = tbPassword.Text;
            Properties.Settings.Default.Save();

            minitextTemplate = MiniTextTemplateStr();
            fullTextTemplate = FullTextTemplateStr();
            keywordsTextTemplate = tbKeywords.Lines[0].ToString();
            titleTextTemplate = tbTitle.Lines[0].ToString();
            descriptionTextTemplate = tbDescription.Lines[0].ToString();
            discountTemplate = discountTemplateStr();

            Thread tabl = new Thread(() => UpdateTovarAccessory());
            forms = tabl;
            forms.IsBackground = true;
            forms.Start();
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

        #endregion

        #region Обработка разделов

        private void UpdateTovarAccessory()
        {
            ControlsFormEnabledFalse();
            CookieContainer cookie = nethouse.CookieNethouse(tbLogin.Text, tbPassword.Text);
            if (cookie.Count == 1)
            {
                MessageBox.Show("Логин или пароль для сайта введены не верно", "Ошибка логина/пароля");
                ControlsFormEnabledTrue();
                return;
            }

            File.Delete("naSite.csv");
            File.Delete("allTovars");
            nethouse.NewListUploadinBike18("naSite");

            chekedEditMiniText = cbMiniText.Checked;

            FileInfo file = new FileInfo("Прайс-лист ТД Мегаполис 07.07.2017 Москва.xlsx");
            ExcelPackage p = new ExcelPackage(file);

            ExcelWorksheet w = p.Workbook.Worksheets[5];
            int q = w.Dimension.Rows;
            string razdelSnegohod = null;
            for (int i = 7; q > i; i++)
            {
                if (w.Cells[i, 3].Value == null)
                {
                    razdelSnegohod = (string)w.Cells[i, 1].Value;
                }
                else
                {
                    double articl;
                    try { articl = (double)w.Cells[i, 3].Value; }
                    catch
                    {
                        continue;
                    }
                    allTovarInFile(articl);
                    double quantity = (double)w.Cells[i, 10].Value;
                    double priceIrbisDiler = (double)w.Cells[i, 7].Value;
                    double actualPrice = Price(priceIrbisDiler);
                    string action = (string)w.Cells[i, 14].Value;
                    string name = (string)w.Cells[i, 4].Value;
                    name = name.Replace("\"", "");

                    ExcelRange er = w.Cells[i, 2];
                    if (er.Hyperlink == null && name.Contains("Наклей"))
                    {
                        continue;
                    }
                    DownloadImages(er, articl);

                    if (action != "")
                        action = actionText(action);

                    string urlTovar = nethouse.searchTovar(name, articl.ToString());
                    if (urlTovar == null)
                        urlTovar = nethouse.searchTovar(name, "IRB_" + articl.ToString());

                    if (urlTovar == "" || urlTovar == null)
                    {
                        WriteInCSV(articl.ToString(), name, razdelSnegohod, actualPrice.ToString());
                    }
                    else
                    {
                        boldOpen = boldOpenSite;
                        List<string> tovarList = new List<string>();
                        tovarList = nethouse.GetProductList(cookie, urlTovar);

                        UpdateInfoTova(cookie, tovarList, urlTovar, name, articl, quantity, action, actualPrice);
                    }
                }
            }

            UploadTovarInSite(cookie);

            otv = httpRequest.getRequest("https://bike18.ru/products/category/aksessuary-i-instrumenty-virz-dlya-mototehniki");
            MatchCollection razdelSite = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            string[] allprod = File.ReadAllLines("allTovars");
            for (int i = 0; razdelSite.Count > i; i++)
            {
                otv = httpRequest.getRequest("https://bike18.ru" + razdelSite[i].ToString() + "?page=all");
                MatchCollection product = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Matches(otv);
                for (int n = 0; product.Count > n; n++)
                {
                    string urlTovar = product[n].ToString();
                    otv = httpRequest.getRequest(urlTovar);

                    DeleteTovarNonePrice(cookie, urlTovar, allprod);
                }
            }

            MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);
            ControlsFormEnabledTrue();
        }

        private void UpdateInfoTova(CookieContainer cookie, List<string> tovarList, string urlTovar, string name, double articl, double quantity, string action, double actualPrice)
        {
            bool izmen = false;
            bool del = false;
            if (tovarList.Count == 0)
            {
                StreamWriter sw = new StreamWriter("badUrl.csv", true, Encoding.GetEncoding(1251));
                sw.WriteLine(urlTovar);
                sw.Close();
                return;
            }

            if (chekedEditMiniText)
            {
                string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";
                string nameBold = boldOpen + name + boldClose;
                string discount = discountTemplate;

                string miniText = minitextTemplate;
                miniText = miniText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace(" | РАЗДЕЛ", "").Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString()).Replace("<p><br /></p><p><br /></p><p><br /></p><p>", "<p><br /></p>");
                miniText = miniText.Remove(miniText.LastIndexOf("<p>"));
                tovarList[7] = miniText;

                // Обновление СЕО

                string descriptionText = null;
                string keywordsText = null;
                string titleText = null;

                titleText = tbTitle.Lines[0].ToString();
                descriptionText = tbDescription.Lines[0].ToString();
                keywordsText = tbKeywords.Lines[0].ToString();

                titleText = titleText.Replace("СКИДКА", discount).Replace(" | РАЗДЕЛ", "").Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                descriptionText = descriptionText.Replace("СКИДКА", discount).Replace(" | РАЗДЕЛ", "").Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                keywordsText = keywordsText.Replace("СКИДКА", discount).Replace(" | РАЗДЕЛ", "").Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

                titleText = textRemove(titleText, 255);
                descriptionText = textRemove(descriptionText, 200);
                keywordsText = textRemove(keywordsText, 100);

                tovarList[11] = descriptionText;
                tovarList[12] = keywordsText;
                tovarList[13] = titleText;

                // Обновление СЕО

                izmen = true;
            }

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
                    izmen = true;
                }
            }

            if (tovarList[39] != action)
            {
                tovarList[39] = action;
                izmen = true;
            }

            if (izmen & !del)
            {
                tovarList[42] = nethouse.alsoBuyTovars(tovarList);
                nethouse.SaveTovar(cookie, tovarList);
            }
        }

        private void UpdateTovarSnegohod()
        {
            ControlsFormEnabledFalse();
            CookieContainer cookie = nethouse.CookieNethouse(tbLogin.Text, tbPassword.Text);
            if (cookie.Count == 1)
            {
                MessageBox.Show("Логин или пароль для сайта введены не верно", "Ошибка логина/пароля");
                ControlsFormEnabledTrue();
                return;
            }

            File.Delete("naSite.csv");
            File.Delete("allTovars");
            nethouse.NewListUploadinBike18("naSite");
            List<string> newProduct = new List<string>();

            chekedEditMiniText = cbMiniText.Checked;

            FileInfo file = new FileInfo("Прайс-лист ТД Мегаполис 07.07.2017 Москва.xlsx");
            ExcelPackage p = new ExcelPackage(file);

            ExcelWorksheet w = p.Workbook.Worksheets[4];
            int q = w.Dimension.Rows;
            string razdelSnegohod = null;
            for (int i = 7; q > i; i++)
            {
                if (w.Cells[i, 1].Value == null)
                {
                    razdelSnegohod = (string)w.Cells[i, 2].Value;
                }
                else
                {
                    double articl;
                    try { articl = (double)w.Cells[i, 1].Value; }
                    catch
                    {
                        continue;
                    }
                    allTovarInFile(articl);
                    double quantity = (double)w.Cells[i, 9].Value;
                    double priceIrbisDiler = (double)w.Cells[i, 6].Value;
                    double actualPrice = Price(priceIrbisDiler);
                    string action = (string)w.Cells[i, 14].Value;
                    string name = (string)w.Cells[i, 3].Value;
                    name = name.Replace("\"", "");

                    ExcelRange er = w.Cells[i, 2];
                    DownloadImages(er, articl);

                    if (action != "")
                        action = actionText(action);

                    string urlTovar = nethouse.searchTovar(name, articl.ToString());
                    if (urlTovar == null)
                        urlTovar = nethouse.searchTovar(name, "IRB_" + articl.ToString());

                    if (urlTovar == "")
                    {
                        WriteInCSV(articl.ToString(), name, razdelSnegohod, actualPrice.ToString());
                    }
                    else
                    {
                        boldOpen = boldOpenSite;
                        List<string> tovarList = new List<string>();
                        tovarList = nethouse.GetProductList(cookie, urlTovar);

                        UpdateInfoTova(cookie, tovarList, urlTovar, name, articl, quantity, action, actualPrice);
                    }
                }
            }


            UploadTovarInSite(cookie);

            otv = httpRequest.getRequest("https://bike18.ru/products/category/zapchasti-dlya-snegohodov-i-motobuksirovshchikov");
            MatchCollection razdelSite = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            string[] allprod = File.ReadAllLines("allTovars");
            for (int i = 0; razdelSite.Count > i; i++)
            {
                otv = httpRequest.getRequest("https://bike18.ru" + razdelSite[i].ToString() + "?page=all");
                MatchCollection product = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Matches(otv);
                for (int n = 0; product.Count > n; n++)
                {
                    string urlTovar = product[n].ToString();
                    otv = httpRequest.getRequest(urlTovar);

                    DeleteTovarNonePrice(cookie, urlTovar, allprod);
                }
            }

            MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);
            ControlsFormEnabledTrue();
        }

        private void UpdateTovar()
        {
            ControlsFormEnabledFalse();
            CookieContainer cookie = nethouse.CookieNethouse(tbLogin.Text, tbPassword.Text);
            if (cookie.Count == 1)
            {
                MessageBox.Show("Логин или пароль для сайта введены не верно", "Ошибка логина/пароля");
                ControlsFormEnabledTrue();
                return;
            }

            File.Delete("naSite.csv");
            File.Delete("allTovars");
            nethouse.NewListUploadinBike18("naSite");
            List<string> newProduct = new List<string>();

            chekedEditMiniText = cbMiniText.Checked;

            FileInfo file = new FileInfo("Прайс-лист ТД Мегаполис 07.07.2017 Москва.xlsx");
            ExcelPackage p = new ExcelPackage(file);

            ExcelWorksheet w = p.Workbook.Worksheets[3];
            int q = w.Dimension.Rows;
            lblAll.Invoke(new Action(() => lblAll.Text = q.ToString()));

            #region Раздел запчасти
            for (int i = 8; q > i; i++)
            {
                lblProduct.Invoke(new Action(() => lblProduct.Text = i.ToString()));
                if (w.Cells[i, 1].Value == null)
                    break;
                double articl;
                try { articl = (double)w.Cells[i, 1].Value; }
                catch
                {
                    continue;
                }

                allTovarInFile(articl);
                double quantity = (double)w.Cells[i, 9].Value;
                double priceIrbisDiler;
                try { priceIrbisDiler = (double)w.Cells[i, 6].Value; }
                catch
                {
                    continue;
                }
                double actualPrice = Price(priceIrbisDiler);
                string action = (string)w.Cells[i, 14].Value;
                string name = (string)w.Cells[i, 3].Value;
                name = name.Replace("\"", "");

                ExcelRange er = w.Cells[i, 2];
                DownloadImages(er, articl);

                if (action != null)
                    action = actionText(action);
                else
                    action = "";

                string urlTovar = nethouse.searchTovar(name, articl.ToString());
                if (urlTovar == null)
                    urlTovar = nethouse.searchTovar(name, "IRB_" + articl.ToString());

                if (urlTovar == null)
                {
                    int space = name.IndexOf(" ");
                    string strRazdel = name.Remove(space, name.Length - space);
                    WriteInCSV(articl.ToString(), name, strRazdel, actualPrice.ToString());
                }
                else
                {
                    boldOpen = boldOpenSite;
                    List<string> tovarList = new List<string>();
                    tovarList = nethouse.GetProductList(cookie, urlTovar);

                    UpdateInfoTova(cookie, tovarList, urlTovar, name, articl, quantity, action, actualPrice);
                }
            }
            #endregion

            UploadTovarInSite(cookie);

            otv = httpRequest.getRequest("https://bike18.ru/products/category/katalog-zapchastey-irbis");
            MatchCollection razdelSite = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            string[] allprod = File.ReadAllLines("allTovars");
            for (int i = 0; razdelSite.Count > i; i++)
            {
                otv = httpRequest.getRequest("https://bike18.ru" + razdelSite[i].ToString() + "?page=all");
                MatchCollection product = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Matches(otv);
                for (int n = 0; product.Count > n; n++)
                {
                    string urlTovar = product[n].ToString();
                    otv = httpRequest.getRequest(urlTovar);

                    DeleteTovarNonePrice(cookie, urlTovar, allprod);
                }
            }

            otv = httpRequest.getRequest("https://bike18.ru/products/category/rashodniki-dlya-tehniki");
            razdelSite = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            for (int i = 0; razdelSite.Count > i; i++)
            {
                otv = httpRequest.getRequest("https://bike18.ru" + razdelSite[i].ToString() + "?page=all");
                MatchCollection product = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Matches(otv);
                for (int n = 0; product.Count > n; n++)
                {
                    string urlTovar = product[n].ToString();
                    otv = httpRequest.getRequest(urlTovar);

                    DeleteTovarNonePrice(cookie, urlTovar, allprod);
                }
            }
            MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);
            ControlsFormEnabledTrue();
        }

        #endregion

        private void WriteInCSV(string articl, string name, string razdel, string actualPrice)
        {
            boldOpen = boldOpenCSV;
            string slug = chpu.vozvr(name);
            razdel = ReturnRazdel(razdel);

            string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";

            string nameBold = boldOpen + name + boldClose;

            string miniText = minitextTemplate;
            string fullText = fullTextTemplate;
            string titleText = titleTextTemplate;
            string descriptionText = descriptionTextTemplate;
            string keywordsText = keywordsTextTemplate;
            string discount = discountTemplate;
            discount = discount.Replace("\"", "\"\"");

            miniText = miniText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace(" | РАЗДЕЛ", "").Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString()).Replace("<p><br /></p><p><br /></p><p><br /></p><p>", "<p><br /></p>");
            miniText = miniText.Remove(miniText.LastIndexOf("<p>"));

            fullText = fullText.Replace("СКИДКА", discount).Replace("ДУБЛЬ", dblProduct).Replace(" | РАЗДЕЛ", "").Replace("НАЗВАНИЕ", nameBold).Replace("АРТИКУЛ", articl.ToString());
            fullText = fullText.Remove(fullText.LastIndexOf("<p>"));

            titleText = titleText.Replace("СКИДКА", discount).Replace(" | РАЗДЕЛ", "").Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

            descriptionText = descriptionText.Replace("СКИДКА", discount).Replace(" | РАЗДЕЛ", "").Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

            keywordsText = keywordsText.Replace("СКИДКА", discount).Replace(" | РАЗДЕЛ", "").Replace("ДУБЛЬ", dblProduct).Replace("НАЗВАНИЕ", name).Replace("АРТИКУЛ", articl.ToString());

            titleText = textRemove(titleText, 255);
            descriptionText = textRemove(descriptionText, 200);
            keywordsText = textRemove(keywordsText, 100);
            slug = textRemove(slug, 64);

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

        private string ReturnRazdel(string razdel)
        {
            throw new NotImplementedException();
        }

        private void DownloadImages(ExcelRange er, double articl)
        {
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

        private string discountTemplateStr()
        {
            string discount = "<p style=\"text-align: right;\"><span style=\"font-weight: bold; font-weight: bold;\"> 1. <a href=\"https://bike18.ru/oplata-dostavka\">Выгодные условия доставки по всей России!</a></span></p><p style=\"text-align: right;\"><span style=\"font-weight: bold; font-weight: bold;\"> 2. <a href=\"https://bike18.ru/stock\">Нашли дешевле!? 110% разницы Ваши!</a></span></p><p style=\"text-align: right;\"><span style=\"font-weight: bold; font-weight: bold;\"> 3. <a href=\"https://bike18.ru/service\">Также обращайтесь в наш сервис центр в Ижевске!</a></span></p>";
            return discount;
        }

        private double Price(double priceDiler)
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

        private string irbisZapchastiRazdel(string strRazdel)
        {
            string razdel = "Запчасти и расходники => Каталог запчастей IRBIS => ";
            switch (strRazdel)
            {
                case "гусеницы":
                    razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => Гусеницы";
                    break;
                case "снегоходы Буран":
                    razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => Снегоходы Буран";
                    break;
                case "снегоходы Тайга":
                    razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => Снегоходы Тайга";
                    break;
                case "снегоходы Dingo":
                    razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => Снегоходы Dingo";
                    break;
                case "мотобуксировщики Мухтар":
                    razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => Мотобуксировщики Мухтар";
                    break;
                case "Аккумуляторная":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Аккумуляторы";
                    break;
                case "Амортизатор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Амортизаторы";
                    break;
                case "Амортизаторы":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Амортизаторы";
                    break;
                case "Багажник":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Багажники";
                    break;
                case "Бак":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Баки масляные, топливные, системы охлаждения";
                    break;
                case "Барабан":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Блоки переключения, бендиксы, барабаны";
                    break;
                case "Бачок":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Баки масляные, топливные, системы охлаждения";
                    break;
                case "Бендикс":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Блоки переключения, бендиксы, барабаны";
                    break;
                case "Блок":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Блоки переключения, бендиксы, барабаны";
                    break;
                case "Блоки":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Блоки переключения, бендиксы, барабаны";
                    break;
                case "Болты":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Болты, буксы";
                    break;
                case "Вал":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Валы";
                    break;
                case "Валы":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Валы";
                    break;
                case "Вариатор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Вариаторы";
                    break;
                case "Вентиль":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Вентили";
                    break;
                case "Вилка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Вилки переключения передач";
                    break;
                case "Втулка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Втулки";
                    break;
                case "Генератор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Генераторы в сборе";
                    break;
                case "Глушитель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Глушители";
                    break;
                case "Головка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Головки цилиндра";
                    break;
                case "грузики":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Грузики вариатора";
                    break;
                case "Датчик":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Датчики";
                    break;
                case "Двигатель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Двигатели в сборе";
                    break;
                case "Демпфер":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Демпферы";
                    break;
                case "Демпферные":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Демпферы";
                    break;
                case "Диск":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Колесные диски";
                    break;
                case "Диски":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Диски сцепления";
                    break;
                case "Жгут":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Жгуты проводов";
                    break;
                case "Замков":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Замки";
                    break;
                case "Замок":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Замки";
                    break;
                case "Защита":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Защита двигателя";
                    break;
                case "Звезда":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Звезды";
                    break;
                case "Зубчатый":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Зубчатые сектора";
                    break;
                case "Камера":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Камеры";
                    break;
                case "Индикатор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Индикаторы";
                    break;
                case "Карбюратор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Карбюраторы";
                    break;
                case "Катушка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Катушки зажигания";
                    break;
                case "Клапан":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Клапаны";
                    break;
                case "Клапаны":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Клапаны";
                    break;
                case "Клипса":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Клипсы";
                    break;
                case "Кнопка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Кнопки";
                    break;
                case "Кнопки":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Кнопки";
                    break;
                case "Кожух":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Кожухи";
                    break;
                case "Коллектор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Коллекторы";
                    break;
                case "Колодки":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Тормозные колодки";
                    break;
                case "Колпачок":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Свечные колпачки";
                    break;
                case "Кольца":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Кольца";
                    break;
                case "Кольцо":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Кольца";
                    break;
                case "Коммутатор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Коммутаторы";
                    break;
                case "Коромысла":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Коромысла";
                    break;
                case "Корпус":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Корпуса картеров, предохранителей";
                    break;
                case "Кран":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Топливные краны";
                    break;
                case "Крепление":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Крепления и кронштейны";
                    break;
                case "Кронштейн":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Крепления и кронштейны";
                    break;
                case "Крыло":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Крылья";
                    break;
                case "Крыльчатка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Крыльчатки";
                    break;
                case "Крышка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Крышки";
                    break;
                case "Лампа":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Лампы";
                    break;
                case "Машинка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Тормозные машинки";
                    break;
                case "Мембрана":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Мембраны карбюратора";
                    break;
                case "Муфта":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Обгонные муфты";
                    break;
                case "Наконечник":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Наконечники рулевой тяги";
                    break;
                case "Направляющие":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Направляющие цепи";
                    break;
                case "Насос":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Насосы";
                    break;
                case "Натяжитель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Натяжители цепи";
                    break;
                case "Обтекатели":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Обтекатели";
                    break;
                case "Обтекатель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Обтекатели";
                    break;
                case "Опора":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Опоры";
                    break;
                case "Ось":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Оси";
                    break;
                case "Палец":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Поршневые пальцы";
                    break;
                case "Панель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Панели приборов";
                    break;
                case "Патрубок":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Патрубки";
                    break;
                case "Педаль":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Педали тормоза";
                    break;
                case "Пластик":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Пластик";
                    break;
                case "Подножка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Подножки";
                    break;
                case "Подножки":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Подножки";
                    break;
                case "Подшипник":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Подшипники";
                    break;
                case "Подшипники":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Подшипники";
                    break;
                case "Поршневой":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Поршни";
                    break;
                case "Привод":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Приводы спидометра";
                    break;
                case "Прокладка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Прокладки";
                    break;
                case "Прокладки":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Прокладки";
                    break;
                case "Пружина":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Пружины";
                    break;
                case "Пружины":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Пружины";
                    break;
                case "Радиатор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Радиаторы";
                    break;
                case "Рама":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Рамы";
                    break;
                case "Реле":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Реле";
                    break;
                case "Реле-регулятор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Реле";
                    break;
                case "Ремень":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ремни вариатора";
                    break;
                case "Ремкомплект":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ремкомплекты карбюраторов";
                    break;
                case "Решетка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Решетки";
                    break;
                case "Ролик":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ролики натяжителя цепи";
                    break;
                case "Ротор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Роторы";
                    break;
                case "Руль":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Рули";
                    break;
                case "Ручка":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ручки, рычаги";
                    break;
                case "Ручки":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ручки, рычаги";
                    break;
                case "Рычаг":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ручки, рычаги";
                    break;
                case "Рычаги":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ручки, рычаги";
                    break;
                case "Сайлентблок":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Сайлентблоки, сальники";
                    break;
                case "Сальник":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Сайлентблоки, сальники";
                    break;
                case "Сальники":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Сайлентблоки, сальники";
                    break;
                case "Свеча":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Свечи зажигания";
                    break;
                case "Сигнал":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Звуковые сигналы";
                    break;
                case "Сиденье":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Сиденья";
                    break;
                case "Спица":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Спицы";
                    break;
                case "Стартер":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Статоры генератора";
                    break;
                case "Статор":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Статоры генератора";
                    break;
                case "Ступица":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Ступицы";
                    break;
                case "Суппорт":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Суппорты";
                    break;
                case "Сцепление":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Сцепление в сборе";
                    break;
                case "Толкатель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Толкатели";
                    break;
                case "Тормоз":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Тормоза";
                    break;
                case "Траверса":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Траверсы";
                    break;
                case "Трос":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Тросы";
                    break;
                case "Турбина":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Трубки, турбины";
                    break;
                case "Тяга":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Тяги";
                    break;
                case "Указатели":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Указатели поворотов";
                    break;
                case "Успокоитель":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Успокоители цепи";
                    break;
                case "Фара":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Фары";
                    break;
                case "Фильтр":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Фильтры";
                    break;
                case "Фильтрующий":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Фильтры";
                    break;
                case "Фонарь":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Фары";
                    break;
                case "Цапфа":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Цапфы";
                    break;
                case "Цепь":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Цепи";
                    break;
                case "Цилиндро-поршневая":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => ЦПГ";
                    break;
                case "Шестерни":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Шестерни и шайбы";
                    break;
                case "Шестерня":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Шестерни и шайбы";
                    break;
                case "Шина":
                    razdel = "Запчасти и расходники => Расходники для мототехники => Моторезина";
                    break;
                case "Шланг":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Шланги";
                    break;
                case "Электроклапан":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Электроклапана карбюраторов";
                    break;
                case "Электростартер":
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Электростартеры";
                    break;
                case "Защита для мототехники":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Защита для мототехники";
                    break;
                case "Лебедки":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Лебедки";
                    break;
                case "Стекла ветровые":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Стекла ветровые";
                    break;
                case "Зеркала":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Зеркала";
                    break;
                case "Кофры, сумки":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Кофры, сумки";
                    break;
                case "Канистры для мототехники":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Канистры для мототехники";
                    break;
                case "Цепи на колеса":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Цепи на колеса";
                    break;
                case "Электроника":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Электроника";
                    break;
                case "Оптика, LED лампы":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Оптика, LED лампы";
                    break;
                case "Рули, ручки руля, наконечники руля, защита":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Рули, ручки руля, наконечники руля, защита";
                    break;
                case "Стайлинг, рамки номеров гос. регистрации":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Стайлинг, рамки номеров гос. регистрации";
                    break;
                case "Проставки амортизаторов":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Проставки амортизаторов";
                    break;
                case "Чехлы":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Чехлы";
                    break;
                case "Противоугонные устройства":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Противоугонные устройства";
                    break;
                case "Декоративные метизы":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Декоративные метизы";
                    break;
                case "Инструменты":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Инструменты";
                    break;
                case "Литература":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Литература";
                    break;
                case "Наклейки":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Наклейки";
                    break;
                case "Зимние аксессуары":
                    razdel = "Аксессуары и инструменты => Аксессуары и инструменты VIRZ для мототехники => Зимние аксессуары";
                    break;
                default:
                    razdel = "Запчасти и расходники => Каталог запчастей IRBIS => Разное";
                    break;
            }
            return razdel;
        }

        private string MiniTextTemplateStr()
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

        private string FullTextTemplateStr()
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

        private void ControlsFormEnabledTrue()
        {
            btnActual.Invoke(new Action(() => btnActual.Enabled = true));
            button1.Invoke(new Action(() => button1.Enabled = true));
            btnSaveTemplates.Invoke(new Action(() => btnSaveTemplates.Enabled = true));
            btnUpdateImage.Invoke(new Action(() => btnUpdateImage.Enabled = true));
            rtbFullText.Invoke(new Action(() => rtbFullText.Enabled = true));
            rtbMiniText.Invoke(new Action(() => rtbMiniText.Enabled = true));
            tbDescription.Invoke(new Action(() => tbDescription.Enabled = true));
            tbKeywords.Invoke(new Action(() => tbKeywords.Enabled = true));
            tbLogin.Invoke(new Action(() => tbLogin.Enabled = true));
            tbPassword.Invoke(new Action(() => tbPassword.Enabled = true));
            tbTitle.Invoke(new Action(() => tbTitle.Enabled = true));
            cbMiniText.Invoke(new Action(() => cbMiniText.Enabled = true));
            btnAccessory.Invoke(new Action(() => btnAccessory.Enabled = true));
        }

        private void ControlsFormEnabledFalse()
        {
            btnActual.Invoke(new Action(() => btnActual.Enabled = false));
            button1.Invoke(new Action(() => button1.Enabled = false));
            btnSaveTemplates.Invoke(new Action(() => btnSaveTemplates.Enabled = false));
            btnUpdateImage.Invoke(new Action(() => btnUpdateImage.Enabled = false));
            rtbFullText.Invoke(new Action(() => rtbFullText.Enabled = false));
            rtbMiniText.Invoke(new Action(() => rtbMiniText.Enabled = false));
            tbDescription.Invoke(new Action(() => tbDescription.Enabled = false));
            tbKeywords.Invoke(new Action(() => tbKeywords.Enabled = false));
            tbLogin.Invoke(new Action(() => tbLogin.Enabled = false));
            tbPassword.Invoke(new Action(() => tbPassword.Enabled = false));
            tbTitle.Invoke(new Action(() => tbTitle.Enabled = false));
            cbMiniText.Invoke(new Action(() => cbMiniText.Enabled = false));
            btnAccessory.Invoke(new Action(() => btnAccessory.Enabled = false));
        }

        private void UploadTovarInSite(CookieContainer cookie)
        {
            System.Threading.Thread.Sleep(20000);
            string[] naSite1 = File.ReadAllLines("naSite.csv", Encoding.GetEncoding(1251));
            if (naSite1.Length > 1)
            {
                nethouse.UploadCSVNethouse(cookie, "naSite.csv");
            }
        }

        private void DeleteTovarNonePrice(CookieContainer cookie, string urlTovar, string[] allprod)
        {
            if (otv == "err")
            {
                StreamWriter s = new StreamWriter("badURL.txt", true);
                s.WriteLine(urlTovar);
                s.Close();
                return;
            }

            string artProd = new Regex("(?<=Артикул:)[\\w\\W]*?(?=</div)").Match(otv).ToString().Trim();
            bool b = false;
            string reg = new Regex("[0-9]{13}").Match(artProd).ToString();
            if (reg == "")
            {
                return;
            }

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
}