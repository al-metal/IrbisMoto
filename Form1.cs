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
        string otv = null;
        int deleteTovar = 0;
        int editPrice = 0;
        WebClient webClient = new WebClient();
        FileEdit files = new FileEdit();

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
            File.Delete("naSite.csv");

            List<string> newProduct = new List<string>();
            newProduct.Add("id");                                                                               //id
            newProduct.Add("Артикул *");                                                 //артикул
            newProduct.Add("Название товара *");                                          //название
            newProduct.Add("Стоимость товара *");                                    //стоимость
            newProduct.Add("Стоимость со скидкой");                                       //со скидкой
            newProduct.Add("Раздел товара *");                                         //раздел товара
            newProduct.Add("Товар в наличии *");                                                    //в наличии
            newProduct.Add("Поставка под заказ *");                                                 //поставка
            newProduct.Add("Срок поставки (дни) *");                                           //срок поставки
            newProduct.Add("Краткий текст");                                 //краткий текст
            newProduct.Add("Текст полностью");                                          //полностью текст
            newProduct.Add("Заголовок страницы (title)");                               //заголовок страницы
            newProduct.Add("Описание страницы (description)");                                 //описание
            newProduct.Add("Ключевые слова страницы (keywords)");                                 //ключевые слова
            newProduct.Add("ЧПУ страницы (slug)");                                   //ЧПУ
            newProduct.Add("С этим товаром покупают");                              //с этим товаром покупают
            newProduct.Add("Рекламные метки");
            newProduct.Add("Показывать на сайте *");                                           //показывать
            newProduct.Add("Удалить *");                                    //удалить
            files.fileWriterCSV(newProduct, "naSite");

            FileInfo file = new FileInfo("Прайс.xlsx");
            ExcelPackage p = new ExcelPackage(file);
            ExcelWorksheet w = p.Workbook.Worksheets[3];
            int q = w.Dimension.Rows;

            for (int i = 8; q > i; i++)
            {
                if (w.Cells[i, 1].Value == null)
                    break;

                double articl = (double)w.Cells[i, 1].Value;
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
                urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
                List<string> tovarList = webRequest.arraySaveimage(urlTovar);

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

                string boldOpen = "<span style=\"font-weight: bold; font-weight: bold; \">";
                string boldClose = "</span>";

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

                if (quantity == 0)
                {
                    if (urlTovar != "")
                    {
                        if (action != "")
                        {
                            tovarList[39] = action;
                            tovarList[9] = actualPrice.ToString();
                            tovarList[1] = slug;
                            tovarList[7] = miniText;
                            tovarList[8] = fullText;
                            tovarList[11] = descriptionText;
                            tovarList[12] = keywordsText;
                            tovarList[13] = titleText;
                            tovarList[3] = "10833347";
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
                    if (urlTovar != "")
                    {
                        double priceBike18 = Convert.ToDouble(tovarList[9].ToString());
                        if (actualPrice != priceBike18)
                        {
                            tovarList[39] = action;
                            tovarList[9] = actualPrice.ToString();
                            tovarList[1] = slug;
                            tovarList[7] = miniText;
                            tovarList[8] = fullText;
                            tovarList[11] = descriptionText;
                            tovarList[12] = keywordsText;
                            tovarList[13] = titleText;
                            tovarList[3] = "10833347";
                            webRequest.saveImage(tovarList);
                            editPrice++;
                        }
                        else if (tovarList[39].ToString() != action)
                        {
                            tovarList[39] = action;
                            tovarList[1] = slug;
                            tovarList[7] = miniText;
                            tovarList[8] = fullText;
                            tovarList[11] = descriptionText;
                            tovarList[12] = keywordsText;
                            tovarList[13] = titleText;
                            tovarList[3] = "10833347";
                            webRequest.saveImage(tovarList);
                            editPrice++;
                        }
                        else
                        {
                            tovarList[1] = slug;
                            tovarList[7] = miniText;
                            tovarList[8] = fullText;
                            tovarList[11] = descriptionText;
                            tovarList[12] = keywordsText;
                            tovarList[13] = titleText;
                            tovarList[3] = "10833347";
                            webRequest.saveImage(tovarList);
                            editPrice++;
                        }
                    }
                    else
                    {
                        string stock = (string)w.Cells[i, 14].Value;
                        bool kioshi = name.Contains("KIYOSHI");
                        if (!kioshi)
                        {
                            newProduct = new List<string>();
                            newProduct.Add(""); //id
                            newProduct.Add("\"" + articl + "\""); //артикул
                            newProduct.Add("\"" + name + "\"");  //название
                            newProduct.Add("\"" + actualPrice + "\""); //стоимость
                            newProduct.Add("\"" + "" + "\""); //со скидкой
                            newProduct.Add("\"" + razdel + "\""); //раздел товара
                            newProduct.Add("\"" + "100" + "\""); //в наличии
                            newProduct.Add("\"" + "0" + "\"");//поставка
                            newProduct.Add("\"" + "1" + "\"");//срок поставки
                            newProduct.Add("\"" + miniText + "\"");//краткий текст
                            newProduct.Add("\"" + fullText + "\"");//полностью текст
                            newProduct.Add("\"" + titleText + "\""); //заголовок страницы
                            newProduct.Add("\"" + descriptionText + "\""); //описание
                            newProduct.Add("\"" + keywordsText + "\"");//ключевые слова
                            newProduct.Add("\"" + slug + "\""); //ЧПУ
                            newProduct.Add(""); //с этим товаром покупают
                            newProduct.Add("");   //рекламные метки
                            newProduct.Add("\"" + "1" + "\"");  //показывать
                            newProduct.Add("\"" + "0" + "\""); //удалить

                            files.fileWriterCSV(newProduct, "naSite");
                        }
                    }
                }
            }

            w = p.Workbook.Worksheets[4];
            q = w.Dimension.Rows;
            string razdelSnegohod = null;
            string podrazdel = null;
            for (int i = 7; q > i; i++)
            {
                if (w.Cells[i, 1].Value == null)
                {
                    razdelSnegohod = (string)w.Cells[i, 2].Value;
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
                }
                else
                {
                    double articl = (double)w.Cells[i, 1].Value;
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
                    }

                    otv = webRequest.getRequest("http://bike18.ru/products/search/page/1?sort=0&balance=&categoryId=&min_cost=&max_cost=&text=" + articl);
                    string urlTovar = new Regex("(?<=<a href=\").*(?=\"><div class=\"-relative item-image\")").Match(otv).ToString();
                    urlTovar = urlTovar.Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
                    List<string> tovarList = webRequest.arraySaveimage(urlTovar);

                    string slug = chpu.vozvr(name);
                    string razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => " + podrazdel;


                    string miniText = null;
                    string titleText = null;
                    string descriptionText = null;
                    string keywordsText = null;
                    string fullText = null;
                    string discount = null;
                    string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";

                    string boldOpen = "<span style=\"font-weight: bold; font-weight: bold; \">";
                    string boldClose = "</span>";

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

                    if (quantity == 0)
                    {
                        if (urlTovar != "")
                        { 
                            if (action != "")
                            {
                                tovarList[39] = action;
                                tovarList[9] = actualPrice.ToString();
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
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
                        if (urlTovar != "")
                        {
                            double priceBike18 = Convert.ToDouble(tovarList[9].ToString());
                            if (actualPrice != priceBike18)
                            {
                                tovarList[39] = action;
                                tovarList[9] = actualPrice.ToString();
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                            else if (tovarList[39].ToString() != action)
                            {
                                tovarList[39] = action;
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                            else
                            {
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                        }
                        else
                        {                           
                            string stock = (string)w.Cells[i, 14].Value;

                            newProduct = new List<string>();
                            newProduct.Add(""); //id
                            newProduct.Add("\"" + articl + "\""); //артикул
                            newProduct.Add("\"" + name + "\"");  //название
                            newProduct.Add("\"" + actualPrice + "\""); //стоимость
                            newProduct.Add("\"" + "" + "\""); //со скидкой
                            newProduct.Add("\"" + razdel + "\""); //раздел товара
                            newProduct.Add("\"" + "100" + "\""); //в наличии
                            newProduct.Add("\"" + "0" + "\"");//поставка
                            newProduct.Add("\"" + "1" + "\"");//срок поставки
                            newProduct.Add("\"" + miniText + "\"");//краткий текст
                            newProduct.Add("\"" + fullText + "\"");//полностью текст
                            newProduct.Add("\"" + titleText + "\""); //заголовок страницы
                            newProduct.Add("\"" + descriptionText + "\""); //описание
                            newProduct.Add("\"" + keywordsText + "\"");//ключевые слова
                            newProduct.Add("\"" + slug + "\""); //ЧПУ
                            newProduct.Add(""); //с этим товаром покупают
                            newProduct.Add("");   //рекламные метки
                            newProduct.Add("\"" + "1" + "\"");  //показывать
                            newProduct.Add("\"" + "0" + "\""); //удалить

                            files.fileWriterCSV(newProduct, "naSite");
                        }
                    }
                }
            }

            string razdelkiyoshi = null;
            podrazdel = null;
            w = p.Workbook.Worksheets[5];
            q = w.Dimension.Rows;

            for (int i = 7; q > i; i++)
            {
                if (w.Cells[i, 2].Value == null)
                {
                    razdelkiyoshi = (string)w.Cells[i, 1].Value;

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

                }
                else
                {

                    double articl = (double)w.Cells[i, 2].Value;
                    double quantity = (double)w.Cells[i, 10].Value;
                    double priceIrbisDiler = (double)w.Cells[i, 7].Value;
                    double actualPrice = Price(priceIrbisDiler);
                    string action = (string)w.Cells[i, 14].Value;
                    string name = (string)w.Cells[i, 3].Value;
                    name = name.Replace("\"", "").Replace("\n", "");

                    string slug = chpu.vozvr(name);
                    string razdel = "Запчасти и расходники => Запчасти для снегоходов и мотобуксировщиков => " + podrazdel;


                    string miniText = null;
                    string titleText = null;
                    string descriptionText = null;
                    string keywordsText = null;
                    string fullText = null;
                    string discount = null;
                    string dblProduct = "НАЗВАНИЕ также подходит для: аналогичных моделей.";

                    string boldOpen = "<span style=\"font-weight: bold; font-weight: bold; \">";
                    string boldClose = "</span>";

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
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
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
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                            else if (tovarList[39].ToString() != action)
                            {
                                tovarList[39] = action;
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                            else
                            {
                                tovarList[1] = slug;
                                tovarList[7] = miniText;
                                tovarList[8] = fullText;
                                tovarList[11] = descriptionText;
                                tovarList[12] = keywordsText;
                                tovarList[13] = titleText;
                                tovarList[3] = "10833347";
                                webRequest.saveImage(tovarList);
                                editPrice++;
                            }
                        }
                        else
                        {
                            string stock = (string)w.Cells[i, 14].Value;
                            
                            newProduct = new List<string>();
                            newProduct.Add(""); //id
                            newProduct.Add("\"" + articl + "\""); //артикул
                            newProduct.Add("\"" + name + "\"");  //название
                            newProduct.Add("\"" + actualPrice + "\""); //стоимость
                            newProduct.Add("\"" + "" + "\""); //со скидкой
                            newProduct.Add("\"" + razdel + "\""); //раздел товара
                            newProduct.Add("\"" + "100" + "\""); //в наличии
                            newProduct.Add("\"" + "0" + "\"");//поставка
                            newProduct.Add("\"" + "1" + "\"");//срок поставки
                            newProduct.Add("\"" + miniText + "\"");//краткий текст
                            newProduct.Add("\"" + fullText + "\"");//полностью текст
                            newProduct.Add("\"" + titleText + "\""); //заголовок страницы
                            newProduct.Add("\"" + descriptionText + "\""); //описание
                            newProduct.Add("\"" + keywordsText + "\"");//ключевые слова
                            newProduct.Add("\"" + slug + "\""); //ЧПУ
                            newProduct.Add(""); //с этим товаром покупают
                            newProduct.Add("");   //рекламные метки
                            newProduct.Add("\"" + "1" + "\"");  //показывать
                            newProduct.Add("\"" + "0" + "\""); //удалить

                            files.fileWriterCSV(newProduct, "naSite");
                        }
                    }
                }


            }

            System.Threading.Thread.Sleep(20000);
            string trueOtv = null;
            string[] naSite1 = File.ReadAllLines("naSite.csv", Encoding.GetEncoding(1251));
            if (naSite1.Length > 1)
            {
                do
                {
                    string otvimg = DownloadNaSite();
                    string check = "{\"success\":true,\"imports\":{\"state\":1,\"errorCode\":0,\"errorLine\":0}}";
                    do
                    {
                        System.Threading.Thread.Sleep(2000);
                        otvimg = ChekedLoading();
                    }
                    while (otvimg == check);

                    trueOtv = new Regex("(?<=\":{\"state\":).*?(?=,\")").Match(otvimg).ToString();
                    string error = new Regex("(?<=errorCode\":).*?(?=,\")").Match(otvimg).ToString();
                    if (error == "13")
                    {
                        string errstr = new Regex("(?<=errorLine\":).*?(?=,\")").Match(otvimg).ToString();
                        string[] naSite = File.ReadAllLines("naSite.csv", Encoding.GetEncoding(1251));
                        int u = Convert.ToInt32(errstr) - 1;
                        string[] strslug3 = naSite[u].Split(';');
                        int slugint = strslug3.Length - 5;
                        string strslug = strslug3[slugint].ToString();
                        int slug = strslug.Length;
                        string strslug2 = strslug.Remove(slug - 2);
                        strslug2 += "1";
                        naSite[u] = naSite[u].Replace(strslug, strslug2);
                        File.WriteAllLines("naSite.csv", naSite, Encoding.GetEncoding(1251));
                    }
                    if (error == "37")
                    {
                        string errstr = new Regex("(?<=errorLine\":).*?(?=,\")").Match(otvimg).ToString();
                        string[] naSite = File.ReadAllLines("naSite.csv", Encoding.GetEncoding(1251));
                        int u = Convert.ToInt32(errstr) - 1;
                        string[] strslug3 = naSite[u].Split(';');
                        int slugint = strslug3.Length - 5;
                        string strslug = strslug3[slugint].ToString();
                        int slug = strslug.Length;
                        string strslug2 = strslug.Remove(slug - 2);
                        strslug2 += "1";
                        naSite[u] = naSite[u].Replace(strslug, strslug2);
                        File.WriteAllLines("naSite.csv", naSite, Encoding.GetEncoding(1251));
                    }
                }
                while (trueOtv != "2");
            }

            MessageBox.Show("Удалено " + deleteTovar + " позиций товара\n " + "Отредактировано цен на товары " + editPrice);
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
            string disount = "<p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> Сделай ТРОЙНОЙ удар по нашим ценам! </span></p><p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> 1. <a target=\"\"_blank\"\" href =\"\"http://bike18.ru/stock\"\"> Скидки за отзывы о товарах!</a> </span></p><p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> 2. <a target=\"\"_blank\"\" href =\"\"http://bike18.ru/stock\"\"> Друзьям скидки и подарки!</a> </span></p><p style=\"\"text-align: right;\"\"><span style=\"\"font -weight: bold; font-weight: bold;\"\"> 3. <a target=\"\"_blank\"\" href =\"\"http://bike18.ru/stock\"\"> Нашли дешевле!? 110% разницы Ваши!</a></span></p>";
            return disount;
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

        public string DownloadNaSite()
        {
            CookieContainer cookie = webRequest.webCookieBike18();
            string epoch = (DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds.ToString().Replace(",", "");
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://bike18.nethouse.ru/api/export-import/import-from-csv?fileapi" + epoch);
            req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:44.0) Gecko/20100101 Firefox/44.0";
            req.Method = "POST";
            req.ContentType = "multipart/form-data; boundary=---------------------------12709277337355";
            req.CookieContainer = cookie;
            req.Headers.Add("X-Requested-With", "XMLHttpRequest");
            byte[] csv = File.ReadAllBytes("naSite.csv");
            byte[] end = Encoding.ASCII.GetBytes("\r\n-----------------------------12709277337355\r\nContent-Disposition: form-data; name=\"_catalog_file\"\r\n\r\nnaSite.csv\r\n-----------------------------12709277337355--\r\n");
            byte[] ms1 = Encoding.ASCII.GetBytes("-----------------------------12709277337355\r\nContent-Disposition: form-data; name=\"catalog_file\"; filename=\"naSite.csv\"\r\nContent-Type: text/csv\r\n\r\n");
            req.ContentLength = ms1.Length + csv.Length + end.Length;
            Stream stre1 = req.GetRequestStream();
            stre1.Write(ms1, 0, ms1.Length);
            stre1.Write(csv, 0, csv.Length);
            stre1.Write(end, 0, end.Length);
            stre1.Close();
            HttpWebResponse resimg = (HttpWebResponse)req.GetResponse();
            StreamReader ressrImg = new StreamReader(resimg.GetResponseStream());
            string otvimg = ressrImg.ReadToEnd();
            return otvimg;
        }

        public string ChekedLoading()
        {
            CookieContainer cookie = webRequest.webCookieBike18();
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://bike18.nethouse.ru/api/export-import/check-import");
            req.Accept = "application/json, text/plain, */*";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:44.0) Gecko/20100101 Firefox/44.0";
            req.Method = "POST";
            req.ContentLength = 0;
            req.ContentType = "application/x-www-form-urlencoded";
            req.CookieContainer = cookie;
            Stream stre1 = req.GetRequestStream();
            stre1.Close();
            HttpWebResponse resimg = (HttpWebResponse)req.GetResponse();
            StreamReader ressrImg = new StreamReader(resimg.GetResponseStream());
            string otvimg = ressrImg.ReadToEnd();
            return otvimg;
        }

        private void btnUpdateImage_Click(object sender, EventArgs e)
        {
            int countUpdateImage = 0;
            otv = webRequest.getRequest("http://bike18.ru/products/category/1183836");
            MatchCollection razdel = new Regex("(?<=<div class=\"category-capt-txt -text-center\"><a href=\").*?(?=\" class=\"blue\">)").Matches(otv);
            for(int i = 0; razdel.Count > i; i++)
            {
                otv = webRequest.getRequest(razdel[i].ToString() + "/page/all");
                MatchCollection tovar = new Regex("(?<=<div class=\"product-link -text-center\"><a href=\").*?(?=\" >)").Matches(otv);
                for(int n = 0; tovar.Count > n; n++)
                {
                    otv = webRequest.getRequest(tovar[n].ToString());
                    string urlImageTovar = new Regex("(?<=class=\"avatar-view \"><link rel=\"image_src\" href=\").*?(?=\">)").Match(otv).ToString();
                    if(urlImageTovar == "")
                    {
                        string articl = new Regex("(?<= Артикул:)[\\w\\W]*?(?=</div>)").Match(otv).ToString();
                        articl = articl.Trim();
                        if(File.Exists("pic\\" + articl + ".jpg"))
                        {
                            CookieContainer cookie = webRequest.webCookieBike18();
                            string urlTovar = tovar[n].ToString().Replace("http://bike18.ru/", "http://bike18.nethouse.ru/");
                            MatchCollection prId = new Regex("(?<=data-id=\").*?(?=\")").Matches(otv);
                            int prodId = Convert.ToInt32(prId[0].ToString());
                            otv = webRequest.PostRequest(urlTovar);

                            Image newImg = Image.FromFile("pic\\" + articl + ".jpg");
                            double widthImg = newImg.Width;
                            double heigthImg = newImg.Height;
                            if (widthImg > heigthImg)
                            {
                                double dblx = widthImg * 0.9;
                                if (dblx < heigthImg)
                                {
                                    heigthImg = heigthImg * 0.9;
                                }
                                else
                                    widthImg = widthImg * 0.9;
                            }
                            else
                            {
                                double dblx = heigthImg * 0.9;
                                if (dblx < widthImg)
                                {
                                    widthImg = widthImg * 0.9;
                                }
                                else
                                    heigthImg = heigthImg * 0.9;
                            }

                            string otvimg = DownloadImages(articl);
                            string urlSaveImg = new Regex("(?<=url\":\").*?(?=\")").Match(otvimg).Value.Replace("\\/", "%2F");
                            string otvSave = SaveImages(urlSaveImg, prodId, widthImg, heigthImg);
                            List<string> listProd = webRequest.arraySaveimage(urlTovar);
                            listProd[3] = "10833347";
                            webRequest.saveImage(listProd);
                            countUpdateImage++;
                        }
                    }
                }
            }

            otv = webRequest.getRequest("http://bike18.ru/products/category/1289775");

            otv = webRequest.getRequest("http://bike18.ru/products/category/2182755");

            MessageBox.Show("Обновлено картинок: " + countUpdateImage);
        }

        public string DownloadImages(string artProd)
        {
            CookieContainer cookie = webRequest.webCookieBike18();
            string epoch = (DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds.ToString().Replace(",", "");
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://bike18.nethouse.ru/putimg?fileapi" + epoch);
            req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:44.0) Gecko/20100101 Firefox/44.0";
            req.Method = "POST";
            req.ContentType = "multipart/form-data; boundary=---------------------------12709277337355";
            req.CookieContainer = cookie;
            req.Headers.Add("X-Requested-With", "XMLHttpRequest");
            byte[] pic = File.ReadAllBytes("Pic\\" + artProd + ".jpg");
            byte[] end = Encoding.ASCII.GetBytes("\r\n-----------------------------12709277337355\r\nContent-Disposition: form-data; name=\"_file\"\r\n\r\n" + artProd + ".jpg\r\n-----------------------------12709277337355--\r\n");
            byte[] ms1 = Encoding.ASCII.GetBytes("-----------------------------12709277337355\r\nContent-Disposition: form-data; name=\"file\"; filename=\"" + artProd + ".jpg\"\r\nContent-Type: image/jpeg\r\n\r\n");
            req.ContentLength = ms1.Length + pic.Length + end.Length;
            Stream stre1 = req.GetRequestStream();
            stre1.Write(ms1, 0, ms1.Length);
            stre1.Write(pic, 0, pic.Length);
            stre1.Write(end, 0, end.Length);
            stre1.Close();
            HttpWebResponse resimg = (HttpWebResponse)req.GetResponse();
            StreamReader ressrImg = new StreamReader(resimg.GetResponseStream());
            string otvimg = ressrImg.ReadToEnd();
            return otvimg;
        }

        public string SaveImages(string urlSaveImg, int prodId, double widthImg, double heigthImg)
        {
            CookieContainer cookie = webRequest.webCookieBike18();
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create("http://bike18.nethouse.ru/api/catalog/save-image");
            req.Accept = "application/json, text/plain, */*";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:44.0) Gecko/20100101 Firefox/44.0";
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            req.CookieContainer = cookie;
            byte[] saveImg = Encoding.ASCII.GetBytes("url=" + urlSaveImg + "&id=0&type=4&objectId=" + prodId + "&imgCrop[x]=0&imgCrop[y]=0&imgCrop[width]=" + widthImg + "&imgCrop[height]=" + heigthImg + "&imageId=0&iObjectId=" + prodId + "&iImageType=4&replacePhoto=0");
            req.ContentLength = saveImg.Length;
            Stream srSave = req.GetRequestStream();
            srSave.Write(saveImg, 0, saveImg.Length);
            srSave.Close();
            HttpWebResponse resSave = (HttpWebResponse)req.GetResponse();
            StreamReader ressrSave = new StreamReader(resSave.GetResponseStream());
            string otvSave = ressrSave.ReadToEnd();
            return otvSave;
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

    }
}
