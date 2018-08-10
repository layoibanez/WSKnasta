using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using TLSharp.Core;
using TeleSharp.TL.Messages;
//using TeleSharp.TL;
using System.Net.Mail;
using Telegram.Bot;
using OpenQA.Selenium.Interactions;
using System.IO;
using System.Data;

using ClosedXML.Excel;
using System.Net.Mime;

namespace ConsoleApp1
{
    class Program
    {
        public static  void Main(string[] args)
        {
                      


            ILog log = log4net.LogManager.GetLogger(typeof(Program));
            try
            {


                log.Info("Inicia navegador. v2");

                // Initialize the Chrome Driver
                ChromeOptions opt = new ChromeOptions();
                opt.AddArguments("--start-maximized");
                int descuentoAverificar = 39;
                string mailAEnviar = "layolaif@gmail.com";


                for (int m = 0; m < 7; m++)
                {
                    using (var driver = new ChromeDriver(opt))
                    {

                        string nombreHoja = string.Empty;
                        switch (m)
                        {
                            case 0:
                                nombreHoja = "Tecnologia";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=0&ktegory=2000&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            case 1:
                                nombreHoja = "Electrodomesticos";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=0&ktegory=4000&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            case 2:
                                nombreHoja = "MueblesDecoJardin";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=0&ktegory=5000&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            case 3:
                                nombreHoja = "Deportes";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=0&ktegory=7000&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            case 4:
                                nombreHoja = "Moda";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=0&ktegory=8000&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            case 5:
                                nombreHoja = "General";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=3&ktegory=0&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            case 6:
                                nombreHoja = "Otros";
                                driver.Navigate().GoToUrl("https://knasta.cl/results/?knastaday=0&ktegory=1000&max_price=&min_price=&order=asc&page=1&partners=all");
                                break;
                            default:
                                Console.WriteLine("Default case");
                                break;
                        }



                        DataTable dt = GetDataTableStruct();


                        IList<IWebElement> divsOffer = driver.FindElementsByClassName("ProductList-item");
                        IList<IWebElement> footerPaginas = driver.FindElementByClassName("Pagination-buttons").FindElements(By.XPath(".//*"));
                        int contadorpaginas = footerPaginas.Count;
                        for (int w = 0; w < contadorpaginas - 3; w++)
                        {
                            Console.WriteLine("Pagina " + w);
                            if (w != 0)
                            {

                                Actions act3 = new Actions(driver);
                                act3.MoveToElement(driver.FindElementByClassName("Pagination-buttons"));
                                act3.Perform();


                                driver.FindElementByClassName("Pagination-buttons").FindElements(By.XPath(".//*"))[w].Click();
                            }

                            for (int i = 1; i < divsOffer.Count; i++)
                            {

                                IList<IWebElement> divsContenedor = driver.FindElementsByClassName("ProductBox-content");
                                IWebElement divContenedor = divsContenedor[i];
                                if (divContenedor != null)
                                {
                                    int descuento = int.Parse(divContenedor.FindElements(By.ClassName("ProductBox-diff"))[0].Text.Replace("%", ""));

                                    if (descuento < (descuentoAverificar * -1))
                                    {
                                        Console.WriteLine("descuento a verificar");
                                        string prodcuto = divContenedor.FindElements(By.ClassName("ProductBox-title"))[0].Text;
                                        string precio = divContenedor.FindElements(By.ClassName("ProductBox-price"))[0].Text;

                                        //  string message = prodcuto + " a: " + precio + " con un: " + (descuento * - 1).ToString() + "% de descuento";


                                        Actions act2 = new Actions(driver);
                                        act2.MoveToElement(divContenedor);
                                        act2.Perform();



                                        bool chargedcont = true;
                                        while (chargedcont)
                                        {
                                            try
                                            {


                                                Console.WriteLine("click en contenedor producto " + prodcuto);
                                                divContenedor.Click();


                                                chargedcont = false;
                                            }
                                            catch (Exception ex3)
                                            {
                                                log.Error(ex3);
                                                log.Info("==================================================================");
                                            }

                                        }

                                        Console.WriteLine("espera a cargar contenedor producto");
                                        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(120));

                                        Actions actions = new Actions(driver);
                                        Console.WriteLine("scrol al inicio");
                                        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(120));
                                        chargedcont = false;
                                        bool charged = false;
                                        try {
                                            IWebElement element = driver.FindElementByClassName("ProductView-chartTitle");
                                            Actions act = new Actions(driver);
                                            act.MoveToElement(element);
                                            act.Perform();

                                            Console.WriteLine("carga pagina");
                                            charged = true;
                                        }
                                        catch (Exception ex3)
                                        {
                                           
                                        }
                                        
                                     
                                        while (charged)
                                        {
                                            try
                                            {
                                                if (driver.FindElementsByClassName("HistoryRow-price").Count <= 2 && driver.FindElementsByClassName("HistoryRow-price").Count > 0)
                                                    charged = false;


                                                if (charged)
                                                    driver.FindElementByClassName("HistoryTableDesktop-displayButton").Click();


                                                charged = false;
                                            }
                                            catch (Exception ex2) { }

                                        }
                                        Console.WriteLine("calculo de oferta");
                                        double sumaPrecios = 0;

                                        //Se leen los puntos del grafico
                                        try {
                                            IList<IWebElement> listaPrecios = driver.FindElementsByTagName("circle");
                                            if (listaPrecios.Count > 0)
                                            {
                                                try
                                                {
                                                    for (int j = 0; j < listaPrecios.Count - 2; j++)
                                                    {

                                                        sumaPrecios = sumaPrecios + double.Parse((driver.FindElementsByTagName("circle")[j].GetAttribute("cy").Replace(".", ",")));
                                                    }

                                                }
                                                catch (Exception ex2) { }


                                                double promedio = 0;
                                                    
                                                try
                                                {
                                                    promedio = sumaPrecios / listaPrecios.Count - 2;
                                                }
                                                catch (Exception ex2) { }


                                                try
                                                {
                                                    double descuentoReal = 0;
                                                    try
                                                    {
                                                        descuentoReal = ((double.Parse((listaPrecios[listaPrecios.Count - 1].GetAttribute("cy").Replace(".", ","))) * 100) / promedio) - 100;
                                                    }
                                                    catch (Exception ex2) { }


                                                    if (descuentoReal >= descuentoAverificar)
                                                    {
                                                        string linkTienda = string.Empty;

                                                        try
                                                        {
                                                            IWebElement buybutton = driver.FindElementByClassName("BuyButton-button");
                                                            linkTienda = buybutton.GetAttribute("href");
                                                        }
                                                        catch (Exception ex2) { }


                                                        DataRow drNewRow = dt.NewRow();

                                                        drNewRow["Nombre"] = prodcuto;
                                                        drNewRow["Precio"] = precio;
                                                        drNewRow["Descuento"] = descuento.ToString();

                                                        drNewRow["Link Knasta"] = driver.Url;
                                                        drNewRow["Link Tienda"] = linkTienda;
                                                        dt.Rows.Add(drNewRow);
                                                        //Enviar mensaje
                                                        //sendmail(precio, descuento.ToString(), prodcuto, mailAEnviar, buybutton.GetAttribute("href"), driver.Url);

                                                    }

                                                }
                                                catch (Exception ex2) { }


                                            }
                                        }
                                        catch (Exception ex3)
                                        {
                                            //Sigue con los demas
                                        }

                                      

                                        try { 
                                        driver.Navigate().Back();
                                        }
                                        catch (Exception ex2) { }

                                    }
                                }
                            }
                        }


                        XLWorkbook wb = new XLWorkbook();
                        wb.Worksheets.Add(dt, nombreHoja);
                        sendmail(mailAEnviar, wb, nombreHoja);
                        driver.Close();
                    }
                }

              
                        

              Environment.Exit(0);
                

            }
           catch (Exception ex)
            {
                 log.Error(ex);
                log.Info("==================================================================");

            }


        }
         
    
        public static DataTable GetDataTableStruct()
        {
            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn();
            dc.ColumnName = "Nombre";
            dt.Columns.Add(dc);
            dc = new DataColumn();
            dc.ColumnName = "Precio";
            dt.Columns.Add(dc);
            dc = new DataColumn();
            dc.ColumnName = "Descuento";
            dt.Columns.Add(dc);
          
            dc = new DataColumn();
            dc.ColumnName = "Link Knasta";
            dt.Columns.Add(dc);
            dc = new DataColumn();
            dc.ColumnName = "Link Tienda";
            dt.Columns.Add(dc);
            return dt;
        }

        public static bool sendmail(string mailAEnviar, XLWorkbook wb, string subject)//(string precio, string descuento, string producto, string mailAEnviar, string url, string urlactual)
        {
            SmtpClient server = new SmtpClient("smtp.gmail.com", 587);
            server.Credentials = new System.Net.NetworkCredential("knastalaio@gmail.com", "123knastalaio123");
            server.EnableSsl = true;
            // string mensaje = producto + " a: " + precio + "  con un %" + descuento + " de descuento. Link: " + url + " Url Knasta : " + urlactual;
            string mensaje = "Ofertas";
            try
            {


              

                MailMessage mnsj = new MailMessage();


                // Attachment data = new Attachment(wb, MediaTypeNames.Application.Octet);
                //ContentDisposition disposition = data.ContentDisposition;
                //   disposition.CreationDate = System.IO.File.GetCreationTime(file);
                // disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                //disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                // Add the file attachment to this e-mail message.
                //mnsj.Attachments.Add(data);
                                    

                using (var memoryStream = new MemoryStream())
                {
                    wb.SaveAs(memoryStream, false);
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    var attachment = new Attachment(memoryStream,
                                                        "Ofertas_" + DateTime.Now +
                                                        ".xlsx", MediaTypeNames.Application.Octet);
                    mnsj.Attachments.Add(attachment);
                    //Create the SMTP client object and send the message
                    //var smtpClient = new SmtpClient(smtpServer);
                    //smtpClient.Send(mnsj);
                    mnsj.Subject = subject;
                    mnsj.To.Add(new MailAddress(mailAEnviar));
                    mnsj.From = new MailAddress("knastalaio@gmail.com", "El oferton");

                    mnsj.Body = mensaje;

                    /* Enviar */
                    server.Send(mnsj);
                    memoryStream.Close();
                    
                }
             


                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        //public static async void sendmessage(string precio, string descuento, string producto, string mailAEnviar)
        //{

        //    string message = producto + " " + precio + " " + descuento;
    
        //    try
        //    {

        //        var bot = new Telegram.Bot.TelegramBotClient("548196997:AAFzKwNGaaO6ls7tABHpnugZ6DKvO_A4O5E");
        //        await bot.SendTextMessageAsync("@knastalaio", message);
        //        //  bot.OnMessage += Bot_OnMessage;

        //        Console.WriteLine("Envio correcto para: " + message);


        //    }
        //    catch(Exception ex)
        //    {
        //        Console.WriteLine("Error sending: " + message);
        //    }
           
        //}

        //private static void Bot_OnMessage(object sender, Telegram.Bot.Args.MessageEventArgs e)
        //{
        //    bot.SendTextMessageAsync(e.Message.Chat.Id, "chao");
            
        //}
    }
}
