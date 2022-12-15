using System;
using Aspose.Words;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V105.Input;
using OpenQA.Selenium.DevTools.V105.SystemInfo;
using OpenQA.Selenium.Support.UI;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Reflection.Metadata;
using Document = Aspose.Words.Document;

namespace webscraping
{
    class Program
    {
        public static void Main(string[] args)
        {
            bool starten = true;
            while (starten)
            {
                Console.WriteLine("Welke webscrape optie kies je? \n1 ->  scrapen van data op Youtube door zelfgekozen zoekterm \n2 ->  scrapen van data op ICTjobs door zelfgekozen zoekterm \n3 ->  scrapen van data op Bol.com door zelfgekozen zoekterm\n4 ->  scrapen van data op Google Flights door zelfgekozen zoekterm");
                int optie = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Je hebt gekozen voor optie " + optie);

                if (optie == 1)
                {
                    Console.WriteLine("Op welke zoekterm wil je op Youtube zoeken?");
                    string inputGebruiker = Console.ReadLine();
        
                    IWebDriver driver = new ChromeDriver(@"C:\Users\Kyara\source\repos\DevOpsCaseStudy\bin\Debug\net6.0\chromedriver.exe");
                    driver.Navigate().GoToUrl("https://www.youtube.com");
               
                    // cookies accepteren
                    var submitButton = driver.FindElement(By.XPath("/html/body/ytd-app/ytd-consent-bump-v2-lightbox/tp-yt-paper-dialog/div[4]/div[2]/div[6]/div[1]/ytd-button-renderer[2]/yt-button-shape/button/yt-touch-feedback-shape/div/div[2]"));
                    submitButton.Click();
                    Thread.Sleep(5000);

                    // input invullen in zoekbalk
                    var element = driver.FindElement(By.XPath("/html/body/ytd-app/div/div[1]/ytd-masthead/div[3]/div[2]/ytd-searchbox/form/div[1]/div[1]/input"));
                    element.SendKeys(inputGebruiker);
                    element.Submit();

                    Thread.Sleep(4000);
                    // alle videos zoeken
                    By elem_video_links = By.CssSelector("#contents > ytd-video-renderer");
                    ReadOnlyCollection<IWebElement> videos = driver.FindElements(elem_video_links);
                    Console.WriteLine("The 5 most recently uploaded videos are:");
                    
                    //csv 
                    StringBuilder csvcontent = new StringBuilder();
                    csvcontent.AppendLine("Title, Views, Uploader, Link");

                    /* Voor elke video in de videos lijst haal hiervan informatie*/
                    int count = 1;
                    foreach (IWebElement video in videos)
                    {
                        if (count < 6)
                        {
                            string str_title, str_views, str_uploader, str_link;

                            IWebElement elem_video_title = video.FindElement(By.CssSelector("#video-title"));
                            str_title = elem_video_title.Text;
                    
                            IWebElement elem_video_views = video.FindElement(By.CssSelector("#metadata-line > span:nth-child(3)"));
                            str_views = elem_video_views.Text;

                            IWebElement elem_video_uploader = video.FindElement(By.CssSelector("#text > a"));
                            str_uploader = elem_video_uploader.GetAttribute("innerText");

                            IWebElement elem_video_link = video.FindElement(By.CssSelector("#thumbnail"));
                            str_link = elem_video_link.GetAttribute("href");

                            Console.WriteLine("*** Video " + count + " ***");
                            Console.WriteLine("Video Title: " + str_title);
                            Console.WriteLine("Video Views: " + str_views);
                            Console.WriteLine("Video Uploader: " + str_uploader);
                            Console.WriteLine("Video Link: " + str_link + "\n");
                    
                            // csv data aanmaken
                            csvcontent.AppendLine(str_title + "," + str_views + "," + str_uploader + "," + str_link);
                            //JSON OBJECT
                            JObject videoyoutube = new JObject(
                                new JProperty("Title ", str_title),
                                new JProperty("Views ", str_views),
                                new JProperty ("Uploader ", str_uploader),
                                new JProperty("Link ", str_link));
                            File.AppendAllText(@"C:\Users\Kyara\source\repos\DevopsCaseStudy\fileOptie1YoutubeVideos.json", videoyoutube.ToString());

                            count++;
                        }             
                    }
                    string csvpath = @"C:\Users\Kyara\source\repos\DevOpsCaseStudy\fileOptie1YoutubeVideos.csv";
                    File.AppendAllText(csvpath, csvcontent.ToString());
                    Console.WriteLine("DONE");
                }

                else if (optie == 2)
                {
                    Console.WriteLine("Op welke zoekterm wil je op de website van ictjob.be zoeken?");
                    string zoekterm = Console.ReadLine();
                    
                    IWebDriver driver = new ChromeDriver(@"C:\Users\Kyara\source\repos\DevOpsCaseStudy\bin\Debug\net6.0\chromedriver.exe");
                    driver.Url = "https://www.ictjob.be/nl/";
                    Thread.Sleep(1000);

                    // filter optie
                    Console.WriteLine("Op wat wil je zoeken?\n1-ten minste één van deze woorden\n2-alle woorden\n3-exact zoeken\nGeef alleen de nummer in van de optie die je kiest");
                    int zoekoptie = Convert.ToInt32(Console.ReadLine());
                    if(zoekoptie == 2){
                        var zoekfilter = driver.FindElement(By.CssSelector("#column-keywords-options > label:nth-child(2) > span"));
                        zoekfilter.Click();
                    }
                    else if(zoekoptie == 3){
                        var zoekfilter = driver.FindElement(By.CssSelector("#column-keywords-options > label:nth-child(3) > span"));
                        zoekfilter.Click();
                    }
                    // zoekterm van de gebruiker invullen en submitten
                    driver.FindElement(By.Id("keywords-input")).SendKeys(zoekterm);
                    driver.FindElement(By.Id("main-search-button")).Submit();

                    // filteren op de vacatures
                    Console.WriteLine("Wil je op de job functies nog verder filteren? ");
                    string antwoord = Console.ReadLine();
                    if (antwoord == "JA" || antwoord == "ja" || antwoord == "Ja")
                    {
                        
                        Console.WriteLine("Contracttype\nWelk contracttype wil je?\n1-Onbepaalde duur\n2-Freelance\n3-Bepaalde duur");
                        int number = Convert.ToInt32(Console.ReadLine());
                        Thread.Sleep(2000);
                        if (number == 1)
                        {
                            var filterContractoptie1 = driver.FindElement(By.CssSelector("#search-criteria-display > div > div.main-panel > div > div > ul > li:nth-child(2) > div > ul > li:nth-child(1) > label > a"));
                            filterContractoptie1.Click();
                        }
                        else if (number ==2)
                        {
                            var filterContractoptie2 = driver.FindElement(By.CssSelector("#search-criteria-display > div > div.main-panel > div > div > ul > li:nth-child(2) > div > ul > li:nth-child(2) > label > a"));
                            filterContractoptie2.Click();
                        }
                        else if (number == 3)
                        {
                            var filterContractoptie3 = driver.FindElement(By.CssSelector("#search-criteria-display > div > div.main-panel > div > div > ul > li:nth-child(2) > div > ul > li:nth-child(3) > label > a"));
                            filterContractoptie3.Click();
                        }
                        else
                        {
                            Console.WriteLine("Je hebt geen optie gegeven tussen 1-3.");
                        }

                        Console.WriteLine("Welk taal van de job wil je?");
                        Console.WriteLine("Wil je Nederlands als taal van de job? (typ ja)");
                        string taal_nederlands = Console.ReadLine();
                        Console.WriteLine("Wil je Engels als taal van de job? (typ ja)");
                        string taal_engels = Console.ReadLine();
                        Console.WriteLine("Wil je Frans als taal van de job? (typ ja)");
                        string taal_frans = Console.ReadLine();

                        if (taal_nederlands == "ja" || taal_nederlands == "JA" || taal_nederlands == "Ja")
                        {
                            var filtertaal = driver.FindElement(By.CssSelector("#search-criteria-display > div > div.main-panel > div > div > ul > li:nth-child(1) > div > ul > li:nth-child(1) > label > a"));
                            filtertaal.Click();
                        }
                        if (taal_engels == "ja" || taal_engels == "JA" || taal_engels == "Ja")
                        {
                            var filtertaal = driver.FindElement(By.CssSelector("#search-criteria-display > div > div.main-panel > div > div > ul > li:nth-child(1) > div > ul > li:nth-child(2) > label > a"));
                            filtertaal.Click();
                        }
                        if (taal_frans == "ja" || taal_frans == "JA" || taal_frans == "Ja")
                        {
                            var filtertaal = driver.FindElement(By.CssSelector("#search-criteria-display > div > div.main-panel > div > div > ul > li:nth-child(1) > div > ul > li:nth-child(3) > label > a"));
                            filtertaal.Click();
                        }
                        Thread.Sleep(2000);
                    }
                    Thread.Sleep(10000);
                    // alle vacatures ophalen
                    By elem_job_links = By.CssSelector("#search-result-body > div > ul > li> span.job-info");
                    ReadOnlyCollection<IWebElement> jobs = driver.FindElements(elem_job_links);
                    Console.WriteLine("The 5 most recently jobs are:");

                    int count = 1;
                    // csv file aanmaken + eerste record aanmaken
                    StringBuilder csvcontent = new StringBuilder();
                    csvcontent.AppendLine("Title, Bedrijf, Locatie, Keywords, Link, Datum");

                    // voor elke job in de list van jobs haal je informatie als de count kleiner is dan 6.
                    foreach (IWebElement job in jobs)
                    {
                        if (count < 6) { 
                            string str_title, str_bedrijf, str_locatie, str_keywords, str_link, str_date;
                            IWebElement elem_job_title = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > a > h2"));
                            str_title = elem_job_title.Text;
                    
                            IWebElement elem_job_bedrijf = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > span.job-company"));
                            str_bedrijf = elem_job_bedrijf.Text;

                            IWebElement elem_job_locatie = job.FindElement(By.CssSelector("#search-result-body > div > ul > li> span.job-info > span.job-details > span.job-location > span > span"));
                            str_locatie = elem_job_locatie.Text;

                            IWebElement elem_job_keywords = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > span.job-keywords"));
                            str_keywords = elem_job_keywords.Text;

                            IWebElement elem_job_link = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > a"));
                            str_link = elem_job_link.GetAttribute("href");
                      
                            IWebElement elem_job_date = job.FindElement(By.CssSelector("#search-result-body > div > ul > li > span.job-info > span.job-details > span.job-date"));
                            str_date = elem_job_date.Text;

                            Console.WriteLine("*** Job " + count + " ***");
                            Console.WriteLine("Title: " + str_title);
                            Console.WriteLine("Bedrijf: " + str_bedrijf);
                            Console.WriteLine("Locatie: " + str_locatie);
                            Console.WriteLine("Keywords: " + str_keywords);
                            Console.WriteLine("Link: " + str_link);
                            Console.WriteLine("Date: " + str_date + "\n");

                            // csv data aanmaken
                            csvcontent.AppendLine(str_title + "," + str_bedrijf + "," + str_locatie + "," + str_keywords + "," + str_link + "," + str_date);
                            //JSON OBJECT
                            JObject jobsfile = new JObject(
                                new JProperty("Job Functie", str_title),
                                new JProperty("Bedrijf", str_bedrijf),
                                new JProperty ("Locatie", str_locatie),
                                new JProperty("Keywords", str_keywords),
                                new JProperty("Link", str_link),
                                new JProperty("Date",str_date));
                        
                            File.AppendAllText(@"C:\Users\Kyara\source\repos\DevopsCaseStudy\fileOptie2Jobs.json", jobsfile.ToString());
                            count++;
                        } 
                    
                    }
                    //csv file aanmaken
                    string csvpath = @"C:\Users\Kyara\source\repos\DevOpsCaseStudy\fileOptie2Jobs.csv";
                    File.WriteAllText(csvpath, csvcontent.ToString());
                    Console.WriteLine("DONE");

                }
                else if (optie == 3)
                {
                    Console.WriteLine("Op welke zoekterm wil je op bol.com zoeken?");
                    string zoekterm = Console.ReadLine();

                    IWebDriver driver = new ChromeDriver(@"C:\Users\Kyara\source\repos\DevOpsCaseStudy\bin\Debug\net6.0\chromedriver.exe");
                    driver.Url = "https://www.bol.com/be/nl/";

                    Thread.Sleep(1000);
                    var button = driver.FindElement(By.XPath("//*[@id=\"js-first-screen-accept-all-button\"]"));       
                    button.Submit();
                
                    // taalkeuze
                    Console.WriteLine("Taal keuze \n1-Nederlands\n2-Français \nGeef de nummer van de taal in die je wil gebruiken. ");
                    int languagenumber = Convert.ToInt32(Console.ReadLine());
                    if (languagenumber == 1)
                    {
                        var element_taal = driver.FindElement(By.CssSelector("#modalWindow > div.modal__window.js_modal_window > div.modal__window__content.js_modal_content > wsp-country-language-modal > div:nth-child(8) > p.u-mb--xs.u-pt--xs.c-media.c-media--center > label > span"));
                        element_taal.Click();
                        Console.WriteLine("U hebt gekozen voor Nederlands");
                    }
                    else if (languagenumber == 2)
                    {
                        var element_taal = driver.FindElement(By.CssSelector("#modalWindow > div.modal__window.js_modal_window > div.modal__window__content.js_modal_content > wsp-country-language-modal > div:nth-child(8) > p.u-mb--0.u-pb--xs.u-pt--xs.c-media.c-media--center > label > span"));
                        element_taal.Click();
                        Console.WriteLine("U hebt gekozen voor Français");
                    }
                    else
                    {
                        Console.WriteLine("Je hebt geen optie van 1-2 gekozen dus gaat de taal verder in het Nederlands.");
                    }

                    Thread.Sleep(1000);
                    // button van de taal submitten
                    var buttonland = driver.FindElement(By.XPath("//*[@id=\"modalWindow\"]/div[2]/div[2]/wsp-country-language-modal/button"));
                    buttonland.Submit();

                    // input gebruiker invullen en op de eerste optie klikken
                    Thread.Sleep(1000);
                    var element = driver.FindElement(By.XPath("//*[@id=\"searchfor\"]"));
                    element.SendKeys(zoekterm);

                    Thread.Sleep(5000);
                    var eersteoptie = driver.FindElement(By.CssSelector("#siteSearch"));
                    eersteoptie.Submit();

                    // lijst maken van alle resultaten
                    By elem_resultaten = By.CssSelector("#js_items_content > li");
                    ReadOnlyCollection<IWebElement> resultaten = driver.FindElements(elem_resultaten);
                    Console.WriteLine("The 5 results of the search input: ");

                    // csv file aanmaken + eerste record aanmaken
                    StringBuilder csvcontent = new StringBuilder();
                    csvcontent.AppendLine("Title, Price, Description, Link");
                    int count = 1;

                    // voor elke resultaat informatie halen en tonenen als de count kleiner is dan 6.
                    foreach (IWebElement resultaat in resultaten)
                    {
                        if (count < 6)
                        {
                            string str_title, str_price, str_description, str_link;

                            IWebElement elem_resultaat_title = resultaat.FindElement(By.CssSelector("#js_items_content > li > div.product-item__content > div > div.product-title--inline > wsp-analytics-tracking-event > a"));
                            str_title = elem_resultaat_title.Text;
                    

                            IWebElement elem_resultaat_price = resultaat.FindElement(By.CssSelector("#js_items_content > li > div.product-item__content > wsp-buy-block > div.product-prices > section > div > div"));
                            str_price = elem_resultaat_price.Text;

                            IWebElement elem_resultaat_description = resultaat.FindElement(By.CssSelector("#js_items_content > li > div.product-item__content > div > p"));
                            str_description = elem_resultaat_description.Text;

                            IWebElement elem_resultaat_link = resultaat.FindElement(By.CssSelector("#js_items_content > li > div.product-item__content > div > div.product-title--inline > wsp-analytics-tracking-event > a"));
                            str_link = elem_resultaat_link.GetAttribute("href");

                            Console.WriteLine("*** Resultaat " + count + " ***");
                            Console.WriteLine("Title: " + str_title);
                            Console.WriteLine("Price: " + str_price);
                            Console.WriteLine("Description: " + str_description);
                            Console.WriteLine("Link: " + str_link + "\n");
                    
                            // csv data aanmaken
                            csvcontent.AppendLine(str_title + "," + str_price + "," + str_description + "," + str_link);
                            //JSON OBJECT
                            JObject resultatenoptiedrie = new JObject(
                                new JProperty("Title ", str_title),
                                new JProperty("Price ", str_price),
                                new JProperty ("Description ", str_description),
                                new JProperty("Link ", str_link));
                            File.AppendAllText(@"C:\Users\Kyara\source\repos\DevopsCaseStudy\fileOptie3Bol.json", resultatenoptiedrie.ToString());
                        
                            count++;
                        }
                    }
                    //csv file aanmaken
                    string csvpath = @"C:\Users\Kyara\source\repos\DevOpsCaseStudy\fileOptie3Bol.csv";
                    File.WriteAllText(csvpath, csvcontent.ToString());
                    Console.WriteLine("DONE");
                }
                else if (optie == 4)
                {
                    Console.WriteLine("Op welke zoekterm wil je op Google Flights zoeken?");
                    string inputGebruiker = Console.ReadLine();
                    
                    IWebDriver driver = new ChromeDriver(@"C:\Users\Kyara\source\repos\DevOpsCaseStudy\bin\Debug\net6.0\chromedriver.exe");
                    driver.Navigate().GoToUrl("https://www.google.com/travel/flights");

                    // cookies accepteren
                    var element = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz > div > div > div > div.NIoIEf > div.G4njw > div.AIC7ge > div.CxJub > div.VtwTSb > form:nth-child(2) > div > div > button"));
                    element.Click();
                    Thread.Sleep(4000);

                    // input clicken
                    var inputclick = driver.FindElement(By.CssSelector("#i15 > div.e5F5td.vxNK6d > div"));
                    inputclick.Click();
                    
                    //input invullen
                    var search = driver.FindElement(By.XPath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[6]/div[2]/div[2]/div[1]/div/input"));
                    search.SendKeys(inputGebruiker);
                    Thread.Sleep(4000);

                    // eerste link klikken
                    var link = driver.FindElement(By.CssSelector("#c51 > div.CwL3Ec"));
                    link.Click();

                    // zoeken button vluchten
                    var searchflights = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.vg4Z0e > div:nth-child(1) > div.SS6Dqf.CH4bwe > div.MXvFbd > div > button"));
                    searchflights.Click();
                    
                    Thread.Sleep(5000);

                   // alle resultaten verzamelen
                    By elem_resultaten_heenreis = By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li");
                    ReadOnlyCollection<IWebElement> resultaten = driver.FindElements(elem_resultaten_heenreis);
                    Console.WriteLine("The " + resultaten.Count()  +" results of the best flights to " + inputGebruiker + ": ");
                    int count = 0;
                    // voor alle resultaten informatie halen en tonen
                    foreach (IWebElement resultaat in resultaten)
                    {
                         string str_title, str_duur, str_tussenstop, str_prijs, str_uitstoot;
                         IWebElement elem_resultaat_title = resultaat.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li > div > div.yR1fYc > div > div.OgQvJf.nKlB3b > div.Ir0Voe > div.sSHqwe.tPgKwe.ogfYpf > span:nth-child(1)"));
                         str_title = elem_resultaat_title.Text;

                         IWebElement elem_resultaat_duur = resultaat.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li > div > div.yR1fYc > div > div.OgQvJf.nKlB3b > div.Ak5kof > div"));
                         str_duur = elem_resultaat_duur.Text;

                        IWebElement elem_resultaat_tussenstop = resultaat.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li > div > div.yR1fYc > div > div.OgQvJf.nKlB3b > div.BbR8Ec > div.EfT7Ae.AdWm1c.tPgKwe > span"));
                        str_tussenstop = elem_resultaat_tussenstop.Text;

                        IWebElement elem_resultaat_prijs = resultaat.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li > div > div.yR1fYc > div > div.OgQvJf.nKlB3b > div.U3gSDe > div.BVAVmf.I11szd.POX3ye > div.YMlIz.FpEdX > span"));
                        str_prijs = elem_resultaat_prijs.GetAttribute("aria-label");
                        
                        IWebElement elem_resultaat_uitstoot = resultaat.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li> div > div.yR1fYc > div > div.OgQvJf.nKlB3b > div.y0NSEe.V1iAHe.tPgKwe.ogfYpf > div > div.O7CXue"));
                        str_uitstoot = elem_resultaat_uitstoot.Text;

                        count++;
                        Console.WriteLine("*** Vlucht " + count + " ***");
                        Console.WriteLine("Vliegmaatschappij: " + str_title);
                        Console.WriteLine("Reistijd: " + str_duur);
                        Console.WriteLine("Tussenstops: " + str_tussenstop);
                        Console.WriteLine("Prijs heenreis: " + str_prijs);
                        Console.WriteLine("Uitstoot: " + str_uitstoot + "\n");

                    }
                    Console.WriteLine("Welke optie van 1 tot " + resultaten.Count() + " kies je?");
                    int keuze = Convert.ToInt32(Console.ReadLine());
      
                    while (keuze < 0 || keuze > resultaten.Count())
                    {
                        Console.WriteLine("Je hebt geen optie gegeven van 1 tot " + resultaten.Count() + "\nWelke optie van 1 tot " + resultaten.Count() + " kies je?");
                        int new_keuze = Convert.ToInt32(Console.ReadLine());
                        keuze = new_keuze;

                    }
                    // De keuze die je hebt gemaakt op deze klikken
                    var element_keuze = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li:nth-child(" + keuze + ")"));
                    element_keuze.Click();

                    Thread.Sleep(2000);
                    //terugreis
                    

                    Thread.Sleep(5000);
                    // Alle terugvluchten verzamelen en informatie tonen
                    By elem_resultaten_terugreis = By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li");
                    ReadOnlyCollection<IWebElement> resultaten_terugreis = driver.FindElements(elem_resultaten_terugreis);
                    int resultaten_terugreis_count = resultaten_terugreis.Count();
                    Console.WriteLine("The " + resultaten_terugreis_count + " results of the best flights back to Brussels: ");
                    int count_terugreis = 1;
                    foreach (IWebElement resultaat_reis in resultaten_terugreis)
                    {
                            string str_title, str_reistijd,str_tussenstops, str_prijs, str_uitstoot;
                            IWebElement elem_resultaat_terugreis_title = resultaat_reis.FindElement(By.XPath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[2]/div[3]/ul/li["+count_terugreis+"]/div/div[2]/div/div[2]/div[2]/div[2]"));
                            str_title = elem_resultaat_terugreis_title.Text;

                                                                                                           
                            IWebElement elem_resultaat_terugreis_reistijd = resultaat_reis.FindElement(By.XPath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[2]/div[3]/ul/li["+count_terugreis+"]/div/div[2]/div/div[2]/div[3]/div"));
                            str_reistijd = elem_resultaat_terugreis_reistijd.Text;

                            IWebElement elem_resultaat_terugreis_tussenstops = resultaat_reis.FindElement(By.XPath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[2]/div[3]/ul/li["+count_terugreis+"]/div/div[2]/div/div[2]/div[4]/div[1]/span"));
                            str_tussenstops = elem_resultaat_terugreis_tussenstops.Text;

                            IWebElement elem_resultaat_terugreis_prijs = resultaat_reis.FindElement(By.XPath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[2]/div[3]/ul/li["+count_terugreis+"]/div/div[2]/div/div[2]/div[6]/div[1]/div[2]/span"));
                            str_prijs = elem_resultaat_terugreis_prijs.GetAttribute("aria-label");

                            IWebElement elem_resultaat_terugreis_uitstoot = resultaat_reis.FindElement(By.XPath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[2]/div[3]/ul/li["+count_terugreis+"]/div/div[2]/div/div[2]/div[5]/div/div[1]"));
                            str_uitstoot = elem_resultaat_terugreis_uitstoot.Text;

                            Console.WriteLine("*** Vlucht " + count_terugreis + " ***");
                            Console.WriteLine("Vliegmaatschappij: " + str_title);
                            Console.WriteLine("Reistijd: " + str_reistijd);
                            Console.WriteLine("Tussenstops: " + str_tussenstops);
                            Console.WriteLine("Prijs terugreis: " + str_prijs);
                            Console.WriteLine("Uitstoot: " + str_uitstoot + "\n");
                        count_terugreis++;              
                    }
                    // Keuze maken en op klikken
                    Console.WriteLine("Welke optie van 1 tot " + resultaten_terugreis_count + " kies je?");
                    int keuze_terugreis = Convert.ToInt32(Console.ReadLine());

                    while (keuze_terugreis < 0 || keuze_terugreis > resultaten_terugreis.Count())
                    {
                        Console.WriteLine("Je hebt geen optie gegeven van 1 tot " + resultaten_terugreis_count + "\nWelke optie van 1 tot " + resultaten_terugreis_count + " kies je?");
                        int new_keuze = Convert.ToInt32(Console.ReadLine());
                        keuze_terugreis = new_keuze;

                    }
                    var element_keuze_terugreis = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div.PSZ8D.EA71Tc > div.FXkZv > div:nth-child(4) > ul > li:nth-child("+keuze_terugreis+")"));
                    element_keuze_terugreis.Click();

                    Thread.Sleep(5000);
                    //eind pagina + overzicht maken van de vlucht
                    var totale_prijs = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div > div.SDUAh.Xag90b.jtr7Nd > div.OLfz3c > div.Z9HR2d > div > div > div.Y4xNqf > div > div.QORQHb > span"));
                    var heenvlucht = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div > div.SDUAh.Xag90b.jtr7Nd > div.OLfz3c > div:nth-child(4) > div > div:nth-child(2) > div.rVD9dd > div > div > div > div:nth-child(1) > div.mz0jqb > div > div.KC3CM.zeBrcf > div > div.OgQvJf.nKlB3b > div.Ir0Voe > div.zxVSec.YMlIz.tPgKwe.ogfYpf > span"));
                    var terugvlucht = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div > div.SDUAh.Xag90b.jtr7Nd > div.OLfz3c > div:nth-child(4) > div > div:nth-child(2) > div.rVD9dd > div > div > div > div:nth-child(2) > div.mz0jqb > div > div.KC3CM.zeBrcf > div > div.OgQvJf.nKlB3b > div.Ir0Voe > div.zxVSec.YMlIz.tPgKwe.ogfYpf > span"));
                    var vergelijken = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div > div.SDUAh.Xag90b.jtr7Nd > div.OLfz3c > div:nth-child(4) > div > div.rzE7s > div > span.YMlIz.szUogf"));
                    var vergelijken_tekst = driver.FindElement(By.CssSelector("#yDmH0d > c-wiz.zQTmif.SSPGKf > div > div:nth-child(2) > c-wiz > div.cKvRXe > c-wiz > div > div.SDUAh.Xag90b.jtr7Nd > div.OLfz3c > div:nth-child(4) > div > div.rzE7s > div > div.ysCyjb > div.eoY5cb.nNYSH > div"));

                    Console.WriteLine("Overzicht: \n" + heenvlucht.GetAttribute("aria-label"));
                    Console.WriteLine(terugvlucht.GetAttribute("aria-label"));
                    Console.WriteLine("De reis van Brussel tot " + inputGebruiker + " is vanaf " + totale_prijs.GetAttribute("aria-label")+ " voor 1 passagier.");
                    Console.WriteLine("Vergelijken:\n"+vergelijken.Text + "\n" + vergelijken_tekst.Text);
                    
                    //Word doc maken + speciferen van eigenschappen
                    Document doc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    Font font = builder.Font;
                    font.Size = 12;
                    font.Bold = false;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Calibri";
                    // Tekst in document toevoegen
                    builder.Writeln(heenvlucht.GetAttribute("aria-label") + "\n" + terugvlucht.GetAttribute("aria-label"));
                    doc.Save(@"C:\Users\Kyara\source\repos\DevOpsCaseStudy\Optie4vluchten.docx");
                }
                else
                {
                    Console.WriteLine("Je hebt geen optie tussen 1 of 4 gegeven.");
                }
                Console.WriteLine("Wilt u nog eens opnieuw doen? (Type 'ja')");
                string doorgaan = Console.ReadLine();
                if (doorgaan == "JA" || doorgaan == "Ja" || doorgaan == "ja" || doorgaan == "j")
                {
                    starten = true;
                }
                else
                {
                    starten = false;
                }
            }
           
        }
    }
}

