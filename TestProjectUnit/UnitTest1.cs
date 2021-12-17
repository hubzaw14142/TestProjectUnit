using CsvHelper;
using NUnit.Framework;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;


namespace TestProjectUnit
{
    public class Tests
    {
        public IWebDriver driver;

        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--ignore-certificate-errors");
            options.AddArgument("--ignore-ssl-errors");
            options.AcceptInsecureCertificates = true;

            driver = new ChromeDriver(options);
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://isqueue-dev.pl/");

            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(7);
            driver.Manage().Timeouts().PageLoad = System.TimeSpan.FromSeconds(4);
            driver.Manage().Window.Maximize();

        }

        public void Zmiana_filii(int nr_filii)
        {
            string path = "/html/body/header/div[2]/section/a[2]";
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(3);
            IWebElement wstecz = driver.FindElement(By.XPath(path));
            wstecz.Click();

            nr_filii++;
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(3);
            string xpath11 = "/html/body/main/div[3]/div[2]/div/section[" + nr_filii + "]/div";
            IWebElement filia2 = driver.FindElement(By.XPath(xpath11));
            filia2.Click();
        }

        public void Sprawa_click(int nr_sprawy)
        {
            string xpath2 = " /html/body/main/div/div[2]/div/div[2]/div[" + nr_sprawy + "]/div";
            IWebElement sprawa = driver.FindElement(By.XPath(xpath2));
            sprawa.Click();
            nr_sprawy++;
        }

        public void Wybieranie_sprawy(int nr_filii, int nr_sprawy)
        {

            int sprawy = 0;
            IList<IWebElement> ilosc_spraw = driver.FindElements(By.XPath("/html/body/main/div/div[2]/div[4]/div[2]"));
            sprawy = ilosc_spraw.Count();
            try
            {
                Sprawa_click(nr_sprawy);
            }
            catch
            {
                if (nr_sprawy == sprawy)
                {
                    Zmiana_filii(nr_filii);
                }
                else
                {
                    string path = "/html/body/header/div[2]/section/a[2]";
                    IWebElement wstecz = driver.FindElement(By.XPath(path));
                    wstecz.Click();
                }
            }

            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(3);
        }

       

        public void Formularz(int pominiete_osoby)
        {
            string sciezka = "C:/Users/hzawisza/Desktop/TestProject/Zg³oszenia.csv";

            string line = File.ReadLines(sciezka).Skip(pominiete_osoby).Take(1).First();
            var wlasnosci = line.Split(';');
            User osoba = new User() { Imie = wlasnosci[0], Nazwisko = wlasnosci[1], Telefon = wlasnosci[2], Email = wlasnosci[3], };

            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(3);
            string imie = "/html/body/main/div/div[1]/div[2]/div[1]/input";
            driver.FindElement(By.XPath(imie)).SendKeys(osoba.Imie);
            string nazwisko = "/html/body/main/div/div[1]/div[2]/div[2]/input";
            driver.FindElement(By.XPath(nazwisko)).SendKeys(osoba.Nazwisko);
            string email = "/html/body/main/div/div[1]/div[2]/div[3]/input";
            driver.FindElement(By.XPath(email)).SendKeys(osoba.Email);
            string telefon = "/html/body/main/div/div[1]/div[2]/div[4]/input";
            driver.FindElement(By.XPath(telefon)).SendKeys(osoba.Telefon);
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(2);

            //Przycisk rezerwuj
            string rezerwuj = "/html/body/main/div/div[5]/button";
            driver.FindElement(By.XPath(rezerwuj)).Click();
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(10);

        }

        public void Wybierz_date(string data)
        {
            string path_calendar = "/html/body/main/div/div[2]/div/div[1]/div[1]";
            driver.FindElement(By.XPath(path_calendar)).Click();
            IList<IWebElement> days = driver.FindElements(By.ClassName("day"));
            days.SingleOrDefault(o => o.Text == data).Click();
                 
        }

        public void Wybieranie_godziny(string godzina)
        {
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(5);

            IList<IWebElement> hours= driver.FindElements(By.ClassName("avaiable-hour text-center d-flex"));
            
            foreach(var item in hours)
            {
                if (item.Displayed)
                {
                    hours.Remove(item);
                }
            }

            hours.SingleOrDefault(o => o.Text == godzina).Click();
        }

        [Test]
        public void Test1()
        {
            //Wybieranie filii
            int nr_filii = 1;
            int nr_sprawy = 1;
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(5);
            string xpath1 = "/html/body/main/div[3]/div[2]/div/section[1]/div";

            IWebElement filia = driver.FindElement(By.XPath(xpath1));
            filia.Click();

            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(5);
            string data = "22";
            Wybierz_date(data);

            Wybieranie_sprawy(nr_filii, nr_sprawy);

            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(5);
            string termin = "12:00";
            Wybieranie_godziny(termin);

            

                
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(5);
            Formularz(7);
            driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(15);

            //Start albo powrót
            try
            {
                driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(3);
                string powrót = "/html/body/main/div/div[3]/a[2]/span";
                driver.FindElement(By.XPath(powrót)).Click();
            }
            catch
            {
                driver.Manage().Timeouts().ImplicitWait = System.TimeSpan.FromSeconds(3);
                string powrót = "/html/body/main/div/div/a/span";
                driver.FindElement(By.XPath(powrót)).Click();

            }

            /*
            string entryurl = "http://isqueue-dev.pl/#/success";
            Assert.AreEqual(entryurl, driver.Url, "URL jest nieprawid³owy");
            */
        }

        [TearDown]
        public void QuitDriver()
        {
            driver.Quit();
        }

    }

    class User
    {
        public string Imie { get; set; }
        public string Nazwisko { get; set; }
        public string Email { get; set; }
        public string Telefon { get; set; }
    }

}