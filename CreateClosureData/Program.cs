using CsvHelper;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CreateClosureData
{
    class Program
    {
        static void Main(string[] args)
        {
            //Variables we need
            List<string> ticketList = new List<string>();
            IWebElement dropDownMenu;
            Ticket thisTicket = new Ticket();

            //Creating the webdriver
            var options = new InternetExplorerOptions();
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            IWebDriver driver = new InternetExplorerDriver(options);

            //Creating the webdriver wait
            Type[] ignores = new Type[] { typeof(OpenQA.Selenium.StaleElementReferenceException) };
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.PollingInterval = TimeSpan.FromSeconds(30);
            wait.IgnoreExceptionTypes(ignores);

            //Goto ev
            driver.Navigate().GoToUrl("https://dhgllp.easyvista.com/");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            //Find the search textbox
            dropDownMenu = driver.FindElement(By.XPath("//input[@id='GlobalCurrentQuerycombo-ui']"));
            dropDownMenu.SendKeys(Keys.Control + "a");
            dropDownMenu.SendKeys("Incidents");
            dropDownMenu.SendKeys(Keys.ArrowDown);
            dropDownMenu.SendKeys(Keys.Enter);

            //Search each ticket and get it's info from EV
            ticketList = GetIncidentList();
            foreach(string item in ticketList)
            {
                thisTicket = getAllTicketData(item, driver, wait);

                Console.WriteLine("Requesting Person : " + thisTicket.Requestor);
                Console.WriteLine("Category : " + thisTicket.Category);
                Console.WriteLine("Solved by : " + thisTicket.SolvedBy);
                Console.WriteLine("Description : " + thisTicket.Description);
                Console.WriteLine("**************************************************************");

                WriteAllValues(thisTicket);
            }

            Console.ReadLine();
        }

        /// <summary>
        /// Method to get the tickets information and store it in a string to pass to writeallvalues
        /// </summary>
        /// <param name="incidentNumber"></param>
        /// <returns></returns>
        public static Ticket getAllTicketData(string incidentNumber, IWebDriver driver, WebDriverWait wait)
        {

            IWebElement incidentSearch;
            IWebElement tableElement;
            Ticket thisTicket = new Ticket();
            int i = 0;

            

            //Thread.Sleep(4000);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            incidentSearch = driver.FindElement(By.XPath("//input[@name='GlobalSearchText']"));
            incidentSearch.SendKeys(incidentNumber);
            incidentSearch.SendKeys(Keys.Enter);

            Thread.Sleep(2000);


            //Getting the elments we need off the web page
            thisTicket.Requestor = wait.Until(theDriver => theDriver.FindElement(By.XPath("//input[@name='AM_RECIPIENT.LAST_NAME']"))).GetAttribute("value").ToString();

            thisTicket.Category = wait.Until(theDriver => theDriver.FindElement(By.XPath("//input[contains(@id, 'SD_CATALOG.TITLE_EN')] [contains(@class, 'form_input_ro')]"))).GetAttribute("value").ToString();

            tableElement = wait.Until(theDriver => theDriver.FindElement(By.XPath("//table[@id='tbl_dialog_body_section_1_0']")));
           
            thisTicket.Description = wait.Until(theDriver => theDriver.FindElement(By.XPath("//div[@id='SD_REQUEST_COMMENT1']"))).Text;
            


            IList<IWebElement> tableRow;
            try
            {
                tableRow = tableElement.FindElements(By.TagName("tr"));
            }
            catch (StaleElementReferenceException)
            {
                tableRow = wait.Until(theDriver => tableElement.FindElements(By.TagName("tr")));
            }


            IList<IWebElement> td;
           
            i = 0;
            int z = 0;
            Console.WriteLine("***TR Data***");
            foreach(IWebElement element in tableRow)
            {
                if (i == 1) {
                    Console.WriteLine(element.Text);

                    td = element.FindElements(By.TagName("td"));
                    foreach (IWebElement data in td)
                    {
                        if (z == 4)
                        {
                            Console.Write("DATA: ");
                            Console.WriteLine(data.Text + "**************************************************************");
                            thisTicket.SolvedBy = data.Text;
                            break;
                        }
                        z++;
                    }
                }
                //td = element.FindElement()
                i++;
            }
      
            return thisTicket;
        }

        /// <summary>
        /// Method to write values to csv file
        /// </summary>
        /// <param name="thisTicket"></param>
        /// <returns></returns>
        public static bool WriteAllValues(Ticket thisTicket)
        {
            try
            {
                using (var memoryStream = new MemoryStream())
                using (var streamWriter = new StreamWriter(@"C:\FTG\boi.csv", true))
                using (var cw = new CsvWriter(streamWriter))
                {

                    
                    cw.WriteField(thisTicket.Requestor);
                    cw.WriteField(thisTicket.Category);
                    cw.WriteField(thisTicket.SolvedBy);
                    cw.NextRecord();
                }
                
                
                return true;
            }catch(Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
         
        }

        /// <summary>
        /// Method to read all incident ticket from text file
        /// </summary>
        /// <returns></returns>
        public static List<string> GetIncidentList()
        {
            List<string> iList = new List<string>();
            string text = System.IO.File.ReadAllText(@"C:\Users\jf6856\OneDrive - DHG LLP\GitHub\GetData\GetData\CreateClosureData\TicketCloses.txt"); // Upgrade to relative path

            iList.AddRange(text.Split(' '));
 
            return iList;
        }
        
    }
}
