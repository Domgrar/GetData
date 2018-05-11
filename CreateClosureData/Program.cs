using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateClosureData
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> ticketList = new List<string>();
            IWebElement dropDownMenu;
            Ticket thisTicket = new Ticket();
            var options = new InternetExplorerOptions();
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;

            // Search for ticket in EV
            
            IWebDriver driver = new InternetExplorerDriver(options);
            driver.Navigate().GoToUrl("https://dhgllp.easyvista.com/");

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            dropDownMenu = driver.FindElement(By.XPath("//input[@id='GlobalCurrentQuerycombo-ui']"));

            dropDownMenu.SendKeys(Keys.Control + "a");
            dropDownMenu.SendKeys("Incidents");
            dropDownMenu.SendKeys(Keys.ArrowDown);
            dropDownMenu.SendKeys(Keys.Enter);

            
            


           

    


            ticketList = GetIncidentList();
            
            foreach(string item in ticketList)
            {

                thisTicket = getAllTicketData(item, driver);

                //Block to display all values for testing
                Console.WriteLine("Requesting Person : " + thisTicket.Requestor);
                Console.WriteLine("Category : " + thisTicket.Category);
                Console.WriteLine("Solved by : " + thisTicket.SolvedBy);
                Console.WriteLine("Description : " + thisTicket.Description);
            }

            Console.ReadLine();
        }

        /// <summary>
        /// Method to get the tickets information and store it in a string to pass to writeallvalues
        /// </summary>
        /// <param name="incidentNumber"></param>
        /// <returns></returns>
        public static Ticket getAllTicketData(string incidentNumber, IWebDriver driver)
        {
            
            IWebElement incidentSearch;
            IWebElement requestorElement;
            IWebElement categoryElement;
            IWebElement tableElement;
            IWebElement descriptionElement;
            Ticket thisTicket = new Ticket();
            int i = 0;

            

            //Thread.Sleep(4000);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            incidentSearch = driver.FindElement(By.XPath("//input[@name='GlobalSearchText']"));
            incidentSearch.SendKeys(incidentNumber);
            incidentSearch.SendKeys(Keys.Enter);

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            //Get the elements we need and their text
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.PollingInterval = TimeSpan.FromSeconds(10);
            wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
            requestorElement = wait.Until(theDriver => theDriver.FindElement(By.XPath("//input[@name='AM_RECIPIENT.LAST_NAME']")));

        
            
            //driver.FindElement(By.XPath("//input[@name='AM_RECIPIENT.LAST_NAME']"));        // Gets the manager
                    thisTicket.Requestor = requestorElement.GetAttribute("value").ToString();


           
                    categoryElement = driver.FindElement(By.XPath("//input[contains(@id, 'SD_CATALOG.TITLE_EN')] [contains(@class, 'form_input_ro')]"));
                    thisTicket.Category = categoryElement.GetAttribute("value").ToString();
           

          
                    tableElement = driver.FindElement(By.XPath("//table[@id='tbl_dialog_body_section_1_0']"));
                    thisTicket.SolvedBy = tableElement.FindElement(By.XPath("//td[contains(., ',')]")).Text;
         

         
                    descriptionElement = driver.FindElement(By.XPath("//div[@id='SD_REQUEST_COMMENT1']"));
                    thisTicket.Description = descriptionElement.Text;
           
      

            

            return thisTicket;
        }
        public static bool WriteAllValues(Ticket thisTicket)
        {


            return true;
        }

        public static List<string> GetIncidentList()
        {
            List<string> iList = new List<string>();
            string text = System.IO.File.ReadAllText(@"C:\Users\jf6856\source\repos\CreateClosureData\CreateClosureData\TicketCloses.txt"); // Upgrade to relative path

            iList.AddRange(text.Split(' '));
 
            return iList;
        }
        
    }
}
