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

            Thread.Sleep(2000);

            Type[] ignores = new Type[] { typeof(OpenQA.Selenium.StaleElementReferenceException) };
            //Get the elements we need and their text
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.PollingInterval = TimeSpan.FromSeconds(30);
            wait.IgnoreExceptionTypes(ignores);


            //driver.Navigate().Refresh();

            //driver.FindElement(By.XPath("//input[@name='AM_RECIPIENT.LAST_NAME']"));        // Gets the manager
           
            thisTicket.Requestor = wait.Until(theDriver => theDriver.FindElement(By.XPath("//input[@name='AM_RECIPIENT.LAST_NAME']"))).GetAttribute("value").ToString();
            //thisTicket.Requestor = requestorElement.GetAttribute("value").ToString();

            try // This next https://stackoverflow.com/questions/37837407/c-sharp-selenium-webdriver-raise-invalidoperationexception-ocurred-in-webdriver
            {
                thisTicket.Category = wait.Until(theDriver => theDriver.FindElement(By.XPath("//input[contains(@id, 'SD_CATALOG.TITLE_EN')] [contains(@class, 'form_input_ro')]"))).GetAttribute("value").ToString();
            }catch(Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
                //categoryElement = driver.FindElement(By.XPath("//input[contains(@id, 'SD_CATALOG.TITLE_EN')] [contains(@class, 'form_input_ro')]"));
            //thisTicket.Category = categoryElement.GetAttribute("value").ToString();


            tableElement = wait.Until(theDriver => theDriver.FindElement(By.XPath("//table[@id='tbl_dialog_body_section_1_0']")));
            //tableElement = driver.FindElement(By.XPath("//table[@id='tbl_dialog_body_section_1_0']"));
            //thisTicket.SolvedBy = tableElement.FindElement(By.XPath("//td[contains(., ',')]")).Text;


            thisTicket.Description = wait.Until(theDriver => theDriver.FindElement(By.XPath("//div[@id='SD_REQUEST_COMMENT1']"))).Text;
            //descriptionElement = driver.FindElement(By.XPath("//div[@id='SD_REQUEST_COMMENT1']"));
            //thisTicket.Description = descriptionElement.Text;


            //Disgusting code below

            //IList<IWebElement> tableRow = tableElement.FindElements(By.TagName("tr"));
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
            string solvedBy;
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

                }
                
                
                return true;
            }catch(Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
         
        }

        public static List<string> GetIncidentList()
        {
            List<string> iList = new List<string>();
            string text = System.IO.File.ReadAllText(@"C:\Users\jf6856\OneDrive - DHG LLP\GitHub\GetData\GetData\CreateClosureData\TicketCloses.txt"); // Upgrade to relative path

            iList.AddRange(text.Split(' '));
 
            return iList;
        }
        
    }
}
