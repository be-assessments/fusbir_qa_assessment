using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace PilotProject
{
    public class PilotTest
    {
        private IWebDriver webDriver;
        private ExtentReports extent;
        private ExtentTest test;
        public static string dir = AppDomain.CurrentDomain.BaseDirectory;
        public static string testResultPath = Path.Combine(dir.Replace("bin\\Debug\\net8.0", "TestResults"));
        public static string driverpath = Path.Combine(dir.Replace("bin\\Debug\\net8.0", "Drivers"));

        [OneTimeSetUp]
        public void Setup()
        {
            try
            {
                // Initialize ExtentReports
               
                if (!Directory.Exists(testResultPath))
                {
                    Directory.CreateDirectory(testResultPath);
                }

                var htmlReporter = new ExtentSparkReporter(Path.Combine(testResultPath, "TestReport.html"));
                extent = new ExtentReports();
                extent.AttachReporter(htmlReporter);

                if (extent == null)
                {
                    throw new InvalidOperationException("TestReport could not be initialized.");
                }

                // Initialize ChromeDriver
                test = extent.CreateTest("Setup");
                test.Log(Status.Info, "Initializing ChromeDriver");
                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument("--headless=new"); // Enable headless mode
                
                var chromeDriverPath = ""+driverpath+"chromedriver.exe";
                if (!File.Exists(chromeDriverPath))
                {
                    throw new FileNotFoundException("ChromeDriver executable not found at specified path.", chromeDriverPath);
                }

                webDriver = new ChromeDriver(chromeDriverPath, chromeOptions);
                if (webDriver == null)
                {
                    throw new InvalidOperationException("WebDriver could not be initialized.");
                }

                webDriver.Manage().Window.Maximize();
                test.Log(Status.Info, "Navigating to Google");
                webDriver.Navigate().GoToUrl("https://www.google.com/");
            }
            catch (Exception ex)
            {
                // Log the error to the extent report and re-throw it
                if (test != null)
                {
                    test.Log(Status.Fail, $"Setup failed: {ex.Message}");
                }
                throw; // Re-throw the exception to ensure the test fails
            }
        }

        [SetUp]
        public void BeforeEachTest()
        {
            test = extent.CreateTest("Open OTO FulFillment Website and Assert the expected text");
          
        }

        [Test]
        public void OTO_Fulfillment()
        {
            try
            {
             
                IWebElement GoogleSearchBar = webDriver.FindElement(By.Name("q"));


                // Initialize step number
                int stepNo = 1;
                ExtentTest GoogleSearchBarStep = test.CreateNode(""+stepNo+ ": Enter the value in Google Search Bar");
                var GoogleSearchBarTestData = "OTO Fulfilment";
                GoogleSearchBar.SendKeys(GoogleSearchBarTestData);
                GoogleSearchBarStep.Log(Status.Pass, "Value Entered " + GoogleSearchBarTestData + " sucessfully in Google Search Bar Field");
                
                stepNo =stepNo+1;
                ExtentTest GoogleSearchStep = test.CreateNode("" + stepNo + ": Click on Google Search");
                GoogleSearchBar.SendKeys(Keys.Enter);
                GoogleSearchStep.Pass("User click on Google Search button sucessfully");
                GoogleSearchStep.Log(Status.Info, "Waiting for search results to load");
                Thread.Sleep(2000);

                stepNo = stepNo + 1;
                ExtentTest LocatingWebsiteStep = test.CreateNode(""+stepNo+": Locating website link for 'oto fulfilment'");
                IWebElement WebsiteName = webDriver.FindElement(By.XPath("//cite[contains(text(),'https://www.otofulfilment.com')]"));
                WebsiteName.Click();
                LocatingWebsiteStep.Log(Status.Pass, "user is clicked on the website link sucessfully");
                LocatingWebsiteStep.Log(Status.Info, "Waiting for the new page to load");
                Thread.Sleep(2000);

                stepNo = stepNo + 1;
                ExtentTest LocatingElementTextStep = test.CreateNode(""+ stepNo + ": Locating element with text and validate through Assert");
                IWebElement element = webDriver.FindElement(By.XPath("//span[contains(text(),'Grow your business')]"));
                element.Click();
                string expectedText = "Grow your business with faster, more efficient and hassle-free fulfillment services by OTO Fulfilment!";
                string actualText = element.Text;

                LocatingElementTextStep.Log(Status.Info, $"Verifying the text of the located element. Expected: {expectedText}, Actual: {actualText}");
                Assert.That(actualText.ToLower(), Is.EqualTo(expectedText.ToLower()), $"Expected text '{expectedText}' but found '{actualText}'.");
                LocatingElementTextStep.Log(Status.Pass, "Test Passed Successfully");


                stepNo=stepNo + 1;
                ExtentTest FAQStep = test.CreateNode("" + stepNo + ": Ensuring That the FAQ section works as intended");
                IWebElement FAQ = webDriver!.FindElement(By.XPath("//button[@type='button']//span[contains(text(),'Are inventory levels updated timely?')]"));
                FAQ.Click();
                IWebElement accordionBody = webDriver.FindElement(By.Id("accordion-collapse-body-3"));

                // Check if the element is dwebisplayed
                if (accordionBody.Displayed)
                {
                    // Log the success
                    FAQStep.Log(Status.Pass,"After Clicking on FAQ only one Answer has been shown at a time");

                    AttachScreenshot(FAQStep);
                }
                else
                {
                    // Log the fail
                    FAQStep.Log(Status.Fail, "After clicking on multiple FAQ then multiple answer have been shown ");

                    AttachScreenshot(FAQStep);
                }

            }
            catch (NoSuchElementException ex)
            {
                test.Log(Status.Fail, $"Element not found: {ex.Message}");
            }
            catch (AssertionException ex)
            {
                test.Log(Status.Fail, $"Assertion failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                test.Log(Status.Fail, $"Test Failed: {ex.Message}");
            }
        }
      

        // Helper method to normalize text by removing extra whitespace and newlines
        private string NormalizeText(string text)
        {
            // Replace multiple whitespace characters with a single space
            // and remove leading/trailing whitespace
            return Regex.Replace(text.Trim(), @"\s+", " ");
        }
        private void AttachScreenshot(ExtentTest test)
        {
            try
            {
                var screenshot = ((ITakesScreenshot)webDriver).GetScreenshot();

                // Convert screenshot to Base64String
                var base64Screenshot = screenshot.AsBase64EncodedString;

                // Build media entity from Base64String
                var mediaEntity = MediaEntityBuilder
                    .CreateScreenCaptureFromBase64String(base64Screenshot)
                    .Build();

                // Attach screenshot to the report
                test.Info("Screenshot", mediaEntity);
            }
            catch (Exception ex)
            {
                // Log error if attachment fails
                Console.WriteLine($"Failed to attach screenshot: {ex.Message}");
            }
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            if (webDriver != null)
            {
              
                webDriver.Quit();
                webDriver.Dispose();
            }

            if (extent != null)
            {
               
                extent.Flush();
            }
        }
    }
}
