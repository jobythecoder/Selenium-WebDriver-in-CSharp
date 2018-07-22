
using System;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;

using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;


namespace TechTest
{
    [TestFixture]
    public class TechTestClass
    {

        IWebDriver driver;

        public ExtentReports extent;
        public ExtentTest test;


        [OneTimeSetUp]
        public void setupOnce()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\extent\properties\TechTest_Project_Properties.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter("C:\\extent\\extent.html");
            extent = new ExtentReports();
            //extent.AttachReporter(htmlReporter);

            //String projectName = "Tech Test";
            //String testEnvironment = "Selenium WebDrivere";
            //String browserVersions = "Chrome67";

            //htmlReporter.Configuration().DocumentTitle = "Tech Test - Test Report";
            //extent.AddSystemInfo("Project Name", projectName);
            //extent.AddSystemInfo("Test Environment", testEnvironment);
            //extent.AddSystemInfo("Browser Versions", browserVersions);


            extent.AttachReporter(htmlReporter);
                      
            //ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter("C:\\extent\\extent.html");
            //ExtentReports extent = new ExtentReports();
            //htmlReporter.Configuration().DocumentTitle = "Joby...Document Title";
            //extent.Flush();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            string[] arr = new string[5];
            for (int i = 1; i <= rowCount; i++)
            {

                    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    extent.AddSystemInfo(xlRange.Cells[i, 1].Value2.ToString(), xlRange.Cells[i, 2].Value2.ToString());

            }

            xlApp.Quit();
            extent.Flush();
        }

        [SetUp]
        public void TC_01_RecordAScenario()
        {
            //test = extent.CreateTest(TestContext.CurrentContext.Test.Name);

            driver = new ChromeDriver("C:\\Users\\Joby Joseph\\Documents\\eclipse-workspace\\browser_drivers");
            driver.Url = "https://www.mortgagehouse.com.au/resources/calculators/mortgage-repayment-calculator/?loan_amount=300000&loan_period=30&interest_rate=6&loan_type=principle_and_interest&introductory_rate=no&introductory_interest_rate=5.5&introductory_interest_rate_period=2";
            driver.Manage().Window.Maximize();
            Console.WriteLine("reshma recordascenario A");
            IWebElement loanAmount = driver.FindElement(By.XPath("//*[@id='Repayment_LoanAmount']"));
            IWebElement loanPeriod = driver.FindElement(By.XPath("//*[@id='Repayment_LoanPeriod']"));
            IWebElement interestRate = driver.FindElement(By.XPath("//*[@id='Repayment_InterestRate']"));
            IWebElement loanType = driver.FindElement(By.XPath("//*[@id='Repayment_LoanType']"));
            IWebElement repaymentWeekly = driver.FindElement(By.XPath("//*[@id='repayment-calculator-form']/div[2]/div[2]/div[2]/div[2]/ul/li[1]"));

            loanAmount.Clear();
            loanPeriod.Clear();
            interestRate.Clear();


            loanAmount.SendKeys("500000");
            loanPeriod.SendKeys("30");
            interestRate.SendKeys("6");

            var selectType = new SelectElement(loanType);

            String repaymentWeeklyValue = repaymentWeekly.Text;

            selectType.SelectByText("Interest Only");

            repaymentWeekly.Click();

        }

        [Test]
        public void TC_02_AssertTestsForScenario()
        {

            var test = extent.CreateTest("TC_02 AssertTestsForScenario");


            try
            {
                //var test = extent.CreateTest("Tech Test", "TC_01_RecordAScenario");
                //test.Fail("tc_01 failed");
                //extent.Flush();

                IWebElement MonthlyRepayment = driver.FindElement(By.XPath("//*[@id='repayment-calculator-form']/div[2]/div[2]/div[2]/div[2]/ul/li[3]"));
                IWebElement MonthlyRepaymentValue = driver.FindElement(By.XPath("//*[@id='repayment']"));
                IWebElement TotalInterestPayable = driver.FindElement(By.XPath("//*[@id='total-interest']"));
                IWebElement PageHeading = driver.FindElement(By.XPath("//h1"));

                MonthlyRepayment.Click();

                String monthlyRepaymentValue = MonthlyRepaymentValue.Text;
                String totalInterestPayableValue = TotalInterestPayable.Text;
                String pageHeading = PageHeading.Text;
                TestContext.WriteLine("TC02:Page heading....."+pageHeading);


                Assert.Multiple(() =>
                {
                   Assert.AreEqual("$576.92", monthlyRepaymentValue, "monthlyRepaymentValue match not found");
                   Assert.AreEqual("$870,000", totalInterestPayableValue, "TotalInterestPayable Value Incorrect");
                   Assert.AreEqual("Mortgage Repayment Calculator", pageHeading, "Page Heading Incorrect");

                });
                test.Pass("TC_02");
            }
            catch (Exception e)
            {
                test.Fail("TC_02");
            }

        }


        [Test]
        public void TC03_AssertAlternativeTestsForScenario()
        {
            var test = extent.CreateTest("TC_03 AssertTestsForScenario");

            try
            {

                IWebElement TotalInterestPayable = driver.FindElement(By.XPath("//*[@id='total-interest']"));
                IWebElement PageHeading = driver.FindElement(By.XPath("//h1"));

                Assert.Multiple(() =>
                {
                    Assert.AreEqual("$870,000", TotalInterestPayable.Text, "TotalInterestPayable is not matching");
                    Assert.AreEqual("Mortgage Repayment Calculator", PageHeading.Text, "PageHeading Mismatch");
                });
                test.Pass("TC_03");
            }
            catch (Exception e)
            {
                test.Fail("TC_03");
            }


        }

        [TearDown]
        public void afterMethod()
        {
            driver.Close();



        }
        [OneTimeTearDown]
        public void TearDown()
        {
            extent.Flush();

        }


    }
    
}
