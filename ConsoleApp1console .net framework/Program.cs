using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using SeleniumExtras.WaitHelpers;
//using static System.Net.Mime.MediaTypeNames;

class Program
{
    static void Main(string[] args)
    {
        // Initialize Excel application
        Application excelApp = null;
        Workbook workbook = null;
        Worksheet worksheet = null;

        IWebDriver driver = null;

        try
        {
            // Start Excel application (Excel will be invisible)
            excelApp = new Application();
            excelApp.Visible = false;

            // Load the Excel file
            string excelFilePath = @"C:\Users\Shujath WorkSpace\Selenium\SmokeTestConsole\SmokeTestConsole\Usertest1.xlsx";  // Your Excel file path
            workbook = excelApp.Workbooks.Open(excelFilePath);
            worksheet = (Worksheet)workbook.Sheets[1];  // Access the first sheet (adjust if your data is on a different sheet)

            // Get the number of rows in the worksheet (excluding header)
            int rowCount = worksheet.UsedRange.Rows.Count;

            // Initialize WebDriver (Chrome in this case)
            driver = new ChromeDriver();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(160)); // Wait up to 10 seconds for elements


            // Loop through each row of data (starting from row 2, as row 1 contains headers)
            for (int row = 2; row <= 2; row++)
            {
                try
                {
                    // Extract data from the Excel sheet
                    string url = worksheet.Cells[row, 1].Text;  // Column 1: URL
                    string userTag = worksheet.Cells[row, 2].Text;  // Column 2: Username field selector
                    string passwordTag = worksheet.Cells[row, 3].Text;  // Column 3: Password field selector
                    string username = worksheet.Cells[row, 4].Text;  // Column 4: Username value
                    string password = worksheet.Cells[row, 5].Text;  // Column 5: Password value
                    string submitTag = worksheet.Cells[row, 6].Text;  // Column 6: Submit button selector
                    string Targeturl = worksheet.Cells[row, 7].Text;  //column 7: Targeturl/Targetelement after login

                    // Navigate to the URL
                    driver.Navigate().GoToUrl(url);

                    // Wait for the username field to be visible (using WebDriverWait)
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector(userTag)));

                    // Perform login using the extracted data
                    driver.FindElement(By.CssSelector(userTag)).SendKeys(username);
                    driver.FindElement(By.CssSelector(passwordTag)).SendKeys(password);
                    driver.FindElement(By.CssSelector(submitTag)).Click();

                    // Wait for the next page or element to confirm login (e.g., dashboard)
                    try
                    {
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector(Targeturl)));
                        //d.FindElement(By.CssSelector(".logoutButton")).Displayed); // Wait until login page changes or the logout button is visible

                        var expectedId = worksheet.Cells[row, 7].Value.ToString();

                        // If login is successful, update status in Excel
                        if (Targeturl == expectedId)
                        {
                            worksheet.Cells[row, 8].Value = "Passed";  // Column 7: Status
                        }
                        else
                        {
                            worksheet.Cells[row, 8].Value = "Failed";  // Column 7: Status
                        }

                        Console.WriteLine($"Login attempt for {username} completed. Status: {worksheet.Cells[row, 7].Value}");
                    }
                    catch (WebDriverTimeoutException)
                    {
                        // If we can't find the element or the URL doesn't change, it's a failed login
                        worksheet.Cells[row, 8].Value = "Failed";
                        Console.WriteLine($"Login failed for {username}. Status: Failed");
                    }
                }
                catch (Exception ex)
                {
                    // If an error occurs, update status in Excel
                    worksheet.Cells[row, 8].Value = "Error";  // Column 7: Status
                    Console.WriteLine($"Error for row {row}: {ex.Message}");
                }
            }

            new WebDriverWait(driver, TimeSpan.FromSeconds(120)).Until(ExpectedConditions.ElementToBeClickable(By.Id("cList_9"))).Click();  // index 10 AlZarooni Group
            Console.WriteLine("Successfully clicked on AlZarooni Group.");

            new WebDriverWait(driver, TimeSpan.FromSeconds(120)).Until(ExpectedConditions.ElementToBeClickable(By.Id("btnLogn"))).Click();   // clicking submit button
            Console.WriteLine("Successfully clicked on submit button.");

            new WebDriverWait(driver, TimeSpan.FromSeconds(120)).Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".quickview__close"))).Click(); // clicking  on cancel button
            Console.WriteLine("Successfully clicked on  cancel button.");

            new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementToBeClickable(By.ClassName("btn__primary_login"))).Click();// selecting ALL for location
            Console.WriteLine("Successfully ALL for location.");

            new WebDriverWait(driver, TimeSpan.FromSeconds(120)).Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"revwebbody\"]/app-root/app-home/form/div/app-sidemenu/div/ul/li[2]/i"))).Click(); // Clicking Finance
            Console.WriteLine("Successfully clicked on Finance");

            new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"sidebar_menu\"]/app-submenu[2]/div/div[2]/ul/li[1]/ul/li/a"))).Click(); // Clicking on chart of account 
            Console.WriteLine("Successfully clicked on Chart Of account");

            new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"maincontainer\"]/div/div[1]/div[2]/span[2]"))).Click(); // Clicking on ADD 
            Console.WriteLine("Successfully clicked on Add");

            //new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//ul[@class='groupItems'])[2]"))).Click();
            //Console.WriteLine("Successfully clicked on Chart Of Accounts");

            //new WebDriverWait(driver, TimeSpan.FromSeconds(120)).Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[text()='Accounts']/following-sibling::ul[@class='groupItems']\r\n"))).Click();
            //Console.WriteLine("Successfully clicked on Add");


            //IWebElement txtRegister = driver.FindElement(By.XPath(("//span[@class='secondary-btn' and @title='Add']")));
            //txtRegister.Click();

            // Save changes to the Excel file



            // Pause the browser for 10 seconds
            Thread.Sleep(10000);



            {
                for (int row = 7; row <= 8; row++)
                    try
                    {
                        // Extract data from the Excel sheet

                        string AccountCode = worksheet.Cells[row, 1].Text;  // Column 1
                        string AccountName = worksheet.Cells[row, 2].Text;
                        string SaveButtonTag = worksheet.Cells[row, 6].Text;


                        // Perform login using the extracted data
                        driver.FindElement(By.Id("AccountCode")).SendKeys(AccountCode);
                        driver.FindElement(By.Id("AccountName")).SendKeys(AccountName);
                        driver.FindElement(By.XPath("//*[@id=\"revwebbody\"]/modal-container/div/div/app-model-dialog/div/div[2]/form/app-quick-account/div/div[2]/div/button[1]/i")).Click();

                        //new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementToBeClickable(By.ClassName("primary-btn-wicon"))).Click();  // Save Account
                        //Console.WriteLine("Successfully Added Account.");

                        try
                        {


                            var savebutton = worksheet.Cells[row, 7].Value.ToString();

                            // If login is successful, update status in Excel
                            if (SaveButtonTag == savebutton)
                            {
                                worksheet.Cells[row, 8].Value = "Passed";  // Column 7: Status
                            }
                            else
                            {
                                worksheet.Cells[row, 8].Value = "Failed";  // Column 7: Status
                            }

                            Console.WriteLine($"Login attempt for {AccountCode} completed. Status: {worksheet.Cells[row, 7].Value}");
                        }

                        catch (Exception ex)
                        {
                            // If an error occurs, update status in Excel
                            worksheet.Cells[row, 8].Value = "Error";  // Column 7: Status
                            Console.WriteLine($"Error for row {row}: {ex.Message}");
                        }


                        Thread.Sleep(10000);

                        workbook.Save();

                    }
                    catch (Exception ex)
                    {
                        // If an error occurs, update status in Excel
                        worksheet.Cells[row, 8].Value = "Error";  // Column 7: Status
                        Console.WriteLine($"Error for row {row}: {ex.Message}");
                    }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            // Close Excel and release resources properly
            if (worksheet != null)
            {
                Marshal.ReleaseComObject(worksheet);
            }
            if (workbook != null)
            {
                workbook.Close(false);  // Don't save changes, Excel will save after each update
                Marshal.ReleaseComObject(workbook);
            }
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            //// Close WebDriver (if initialized)
            if (driver != null)
            {
                driver.Quit();
            }
        }

    }
}