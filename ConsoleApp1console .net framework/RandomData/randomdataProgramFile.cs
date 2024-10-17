using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SmokeTestLogin.Utilities.Random_Data;
//using SmokeTestLogin.Utilities.Random_Data;

class randomdataProgramFile
{
    static void Main(string[] args)
    {
        // Path to your Excel file
        string excelFilePath = @"C:\Users\Shujath WorkSpace\Selenium\ConsoleApp1\Book2.xlsx";

        // Initialize the helper class
        var excelHelper = new ExcelSequentialHelper(excelFilePath);

        // Generate random data for different fields
        string sequentialValue = excelHelper.GenerateRandomData("sequential", 0, 0);
        string emailValue = excelHelper.GenerateRandomData("email", 1, 0);
        string nameValue = excelHelper.GenerateRandomData("name", 2, 0);
        string numberValue = excelHelper.GenerateRandomData("number", 3, 0);

        // Use the generated data in Selenium
        IWebDriver driver = new ChromeDriver();
        driver.Navigate().GoToUrl("https://example.com/form");

        // Fill different fields with generated data
        driver.FindElement(By.Id("nameFieldId")).SendKeys(nameValue);
        driver.FindElement(By.Id("emailFieldId")).SendKeys(emailValue);
        driver.FindElement(By.Id("numberFieldId")).SendKeys(numberValue);

        // Submit the form (if required)
        driver.FindElement(By.Id("submitButtonId")).Click();

        // Close the browser
        driver.Quit();
    }
}
