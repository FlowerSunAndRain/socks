using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using ex = Microsoft.Office.Interop.Excel;

namespace Socks
{
    internal class Socks
    {
        static void Main()
        {
            IWebDriver driver = new EdgeDriver();
            driver.Navigate().GoToUrl(@"https://market.yandex.ru/");
            try
            {
                driver.FindElement(By.Id("js-button")).Click();
            }
            catch (NoSuchElementException)
            {
  
            }

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("header-search")));
            driver.FindElement(By.Id("header-search")).SendKeys("Носки с дедом морозом");
            driver.FindElement(By.ClassName("_1z5kk")).Click();

            ex.Application excelApp = new ex.Application();
            ex.Workbook workBook = excelApp.Workbooks.Add();
            ex.Worksheet workSheet = workBook.ActiveSheet;

            workSheet.Cells[1, "A"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[1]/article/div[1]/div[2]/div/a/div/span/span[1]")).Text;
            workSheet.Cells[1, "B"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[1]/article/div[2]/h3/a/span[2]")).Text;
            workSheet.Cells[1, "C"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[1]/article/div[1]/div[2]/div/a")).GetAttribute("href");

            workSheet.Cells[2, "A"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[2]/article/div[1]/div[2]/div/a/div/span/span[1]")).Text;
            workSheet.Cells[2, "B"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[2]/article/div[2]/h3/a/span[2]")).Text;
            workSheet.Cells[2, "C"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[2]/article/div[1]/div[2]/div/a")).GetAttribute("href");

            workSheet.Cells[3, "A"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[3]/article/div[1]/div[2]/div/a/div/span/span[1]")).Text;
            workSheet.Cells[3, "B"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[3]/article/div[2]/h3/a/span[2]")).Text;
            workSheet.Cells[3, "C"] = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[5]/div/div/div/div/div/div/div[5]/div/div/div/div/main/div/div/div/div/div/div/div[1]/div/div/div[3]/article/div[1]/div[2]/div/a")).GetAttribute("href");

            workBook.Close(true, "D:\\Prices.xlsx");
            excelApp.Quit();
            
        }
    }
}
