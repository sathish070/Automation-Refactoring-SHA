using AventStack.ExtentReports;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SHAProject.Utilities;
using SeleniumExtras.WaitHelpers;
using SeleniumExtras.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection.Metadata.Ecma335;
using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using System.Threading;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Diagnostics;
using AngleSharp.Dom;
using OpenQA.Selenium.DevTools.V112.Fetch;
using OpenQA.Selenium.Interactions;

namespace SHAProject.SeleniumHelpers
{
    public class FindElements : Tests
    {
        public IWebDriver? _driver;

        public FindElements(IWebDriver driver)
        {
            _driver = driver;
        }

        public IWebElement WaitForElementVisible(IWebElement element)
        {
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(60));
            return wait.Until(driver =>
            {
                try
                {
                    if (element.Displayed)
                        return element;
                    else
                        return null;
                }
                catch (NoSuchElementException e)
                {
                    return null;
                }
            });
        }

        public bool ClickElement(IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement.Enabled && webElement.Displayed)
                {
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    element.Click();
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Pass, $"{fieldName} - is displayed and clickable on the page");
                    return true;
                }
                else
                {
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"{fieldName} - is displayed  not clickable on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                    return false;
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Element - {element} -not found on page . The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                return false;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail,  $"Unknown error { e.Message} occurred on page ");
                return false;
            }
        }

        public IWebElement VerifyElement(IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement.Displayed)
                {
                    ExtentReport.ExtentTest(currentPage=="Login"?"ExtentTest" : "ExtentTestNode", Status.Pass, $"{fieldName} is displayed");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                }
                else
                {
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"{fieldName} is not displayed");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode",Status.Fail, $"{element} is not found on page.");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode",Status.Fail, $"Unknown error {e.Message} is occur on page.");
            }
            return element;
        }

        public void SendKeys(string sendKeys, IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement.Displayed)
                {
                    element.Clear();
                    element.SendKeys(sendKeys);
                    string elementText = element.GetAttribute("value");
                    if (sendKeys.Equals(elementText)) 
                    {
                        ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Pass, $"The given text - {sendKeys} and element text are equal");
                        ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    }
                    else
                    {
                        ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"The given text - {sendKeys} and element text are not equal");
                        ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                    }
                }
                else
                {
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Element - {fieldName} - is not displayed on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Element - {element} -not found on page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Unknown error {e.Message} occurred on page ");
            }
        }

        public bool ClickElementByJavaScript(IWebElement? element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement.Enabled && webElement.Displayed)
                {
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                    jScript.ExecuteScript("arguments[0].click();", element);
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{fieldName} - is displayed and clickable on the page");
                    return true;
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{fieldName} - is not displayed and clickable on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                    return false;
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} -not found on page . The error is {e.Message}");
                return false;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unknown error {e.Message} occurred on page ");
                return false;
            }
        }

        public void ScrollIntoViewAndClickElementByJavaScript(IWebElement? element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement.Enabled && webElement.Displayed)
                {
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                    jScript?.ExecuteScript("arguments[0].scrollIntoView(true);", element);  
                    jScript?.ExecuteScript("arguments[0].click();", element);
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{fieldName} - is displayed and clickable on the page");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{fieldName} - is not displayed and clickable on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} -not found on page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unknown error {e.Message} occurred on page.");
            }
        }

        public void ScrollIntoView(IWebElement element)
        {
            try
            {
                IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                jScript?.ExecuteScript("arguments[0].scrollIntoView(true);", element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" The error is {e.Message}");
            }
        }

        public void ElementTextVerify(IWebElement element, string givenName, string currentPage , string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement.Displayed && webElement.Text.Contains(givenName))
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{fieldName} field is displayed and text contains -{element.Text.Trim()} ");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{fieldName} field is not displayed and does not contains the text");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} -not found on page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element text verification is failed.The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
        }

        public bool ActionsClass(IWebElement element)
        {
            Actions actions = new Actions(_driver);
            actions.MoveToElement(element).Build().Perform();
            return true;
        }

        public bool ActionsClassClick(IWebElement element)
        {
            Actions actions = new Actions(_driver);
            actions.MoveToElement(element).Click().Perform();
            return true;
        }

        public bool ActionsClassDoubleClick(IWebElement element)
        {
            Actions actions = new Actions(_driver);
            actions.MoveToElement(element).DoubleClick().Perform();
            return true;
        }

        public bool SelectByIndex(IWebElement element, int value )
        {
            SelectElement dropdown = new SelectElement(element);
            dropdown.SelectByIndex(value);
            return true;
        }

        public bool SelectByText(IWebElement element, string text)
        {
            SelectElement dropdown = new SelectElement(element);
            dropdown.SelectByText(text.ToString());
            return true;
        }

        public bool SelectByValue(IWebElement element, string text)
        {
            SelectElement dropdown = new SelectElement(element);
            dropdown.SelectByValue(text);
            return true;
        }
    }
}
