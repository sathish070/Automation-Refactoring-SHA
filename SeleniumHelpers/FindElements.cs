using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using SHAProject.Utilities;
using AventStack.ExtentReports;
using SeleniumExtras.WaitHelpers;

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
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
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
                fieldName =$" {widgetName} {fieldName}";
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Enabled && webElement.Displayed)
                {
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    webElement.Click();
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Pass, $"{fieldName} - is displayed and clickable on the page");
                    return true;
                }
                else
                {
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"{fieldName} - is not displayed and  clickable on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                    return false;
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Element - {element} not found on page. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                return false;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail,  $"Error occured while verifying the element - {fieldName}. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                return false;
            }
        }

        public IWebElement VerifyElement(IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                fieldName =$" {widgetName} {fieldName}";
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Displayed)
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
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode",Status.Fail, $"Element -{element} is not found on the page.");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode",Status.Fail, $"Error occured while verifying the element - {fieldName}. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            return element;
        }

        public void SendKeys(string sendKeys, IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                fieldName =$" {widgetName} {fieldName}";
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Displayed)
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
                    ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Element - {fieldName} is not displayed on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Element - {element} not found on the page . The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Status.Fail, $"Error occured while verifying the element - {fieldName}. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
        }

        public bool ClickElementByJavaScript(IWebElement? element, string currentPage, string fieldName)
        {
            try
            {
                fieldName =$" {widgetName} {fieldName}";
                //_driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Enabled && webElement.Displayed)
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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} not found on the page. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                return false;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error occured while verifying the element - {fieldName}. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                return false;
            }
        }

        public void ScrollIntoViewAndClickElementByJavaScript(IWebElement? element, string currentPage, string fieldName)
        {
            try
            {
                fieldName =$" {widgetName} {fieldName}";
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Enabled && webElement.Displayed)
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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} not found on page. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error occured while verifying the element - {fieldName}. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
        }

        public void ScrollIntoView(IWebElement element)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Displayed)
                {
                    IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                    jScript?.ExecuteScript("arguments[0].scrollIntoView(true);", webElement);
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Element is not display and unable to scroll down to view");
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element -{element} not found on page. The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unable to scroll down to the element. The error is {e.Message}");
            }
        }

        public void ElementTextVerify(IWebElement element, string givenName, string currentPage , string fieldName)
        {
            try
            {
                fieldName =$" {widgetName} {fieldName}";
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Displayed && webElement.Text.Equals(givenName))
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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element -{element} not found on page . The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element -{fieldName} text verification is failed.The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
        }

        public void ActionsClass(IWebElement element)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Displayed)
                {
                    Actions actions = new Actions(_driver);
                    actions.MoveToElement(element).Build().Perform();
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Element is not display and unable to move to the element");
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element -{element} not found on page. The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unable to move to the element.The error is {e.Message}");
            }
        }

        public void ActionsClassClick(IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Enabled && webElement.Displayed)
                {
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    Actions actions = new Actions(_driver);
                    actions.MoveToElement(element).Click().Perform();
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{fieldName} is displayed and clickable on the page");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{fieldName} - is not displayed and so unable to perform move action and click the element on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} not found on page. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unable to move and click the element.The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
        }

        public void ActionsClassDoubleClick(IWebElement element, string currentPage, string fieldName)
        {
            try
            {
                IWebElement webElement = WaitForElementVisible(element);
                if (webElement != null && webElement.Enabled && webElement.Displayed)
                {
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Info, element);
                    Actions actions = new Actions(_driver);
                    //actions.MoveToElement(element).DoubleClick().Perform();
                    actions.DoubleClick(element).Perform();
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{fieldName} is displayed and double clickable on the page");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{fieldName} - is not displayed, so unable to perform move action and double click the element on the page");
                    ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {element} not found on page. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unable to move and double click the element.The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, fieldName, ScreenshotType.Error, element);
            }
        }
        public void SelectFromDropdown(IWebElement element, string currentPage, string selectionMethod, string value, string propertyName)
        {
            try
            {
                SelectElement dropdown = new SelectElement(element);

                switch (selectionMethod.ToLower())
                {
                    case "index":
                        int index = int.Parse(value);
                        dropdown.SelectByIndex(index);
                        break;

                    case "text":
                        dropdown.SelectByText(value);
                        break;

                    case "value":
                        dropdown.SelectByValue(value);
                        break;

                    default:
                        throw new ArgumentException("Invalid selection method.");
                }

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{propertyName}option - {value} was selected from the dropdown.");
                ScreenShot.ScreenshotNow(_driver, currentPage, $"Dropdown option - {value}", ScreenshotType.Info, element);
            }
            catch (ArgumentException ex)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Invalid selection method. The error is {ex.Message}");
                ScreenShot.ScreenshotNow(_driver, currentPage, $"Invalid selection method - {selectionMethod}", ScreenshotType.Error, element);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{propertyName}option - {value} was not selected from the dropdown.");
                ScreenShot.ScreenshotNow(_driver, currentPage, $"Dropdown option - {value}", ScreenshotType.Error, element);
            }
        }
    }
}
