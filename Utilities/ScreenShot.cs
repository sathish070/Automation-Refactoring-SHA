using AventStack.ExtentReports;
using OpenQA.Selenium;
using SHAProject.EditPage;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHAProject.Utilities
{
    public static class ScreenShot
    {
        public static IJavaScriptExecutor? jScript;
        public static IJavaScriptExecutor JavaScriptExecutor(IWebDriver driver)
        {
            IJavaScriptExecutor jScript = (IJavaScriptExecutor)driver;
            return jScript;
        }

        public static void ScreenshotNow(IWebDriver driver, string currentPage, string ImageName, ScreenshotType status = ScreenshotType.Info, IWebElement? element = null, ArrayList elementList = null)
        {
            try
            {
                jScript = JavaScriptExecutor(driver);
                var existingBorder = string.Empty;
                if (element != null)
                {
                    /*Highlight the element with a red border*/
                    existingBorder = (string)jScript.ExecuteScript("return arguments[0].getAttribute('style', arguments[1]);", element, "border");
                    if (ImageName == "DataTable Header widgetName")
                    {
                        jScript.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", element, " border: 3px solid red; position: absolute;");
                    }
                    if (ImageName.Contains("Graph Y-Axis"))
                    {
                        jScript.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", element, " border: 3px solid red; stroke: red; stroke-width:3px;");
                    }
                    else
                    {
                        jScript.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", element, " border: 3px solid red;");
                    }
                    Thread.Sleep(1000);
                }
                if (elementList != null)
                {
                    /* Highlight all elements in the list with a red border*/
                    for (int i = 0; i < elementList.Count; i++)
                    {
                        existingBorder = (string)jScript.ExecuteScript("return arguments[0].getAttribute('style', arguments[1]);", elementList[i], "border");
                        jScript.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", elementList[i], " border: 3px solid red;");
                        Thread.Sleep(1000);
                    }
                }

                ImageName = (ImageName.Replace("/", "-").Replace("|", "-").Replace(":", "").Replace("\r\n", "-").Replace(">", ""));
                string path = Tests.loginFolderPath;
                string Imagefolder = status == ScreenshotType.Info ? "Success" : "Error";
                string ImagePath = path + "\\" + currentPage + "\\" + Imagefolder + "\\" + ImageName + ".png";
                if (Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix)
                {
                    ImagePath = path + "/" + Imagefolder + "/" + ImageName + ".png";
                }

                string screenshotPath = ImagePath;
                string reportScreenshotPath = @"" + path + "\\" + currentPage + "\\" + Imagefolder + "\\" + ImageName + ".png";
                ITakesScreenshot screenshotDriver = driver as ITakesScreenshot;
                Screenshot screenshot = screenshotDriver.GetScreenshot();
                screenshot.SaveAsFile(screenshotPath, ScreenshotImageFormat.Png);
                ExtentReport.ExtentScreenshot(currentPage == "Login" ? "ExtentTest" : "ExtentTestNode", Imagefolder == "Success" ? Status.Pass : Status.Fail, ImageName, reportScreenshotPath);

                if (element != null)
                {
                    Thread.Sleep(1000);
                    /* Set the element back to its original state*/
                    jScript.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", element, existingBorder); //set back to original state
                }
                if (elementList != null)
                {
                    Thread.Sleep(1000);
                    for (int i = 0; i < elementList.Count; i++)
                    {
                        /* Set all elements in the list back to their original state*/
                        jScript.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);", elementList[i], existingBorder); //set back to original state
                    }
                }
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "An error occured in taking screenshot. The error is " + ex.Message);
            }
        }
    }
}
