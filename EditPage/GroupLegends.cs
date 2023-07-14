using System;
using System.Xml;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections;
using SkiaSharp;
using OfficeOpenXml;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;

namespace SHAProject.EditPage
{
    public class GroupLegends : Tests
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public GroupLegends(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc)
        {
            _driver = driver;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
        }

        public void EditWidgetGroupLegends(WidgetCategories wCat, WidgetTypes wType, WidgetItems widget)
        {

            Thread.Sleep(2000);
            try
            {
                IReadOnlyCollection<IWebElement> groupLegends = _driver.FindElements(By.CssSelector(".stress-li")).Take(4).ToList();

                foreach (IWebElement groupLegend in groupLegends)
                {
                    if (groupLegend.Text.Contains("Background"))
                        continue;

                    _findElements.ActionsClass(groupLegend);

                    // verify the group values
                    IWebElement groupError = groupLegend.FindElement(By.CssSelector(".groupmean"));
                    if (groupError.Enabled && groupError.Displayed)
                    {
                        if (groupLegend.Text.Contains(widget.ErrorFormat))
                        {
                            _findElements.VerifyElement(groupError, _currentPage, "Mean & Error Value");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Group details Mean & Error Value are not displayed.");
                            ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error, groupError);
                        }
                    }

                    // verify the single-click event
                    _findElements.ActionsClassClick(groupLegend);

                    // verify the double-click event
                    _findElements.ActionsClassDoubleClick(groupLegend);

                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error);
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page exports functionality. The error is {e.Message}");
            }
        }
    }
}
