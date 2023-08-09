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
using SeleniumExtras.PageObjects;

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
            PageFactory.InitElements(_driver, this);
        }

        [FindsBy(How = How.XPath, Using = "//div[@class=\"col-md-12 groups-blocks\"]")]
        public IWebElement? GroupLegendsField;

        public void GroupLegendsArea()
        {
            _findElements.ScrollIntoView(GroupLegendsField);

            _findElements.VerifyElement(GroupLegendsField, _currentPage, "Edit Widget Page - Group Legends");
        }

        public void EditWidgetGroupLegends(WidgetCategories wCat, WidgetTypes wType, WidgetItems widget)
        {
            Thread.Sleep(2000);
            try
            {
                if(wType != WidgetTypes.DoseResponse)
                {
                    //IReadOnlyCollection<IWebElement> groupLegends = _driver.FindElements(By.CssSelector(".stress-li")).Take(4).ToList();

                    IReadOnlyCollection<IWebElement> groupLegends = _driver.FindElements(By.CssSelector(".stress-li"));

                    foreach (IWebElement groupLegend in groupLegends)
                    {
                        if (groupLegend.Text.Contains("Background"))
                            continue;

                        _findElements.ActionsClass(groupLegend);

                        IWebElement groupLegendText = groupLegend.FindElement(By.CssSelector("span:nth-child(2)"));

                        // verify the group values

                        //string errorType = widget.ErrorFormat  == "Std Dev" ? "SD" : widget.ErrorFormat == "SEM" ? "SM" : "None";

                        string errorType = widget.ErrorFormat  == "Std Dev" ? "Std Dev:" : widget.ErrorFormat == "SEM" ? "SEM:" : "None:";

                        IWebElement groupError = groupLegend.FindElement(By.CssSelector(".groupmean"));
                        if (groupError.Enabled && groupError.Displayed)
                        {
                            if (groupLegend.Text.Contains(errorType))
                            {
                                _findElements.VerifyElement(groupError, _currentPage, "Mean & Error Value");
                            }
                            else
                            {
                                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Group details Mean & Error Value are not displayed.");
                                ScreenShot.ScreenshotNow(_driver, _currentPage, "Error Screenshot", ScreenshotType.Error, groupError);
                            }
                        }

                        // verify the single-click event
                        _findElements.ActionsClassClick(groupLegend, _currentPage, $"Hightlight the group legend {groupLegendText.Text}");

                        //ScreenShot.ScreenshotNow(_driver, _currentPage, $"HightLight the group legend  {groupLegend.Text}", ScreenshotType.Info, groupLegend);

                        // verify the double-click event
                        _findElements.ActionsClassDoubleClick(groupLegend, _currentPage, $"Unselect the group legend -{groupLegendText.Text}");

                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"After unselect the group legend - {groupLegendText.Text}", ScreenshotType.Info, groupLegend);
                    }
                }
                else
                {
                    IReadOnlyCollection<IWebElement> CompoundList = _driver.FindElements(By.CssSelector(".col-md-12.stress-li.selected-li"));

                    foreach (IWebElement Compound in CompoundList)
                    {
                        if (Compound.Displayed)
                        {
                            IWebElement CompoundText = Compound.FindElement(By.CssSelector("span:nth-child(2)"));

                            _findElements.ActionsClass(Compound);

                            ScreenShot.ScreenshotNow(_driver, _currentPage, $"Compound List - {CompoundText.Text}", ScreenshotType.Info, Compound);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in edit widget page exports functionality. The error is {e.Message}");
            }
        }

        public void EditWidgetDataTableGroupLegends(WidgetCategories wCat, WidgetTypes wType, WidgetItems widget)
        {

            Thread.Sleep(2000);
            try
            {
                IReadOnlyCollection<IWebElement> groupLegends = _driver.FindElements(By.CssSelector(".stress-li")).Take(4).ToList();

                foreach (IWebElement groupLegend in groupLegends)
                {
                    if (groupLegend.Text.Contains("Background"))
                        continue;

                    _findElements.ClickElement(groupLegend, _currentPage, $"Selected Group Legends  and Group Legends are highlighted &  The Groups  are showing in the DataTable Widget{groupLegend.Text}");
                }

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page exports functionality. The error is {e.Message}");
            }
        }
    }
}
