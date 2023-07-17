using AventStack.ExtentReports;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.Create_Widgets
{
    public class CreateWidgetFromAddWidget
    {

        public IWebDriver? _driver;
        public FindElements? _findElements;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;
        public CommonFunctions? _commonFunc;

        public CreateWidgetFromAddWidget(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc)
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _commonFunc = commonFunc;
            PageFactory.InitElements(_driver, this);
        }

        [FindsBy(How = How.XPath, Using = "//div[@id='AddWidgetModal']/div/div[1]")]
        public IWebElement? AddWidgetPopUp;

        [FindsBy(How = How.XPath, Using = "//li[@id='standardgraphs']/span")]
        public IWebElement? StandardGraphs;

        [FindsBy(How = How.Id, Using = "btnAddWidget")]
        public IWebElement? AddWidgetBtn;

        [FindsBy(How = How.Id, Using = "icongraph")]
        public IWebElement? AddWidgetIcon;

        public void AddWidgets(WidgetCategories wCat, WidgetTypes wType)
        {
            try
            {
                _findElements.ClickElement(AddWidgetIcon, _currentPage, $"Analysis Page - Add Widget Icon");

                _findElements.VerifyElement(AddWidgetPopUp, _currentPage, $"Add Widget Popup");

                _findElements.ClickElement(StandardGraphs, _currentPage, $"Add Widget -Standard graphs");

                //if (wCat != WidgetCategories.XfStandard)
                //{
                //    _driver.FindElement(By.CssSelector("[data-catgtype='" + GetAddWidgetCatgName(wCat) + "']")).Click();
                //}

                var selectWidget = _driver.FindElement(By.CssSelector("#AddWidgetModal li[data-widgetcategory='" + (int)wCat + "'][data-widgettype='" + (int)wType + "']"));

                _findElements.ClickElementByJavaScript(selectWidget, _currentPage, $"Add widget -{wType.ToString()} ");

                _findElements.ClickElement(AddWidgetBtn, _currentPage, $"Add Widget - Add Widget Button");

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in add widget functionality. The error is {e.Message}");
            }
        }
    }
}
