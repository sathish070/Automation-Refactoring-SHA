using System;
using System.Xml;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;

namespace SHAProject.EditPage
{
    public class GraphProperties
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;

        public GraphProperties(string currentPage, IWebDriver driver, FindElements findElements, CommonFunctions commonFunc)
        {
            _driver = driver;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            PageFactory.InitElements(_driver, this);
        }

        #region Graph Properties Elements

        // Dropdowm properties
        [FindsBy(How = How.CssSelector, Using = ".measurement-property")]
        public IWebElement? MeasurementField;

        [FindsBy(How = How.Id, Using = "ddl_measurement")]
        public IWebElement? MeasurementDropdown;

        [FindsBy(How = How.CssSelector, Using = ".rate-property")]
        public IWebElement? RateField;

        [FindsBy(How = How.Id, Using = "ddl_view")]
        public IWebElement? RateDropdown;

        [FindsBy(How = How.CssSelector, Using = ".errorformat-property")]
        public IWebElement? ErrorFormatField;

        [FindsBy(How = How.Id, Using = "ddl_err")]
        public IWebElement? ErrorFormatDropdown;

        [FindsBy(How = How.CssSelector, Using = ".baselinedrp-property")]
        public IWebElement? BaselineField;

        [FindsBy(How = How.Id, Using = "ddl_baseline")]
        public IWebElement? BaselineDropdown;

        [FindsBy(How = How.CssSelector, Using = ".sortby-property")]
        public IWebElement? SortByField;

        [FindsBy(How = How.Id, Using = "ddl_sort")]
        public IWebElement? SortByDropdown;


        // Toggle button properties
        [FindsBy(How = How.CssSelector, Using = ".display-property")]
        public IWebElement? DisplayField;

        [FindsBy(How = How.Id, Using = "dispaly")]
        public IWebElement? DisplayToggle;

        [FindsBy(How = How.XPath, Using = "//input[@id='rddisplay'][1]")]
        public IWebElement? DisplayGroup;

        [FindsBy(How = How.XPath, Using = "//input[@id='rddisplay'][2]")]
        public IWebElement? DisplayWells;

        [FindsBy(How = How.CssSelector, Using = ".y-property")]
        public IWebElement? YField;

        [FindsBy(How = How.Id, Using = "levelrate")]
        public IWebElement? YToggle;

        [FindsBy(How = How.XPath, Using = "//input[@id=\"rddy1\"][1]")]
        public IWebElement? YRate;

        [FindsBy(How = How.XPath, Using = "//input[@id=\"rddy1\"][2]")]
        public IWebElement? YLevel;


        // Toggle properties
        [FindsBy(How = How.CssSelector, Using = ".normalization-property")]
        public IWebElement? NormalizationField;

        [FindsBy(How = How.Id, Using = "chknormalize")]
        public IWebElement? NormalizationToggle;

        [FindsBy(How = How.CssSelector, Using = ".bgcorrection-property")]
        public IWebElement? BackgroundCorrectionField;

        [FindsBy(How = How.Id, Using = "chkbackground")]
        public IWebElement? BackgroundCorrectionToggle;

        #endregion

        #region Dropdowm properties

        public void Measurement(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(MeasurementField, "Measurement", _currentPage, "Graph Property - Measurement");

            VerifySelectDropdown(MeasurementField, MeasurementDropdown, graphProperties.Measurement, "Measurement");
        }

        public void Rate(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(RateField, "Rate", _currentPage, "Graph Property - Rate");

            VerifySelectDropdown(RateField, RateDropdown, graphProperties.Rate, "Rate");
        }

        public void ErrorFormat(WidgetItems graphProperties, WidgetCategories wCat, WidgetTypes wType)
        {
            _findElements.ElementTextVerify(ErrorFormatField, "ErrorFormat", _currentPage, $"Graph Property - Error Format");

            List<string> ErrorFormatOptions = null;
            //int widgetPosition = _commonFunc.GetWidgetPosition(wCat, wType);

            //if ((int)wType == widgetPosition)
                ErrorFormatOptions = new List<string> { "OFF", "Std Dev", "SEM" };

            VerifyDropdownOptions(ErrorFormatField, ErrorFormatOptions, "Error Format");

            VerifySelectDropdown(ErrorFormatField, ErrorFormatDropdown, graphProperties.ErrorFormat, "Error Format");
        }

        public void Baseline(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(BaselineField, "Baseline", _currentPage, $"Graph Property - Baseline");

            VerifySelectDropdown(BaselineField, BaselineDropdown, graphProperties.Baseline, "Baseline");
        }

        public void SortBy(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(SortByField, "ErrorFormat", _currentPage, $"Graph Property - Sort By");

            VerifySelectDropdown(SortByField, SortByDropdown, graphProperties.SortBy, "Sort By");
        }

        public void VerifyDropdownOptions(IWebElement dropdownElement, List<string> drpOptions, string propertyName)
        {
            IList<IWebElement> dropdownOptions = dropdownElement.FindElements(By.TagName("option"));

            List<string> optionTexts = new List<string>();
            foreach (IWebElement option in dropdownOptions)
            {
                optionTexts.Add(option.Text.Trim());
            }

            bool areEqual = drpOptions.SequenceEqual(optionTexts);
            if (areEqual)
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The expected {propertyName} are verified with {optionTexts.Count} optoins.");
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The expected {propertyName} are unmatched with {optionTexts.Count} optoins.");
        }

        public void VerifySelectDropdown(IWebElement fieldElement, IWebElement dropdownElement, string expectedText, string propertyName)
        {
            try
            {
                IWebElement selectedOption = dropdownElement.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                string defaultText = selectedOption.Text;

                if (expectedText == defaultText)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedText} and actual {propertyName} are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, fieldElement);
                }
                else if (expectedText != defaultText)
                {
                    bool status = _findElements.SelectByText(dropdownElement, expectedText);
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedText} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, fieldElement);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} - {expectedText} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, fieldElement);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {dropdownElement} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Toggle button properties

        public void Display(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(DisplayField, "Display", _currentPage, "Graph Property - Display");

            VerifySelectToggleBtn(DisplayField, DisplayToggle, graphProperties.Display, "Display");
        }

        public void Y(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(YField, "Y", _currentPage, "Graph Property - Y");

            VerifySelectToggleBtn(YField, YToggle, graphProperties.Y, "Y");
        }

        public void VerifySelectToggleBtn(IWebElement fieldElement, IWebElement toggleElement, string expectedText, string propertyName)
        {
            IWebElement displayOption = null;
            try
            {
                IWebElement defaultOption = toggleElement.FindElement(By.CssSelector(".btn-on.active"));

                string defaultText = defaultOption.Text.Trim();

                if (expectedText == defaultText)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedText} and actual {propertyName} are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, fieldElement);
                }
                else if (expectedText != defaultText)
                {
                    displayOption = expectedText == "Group" ? DisplayGroup : DisplayWells;
                    bool status = _findElements.ClickElementByJavaScript(displayOption, _currentPage, $"{propertyName} - {expectedText}");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedText} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, fieldElement);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} - {expectedText} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error, fieldElement);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {displayOption} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Toggle properties
        public void Normalization(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(NormalizationField, "Normalization", _currentPage, "Graph Property - Normalization");

            VerifySelectToggle(NormalizationField, NormalizationToggle, graphProperties.Normalization, "Normalization");
        }

        public void BackgroundCorrection(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(BackgroundCorrectionField, "BackgroundCorrection", _currentPage, "GraphProperty - BackgroundCorrection");

            VerifySelectToggle(BackgroundCorrectionField, BackgroundCorrectionToggle, graphProperties.BackgroundCorrection, "BackgroundCorrection");
        }

        public void VerifySelectToggle(IWebElement fieldElement, IWebElement toggleElement, bool expectedText, string propertyName)
        {
            try
            {
                bool isChecked = toggleElement.Selected;
                if (isChecked && expectedText)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} is {(isChecked ? "ON" : "OFF")} and actual {propertyName} options are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, fieldElement);
                }
                else
                {
                    bool status = _findElements.ClickElementByJavaScript(toggleElement, _currentPage, $"{propertyName} Button");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} is {(expectedText ? "ON" : "OFF")}.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, fieldElement);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} option was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, fieldElement);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {toggleElement} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{propertyName} Normalization is not verified. The error is {e.Message}");
            }
        }

        #endregion

        public void Graphproperties()
        {
            try
            {
                IReadOnlyCollection<IWebElement> graphProperties = _driver.FindElements(By.CssSelector(".graph-ms"));

                foreach (IWebElement graphProperty in graphProperties)
                {
                    if (graphProperty.Displayed)
                        _findElements.ElementTextVerify(graphProperty, graphProperty.Text, _currentPage, $"Graph property - {graphProperty.Text}");
                }

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Graph properties elemnets text has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Graph properties elemnets text has not been verified.");
            }
        }

        public void VerifyExpectedGraphUnits(string ExpectedGraphUnits, WidgetTypes widgetType, bool expectedBtnStatus)
        {
            try
            {
                Thread.Sleep(2000);

                ChartType chartType = _commonFunc.GetChartType(widgetType);
                string graphUnits = _commonFunc.GetGraphUnits(chartType);

                if (graphUnits == ExpectedGraphUnits)
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The graph units -{graphUnits} and exact graph units -{ExpectedGraphUnits} are equal.");
                else
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The graph units -{graphUnits} and exact graph units -{ExpectedGraphUnits} are not equal.");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Verify Normalization units has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Verify Normalization units has not been verified. The error is {e.Message}");
            }
        }
    }
}
