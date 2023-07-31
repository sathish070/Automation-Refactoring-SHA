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
using AngleSharp.Dom;
using System.Diagnostics;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using System.Security.Policy;

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

        [FindsBy(How = How.XPath, Using = "//div[@class=\"graph-ocr-msrmt col-lg-12\"]")]
        public IWebElement? GraphPropertyField;

        [FindsBy(How = How.XPath, Using ="//div[@id='grapharea']/div[1]")]
        public IWebElement? GraphAreaField;

        // Dropdown properties
        [FindsBy(How = How.CssSelector, Using = ".graph-ms.select-measurement.hideprop")]
        public IWebElement? MeasurementField;

        [FindsBy(How = How.Id, Using = "ddl_measurement")]
        public IWebElement? MeasurementDropdown;

        [FindsBy(How = How.CssSelector, Using = ".graph-ms.select-measurement.rate.hiderate")]
        public IWebElement? RateField;

        [FindsBy(How = How.Id, Using = "ddl_view")]
        public IWebElement? RateDropdown;

        [FindsBy(How = How.CssSelector, Using = ".graph-ms.error-form.errorformat")]
        public IWebElement? ErrorFormatField;

        [FindsBy(How = How.Id, Using = "ddl_err")]
        public IWebElement? ErrorFormatDropdown;

        [FindsBy(How = How.CssSelector, Using = "#baselineselection")]
        public IWebElement? BaselineField;

        [FindsBy(How = How.Id, Using = "ddl_baseline")]
        public IWebElement? BaselineDropdown;

        [FindsBy(How = How.CssSelector, Using = ".sortby-property")]
        public IWebElement? SortByField;

        [FindsBy(How = How.Id, Using = "ddl_sort")]
        public IWebElement? SortByDropdown;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"graph-ms error-form hideoligo oligo-property\"]")]
        public IWebElement? OligoField;

        [FindsBy(How = How.Id, Using = "ddl_oligo")]
        public IWebElement? OligoDropdown;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"graph-ms error-form hideoligo-induced induced-property\"]")]
        public IWebElement? InducedField;

        [FindsBy(How = How.Id, Using = "ddl_induced")]
        public IWebElement? InducedDropdown;

        // Toggle button properties
        [FindsBy(How = How.CssSelector, Using = ".graph-ms.select-display.hideprop")]
        public IWebElement? DisplayField;

        [FindsBy(How = How.Id, Using = "dispaly")]
        public IWebElement? DisplayToggle;

        [FindsBy(How = How.XPath, Using = "//input[@name=\"rddisplay\"][@value=\"0\"]")]
        public IWebElement? DisplayGroup;

        [FindsBy(How = How.XPath, Using = "//input[@name=\"rddisplay\"][@value=\"1\"]")]
        public IWebElement? DisplayWells;

        [FindsBy(How = How.CssSelector, Using = ".graph-ms.select-y1.hideprop")]
        public IWebElement? YField;

        [FindsBy(How = How.Id, Using = "levelrate")]
        public IWebElement? YToggle;

        [FindsBy(How = How.XPath, Using = "//input[@id=\"rddy1\"][@value=\"0\"]")]
        public IWebElement? YRate;

        [FindsBy(How = How.XPath, Using = "//input[@id=\"rddy1\"][@value=\"1\"]")]
        public IWebElement? YLevel;

        // Toggle properties
        [FindsBy(How = How.CssSelector, Using = ".graph-ms.select-normal.normalization-toggle")]
        public IWebElement? NormalizationField;

        [FindsBy(How = How.Id, Using = "chknormalize")]
        public IWebElement? NormalizationToggle;

        [FindsBy(How = How.CssSelector, Using =".graph-ms.bg-correction.hideprop")]
        public IWebElement? BackgroundCorrectionField;

        [FindsBy(How = How.Id, Using = "chkbackground")]
        public IWebElement? BackgroundCorrectionToggle;

        #endregion

        #region Dropdown properties

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
            _findElements.ElementTextVerify(ErrorFormatField, "Error Format", _currentPage, $"Graph Property - Error Format");

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
            _findElements.ElementTextVerify(SortByField, "Sort By", _currentPage, $"Graph Property - Sort By");

            VerifySelectDropdown(SortByField, SortByDropdown, graphProperties.SortBy, "Sort By");
        }

        public void Oligo(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(OligoField, "Oligo", _currentPage, $"Graph Property - Oligo");

            VerifySelectDropdown(OligoField, OligoDropdown, graphProperties.Oligo, "Oligo");
        }

        public void Induced(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(InducedField, "Induced", _currentPage, $"Graph Property - Induced");

            VerifySelectDropdown(InducedField, InducedDropdown, graphProperties.Induced, "Induced");
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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The expected {propertyName} are verified with {optionTexts.Count} options.");
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The expected {propertyName} are unmatched with {optionTexts.Count} options.");
        }

        public void VerifySelectDropdown(IWebElement fieldElement, IWebElement dropdownElement, string expectedText, string propertyName)
        {
            try
            {
                IWebElement selectedOption = dropdownElement.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                string defaultText = selectedOption.Text;

                if (dropdownElement.Text.Contains(expectedText))
                {
                    if (expectedText == defaultText)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedText} and actual {propertyName} are equal.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Info, fieldElement);
                    }
                    else if (expectedText != defaultText)
                    {
                        _findElements.SelectFromDropdown(dropdownElement, _currentPage, "text", expectedText, propertyName);
                    }
                }
                else 
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{propertyName} dropdown does not contains the expected text - {expectedText}");
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
                    ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Info, fieldElement);
                }
                else if (expectedText != defaultText)
                {
                    if(propertyName == "Display")
                        displayOption = expectedText == "Group" ? DisplayGroup : DisplayWells;
                    if(propertyName == "Y")
                        displayOption = expectedText == "Rate" ? YRate : YLevel;
                    try
                    {
                        IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                        jScript.ExecuteScript("arguments[0].click();", displayOption);

                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedText} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Info, fieldElement);
                    }
                    catch (Exception e)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} - {expectedText} was not selected. The error is {e.Message}");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Error, fieldElement);
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

            string Opacity = NormalizationToggle.GetCssValue("opacity");
            if (Opacity == "1")
            {
                VerifySelectToggle(NormalizationField, NormalizationToggle, graphProperties.Normalization, "Normalization");
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Normalization is in disable mode and file is not normalized");
            }
        }

        public void BackgroundCorrection(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(BackgroundCorrectionField, "Background Correction", _currentPage, "GraphProperty - BackgroundCorrection");

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
                    ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Info, fieldElement);
                }
                else
                {
                    try
                    {
                        if (propertyName == "Normalization")
                            _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()");

                        if(propertyName == "BackgroundCorrection")
                            _driver.ExecuteJavaScript<string>("return document.getElementById(\"chkbackground\").click()");

                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} is {(expectedText ? "ON" : "OFF")}.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Info, fieldElement);
                    }
                    catch (Exception)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} option was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"{propertyName}", ScreenshotType.Error, fieldElement);
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

        public void GraphProperty()
        {
            _findElements.VerifyElement(GraphPropertyField, _currentPage, $"Edit Widget Page -Graph Property");
        }

        public void GraphArea()
        {
            _findElements.VerifyElement(GraphAreaField, _currentPage, $"Edit Widget Page -Graph Area");
        }

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

        public void VerifyExpectedGraphUnits(string ExpectedGraphUnits, WidgetTypes widgetType)
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

        public void VerifyDropdown()
        {
             try
            {
                Dictionary<string, string> inducedDropdownOptions = new Dictionary<string, string>
                {
                    { "1", "N/A" },
                    { "2", "1st" },
                    { "3", "1st, 2nd" }
                };

                SelectElement selectOligo = new SelectElement(OligoDropdown);
                foreach (IWebElement option in selectOligo.Options)
                {
                    string optionValue = option.GetAttribute("value");

                    if (string.IsNullOrEmpty(optionValue))
                        continue;

                    _findElements.SelectFromDropdown(OligoDropdown, _currentPage, "value", optionValue, $"Add View Popup - Oligo dropdown");

                    SelectElement selectInduced = new SelectElement(InducedDropdown);
                    string selectedOptionText = selectInduced.SelectedOption.Text;

                    foreach (IWebElement Inducedoptions in selectInduced.Options)
                    {
                        string inducedOptionValue = Inducedoptions.Text;

                        if (string.IsNullOrEmpty(inducedOptionValue))
                            continue;

                        string expectedText = inducedDropdownOptions[optionValue];

                        if (expectedText.Contains(inducedOptionValue))
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Verification Passed for option {optionValue}: Expected '{expectedText}', Actual '{selectedOptionText}'");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Verification Failed for option {optionValue}: Expected '{expectedText}', Actual '{selectedOptionText}'");
                        }
                    }

                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Induced option - {selectedOptionText} was selected from the dropdown.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, $"Dropdown option - {selectedOptionText}", ScreenshotType.Info, InducedDropdown);

                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Error ocuured while verify the oligo and induced ");
            }
        }
    }
}
