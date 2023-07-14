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


        [FindsBy(How = How.CssSelector, Using = ".measurement-property")]
        public IWebElement? MeasurementField;

        [FindsBy(How = How.Id, Using = "ddl_measurement")]
        public IWebElement? MeasurementDropdown;

        [FindsBy(How = How.CssSelector, Using = ".rate-property")]
        public IWebElement? RateField;

        [FindsBy(How = How.Id, Using = "ddl_view")]
        public IWebElement? RateDropdown;

        [FindsBy(How = How.CssSelector, Using = ".display-property")]
        public IWebElement? DisplayField;

        [FindsBy(How = How.Id, Using = "dispaly")]
        public IWebElement? DisplayToggle;

        [FindsBy(How = How.XPath, Using ="//input[@id='rddisplay'][1]")]
        public IWebElement? DisplayGroup;

        [FindsBy(How = How.XPath, Using = "//input[@id='rddisplay'][2]")]
        public IWebElement? DisplayWells;

        [FindsBy(How = How.CssSelector, Using = ".y-property")]
        public IWebElement? YField;

        [FindsBy(How = How.Id, Using = "levelrate")]
        public IWebElement? YToggle;

        [FindsBy(How = How.XPath, Using ="//input[@id=\"rddy1\"][1]")]
        public IWebElement? YRate;

        [FindsBy(How = How.XPath, Using ="//input[@id=\"rddy1\"][2]")]
        public IWebElement? YLevel;

        [FindsBy(How = How.CssSelector, Using = ".normalization-property")]
        public IWebElement? NormalizationField;

        [FindsBy(How = How.Id, Using = "chknormalize")]
        public IWebElement? NormalizationToggle;

        [FindsBy(How = How.CssSelector, Using = ".errorformat-property")]
        public IWebElement? ErrorFormatField;

        [FindsBy(How = How.Id, Using = "ddl_err")]
        public IWebElement? ErrorFormatDropdown;

        [FindsBy(How = How.CssSelector, Using = ".bgcorrection-property")]
        public IWebElement? BackgroundCorrectionField;

        [FindsBy(How = How.Id, Using = "chkbackground")]
        public IWebElement? BackgroundCorrectionToggle;

        [FindsBy(How = How.CssSelector, Using = ".baselinedrp-property")]
        public IWebElement? BaselineField;

        [FindsBy(How = How.Id, Using = "ddl_baseline")]
        public IWebElement? BaselineDropdown;

        [FindsBy(How = How.CssSelector, Using = ".oligo-property")]
        public IWebElement? OligoField;

        [FindsBy(How = How.Id, Using = "ddl_oligo")]
        public IWebElement? OligoDropdown;

        public void Graphproprties()
        {
            try
            {
                IReadOnlyCollection<IWebElement> graphProperties = _driver.FindElements(By.CssSelector(".graph-ms"));

                foreach (IWebElement graphProperty in graphProperties)
                {
                    if (graphProperty.Displayed)
                    {
                        _findElements.ElementTextVerify(graphProperty, graphProperty.Text, _currentPage, $"Graph property - {graphProperty.Text}");
                    }
                }

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Graph properties elemnets text has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Graph properties elemnets text has not been verified.");
            }
        }

        #region Graph Properties - Measurement

        public void Measurement(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(MeasurementField, "Measurement", _currentPage, "Graph Property - Measurement");

            ExpectedMesurement(graphProperties);
        }

        public void ExpectedMesurement(WidgetItems graphProperties)
        {
            try
            {
                IWebElement selectedOption = MeasurementDropdown.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                string defaultValue = selectedOption.Text;

                if (defaultValue == graphProperties.Measurement)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Measurement - {graphProperties.Measurement} and actual values are equal.");
                }
                else
                {
                    bool selectedStatus = _findElements.SelectByText(MeasurementDropdown, graphProperties.Measurement);
                    if (selectedStatus)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Measurement - {graphProperties.Measurement} was selected from the dropdown.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, MeasurementField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Measurement - {graphProperties.Measurement} was not selected from the dropdown.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error, MeasurementField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {MeasurementDropdown} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Measurement is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph Properties - Rate

        public void Rate(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(RateField, "Rate", _currentPage, "Graph Property - Rate");

            ExpectedRate(graphProperties);
        }

        public void ExpectedRate(WidgetItems graphProperties)
        {
            try
            {
                IWebElement selectedOption = RateDropdown.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);

                string defaultValue = selectedOption.Text;

                if (graphProperties.Rate == defaultValue)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Rate - {graphProperties.Rate} and actual rate type are equal.");
                }
                else
                {
                    bool selectedStatus = _findElements.SelectByText(RateDropdown, graphProperties.Rate);
                    if (selectedStatus)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Rate - {graphProperties.Rate} was selected from the dropdown.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, RateField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Rate - {graphProperties.Rate} was not selected from the dropdown.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error, RateField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {RateDropdown} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Rate is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph Properties -Display

        public void Display(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(DisplayField, "Display", _currentPage, "Graph Property - Display");

            ExpectedDisplay(graphProperties);
        }

        public void ExpectedDisplay(WidgetItems graphProperties)
        {
            IWebElement displayOption = null;
            try
            {
                IWebElement defaultOption = DisplayToggle.FindElement(By.CssSelector(".btn-on.active"));

                string defaultOptionText = defaultOption.Text.Trim();

                if (graphProperties.Display == defaultOptionText)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Display - {graphProperties.Display} and actual display are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, DisplayField);
                }
                else if (defaultOptionText != graphProperties.Display)
                {
                    displayOption = graphProperties.Display == "Group" ? DisplayGroup : DisplayWells;
                    bool status = _findElements.ClickElementByJavaScript(displayOption, _currentPage, $"Display - {graphProperties.Display}");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Display - {graphProperties.Display} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, DisplayField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Display - {graphProperties.Display} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error, DisplayField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {displayOption} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Display is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph Properties -Y

        public void Y(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(YField, "Y", _currentPage, "Graph Property - Y");

            ExpectedY(graphProperties);
        }

        public void ExpectedY(WidgetItems graphProperties)
        {
            IWebElement rateOption = null;
            try
            {
                IWebElement defaultOption = YToggle.FindElement(By.CssSelector(".btn-on.active"));

                string defaultOptionText = defaultOption.Text.Trim();

                if (graphProperties.Y == defaultOptionText)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Y - {graphProperties.Y} and actual Y text are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, YField);
                }
                else if (defaultOptionText != graphProperties.Y)
                {
                    rateOption = graphProperties.Y == "Rate" ? YRate : YLevel;
                    bool status = _findElements.ClickElementByJavaScript(rateOption, _currentPage, $"Y toggle - {graphProperties.Y}");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Y - {graphProperties.Y} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, YField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Y - {graphProperties.Y} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Error, YField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {rateOption} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Y is not verified. The error is {e.Message}");
            }
         }

        #endregion

        #region Graph Properties -Normalization

        public void Normalization(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(NormalizationField, "Normalization", _currentPage, "Graph Property - Normalization");

            ExpectedNormalization(graphProperties);
        }

        public void ExpectedNormalization(WidgetItems graphProperties)
        {
            try
            {
                bool isChecked = NormalizationToggle.Selected;
                if (isChecked && graphProperties.Normalization)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Normalization is {(isChecked ? "ON" : "OFF")} and actual nortmalization options are equal."); 
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, NormalizationField);
                }
                else
                {
                    bool status =_findElements.ClickElementByJavaScript(NormalizationToggle, _currentPage, "Normalization Button");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Normalization is {(graphProperties.Normalization ? "ON" : "OFF")}.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, NormalizationField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Normalization option was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, NormalizationField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {NormalizationToggle} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Default Normalization is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph properties - Error Format

        public void ErrorFormat(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(ErrorFormatField, "ErrorFormat", _currentPage, $"Graph Property - Error Format");

            ExpectedErrorFormat(graphProperties);
        }

        public void ExpectedErrorFormat(WidgetItems graphProperties)
        {
            try
            {
                IWebElement selectedOption = ErrorFormatDropdown.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                string defaultValue = selectedOption.Text;

                if (defaultValue == graphProperties.ErrorFormat)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Error Format - {graphProperties.ErrorFormat} and actual error format are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, ErrorFormatField);
                }
                else if (defaultValue != graphProperties.ErrorFormat)
                {
                    bool status =_findElements.SelectByText(ErrorFormatDropdown, graphProperties.ErrorFormat);
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Error Format - {graphProperties.ErrorFormat} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, ErrorFormatField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Error Format - {graphProperties.ErrorFormat} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, ErrorFormatField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {ErrorFormatDropdown} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Error Format is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph Properties - Background Correction

        public void BackgroundCorrection(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(BackgroundCorrectionField, "BackgroundCorrection", _currentPage, "GraphProperty - BackgroundCorrection");

            ExpectedBackgroundCorrection(graphProperties);
        }

        public void ExpectedBackgroundCorrection(WidgetItems graphProperties)
        {
            try
            {
                bool isChecked = BackgroundCorrectionToggle.Selected;
                if (isChecked && graphProperties.BackgroundCorrection)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Background Correction is {(isChecked ? "ON" : "OFF")} and actual Background Correction options are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, BackgroundCorrectionField);
                }
                else
                {
                    bool status = _findElements.ClickElementByJavaScript(BackgroundCorrectionToggle, _currentPage, "Background Correction Button");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Background Correction is {(graphProperties.BackgroundCorrection ? "ON" : "OFF")}.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, BackgroundCorrectionField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Background Correction option was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, BackgroundCorrectionField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {BackgroundCorrectionToggle} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Default Background Correction is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph Properties - Baseline

        public void Baseline(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(BaselineField, "Baseline", _currentPage, $"Graph Property - Baseline");

            ExpectedBaseline(graphProperties);
        }

        public void ExpectedBaseline(WidgetItems graphProperties)
        {
            try
            {
                IWebElement selectedOption = BaselineDropdown.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                string defaultValue = selectedOption.Text;

                if (defaultValue == graphProperties.Baseline)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Baseline - {graphProperties.Baseline} and actual baseline are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, BaselineField);
                }
                else if (defaultValue != graphProperties.Baseline)
                {
                    bool status = _findElements.SelectByText(BaselineDropdown, graphProperties.Baseline);
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Baseline - {graphProperties.Baseline} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, BaselineField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Baseline - {graphProperties.Baseline} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, BaselineField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {BaselineDropdown} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Baseline is not verified. The error is {e.Message}");
            }
        }

        #endregion

        #region Graph Properties - Oligo
        public void Oligo(WidgetItems graphProperties)
        {
            _findElements.ElementTextVerify(OligoField, "Oligo", _currentPage, $"Graph Property - Oligo");

            ExpectedOligo(graphProperties);
        }

        public void ExpectedOligo(WidgetItems graphProperties)
        {
            try
            {
                IWebElement selectedOption = OligoDropdown.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                string defaultValue = selectedOption.Text;

                if (defaultValue == graphProperties.Oligo)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Baseline - {graphProperties.Baseline} and actual baseline are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, OligoField);
                }
                else if (defaultValue != graphProperties.Oligo)
                {
                    bool status = _findElements.SelectByText(OligoDropdown, graphProperties.Oligo);
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected Oligo - {graphProperties.Oligo} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, OligoField);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Baseline - {graphProperties.Oligo} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, OligoField);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {OligoDropdown} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected Baseline is not verified. The error is {e.Message}");
            }
        }

        #endregion

        public void VerifyNormalizationUnits(string ExactGraphUnits, WidgetTypes widgetType, bool expectedBtnStatus)
        {
            try
            {
                Thread.Sleep(2000);

                /* Check if the normalization toggle button is enabled and log its status*/
                bool IsNormalizationBtnEnabled = NormalizationToggle.Enabled;
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, IsNormalizationBtnEnabled ? "Normalization toggle button is enabled" : "Normalization toggle button is disabled");

                /* Check if the normalization toggle button is turned on and log its status*/
                bool IsNormalizationBtnToggledOn = NormalizationToggle.Selected;
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, IsNormalizationBtnToggledOn ? "Normalization toggle button is turned on" : "Normalization toggle button is turned off");

                /* If the button is not in the expected state, click it to toggle it and wait for 2 seconds*/
                if (IsNormalizationBtnEnabled && expectedBtnStatus && IsNormalizationBtnToggledOn == false)
                {
                    _findElements.ClickElementByJavaScript(NormalizationToggle, _currentPage,$"Normalization Button");
                    Thread.Sleep(2000);
                }
                else if (IsNormalizationBtnEnabled && expectedBtnStatus == false && IsNormalizationBtnToggledOn)
                {
                    _findElements.ClickElementByJavaScript(NormalizationToggle, _currentPage, $"Normalization Button");
                    Thread.Sleep(2000);
                }

                ChartType chartType = _commonFunc.GetChartType(widgetType);
                string graphUnits = _commonFunc.GetGraphUnits(chartType);
                if(graphUnits == ExactGraphUnits)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The graph units -{graphUnits} and exact graph units -{ExactGraphUnits} are equal.");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The graph units -{graphUnits} and exact graph units -{ExactGraphUnits} are not equal.");
                }
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Verify Normalization units has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Verify Normalization units has not been verified. The error is {e.Message}");
            }
        }
    }
}
