using System;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;
using SeleniumExtras.PageObjects;

namespace SHAProject.EditPage
{
    public class GraphSettings
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public string _currentPage = string.Empty;
        public CommonFunctions? _commonFunc;

        public GraphSettings(string currentPage, IWebDriver driver, FindElements findElements, CommonFunctions commonFunc)
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements = findElements;
            _commonFunc = commonFunc;
            PageFactory.InitElements(_driver, this);
        }

        #region Graph Setting Elements

        [FindsBy(How = How.XPath, Using = "(//img [@src='/images/svg/Settings.svg'])[1]")]
        public IWebElement GraphSettingIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id='graphSettings']/div/div")]
        public IWebElement GraphSettingsDisplayPopup;

        [FindsBy(How = How.XPath, Using = "//label[@for='zeroline']")]
        public IWebElement ZeroLineCheckBox;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettingsenergy")]
        public IWebElement DoseZerolineTag;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers']")]
        public IWebElement LineMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement LineMarkersTag;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement DoseLinemarkersTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphsettings.hid-rate-highlight")]
        public IWebElement RateHightlightTag;

        [FindsBy(How = How.CssSelector, Using = "//label[@for='ratehighlight']")]
        public IWebElement RateHightlightCheckBox;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(8)")]
        public IWebElement InjectionMarkersTag;

        [FindsBy(How = How.CssSelector, Using = "//label[@for='injectionmarkers']")]
        public IWebElement InjectionMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(9)")]
        public IWebElement zoomOptionTag;

        [FindsBy(How = How.XPath, Using = "(//div[@class='row g-set-form yaxis hidyaxis'][1])[1]")]
        public IWebElement YaxisMaxTag;

        [FindsBy(How = How.XPath, Using = "(//div[@class='row g-set-form yaxis hidyaxis'][2])[1]")]
        public IWebElement YaxisMinTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphATPsettings")]
        public IWebElement YintervalTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement TimeaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement TimeaxisMinTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphATPsettings.timeIntervalMTI")]
        public IWebElement TimeintervalTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(3)")]
        public IWebElement AutoscaleTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(4)")]
        public IWebElement XaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(5)")]
        public IWebElement XaxisMinTag;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement DoseaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement DoseaxisMinTag;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hiddoseaxis.dose-autoscale")]
        public IWebElement DoseAutoScaleTag;

        [FindsBy(How = How.CssSelector, Using ="#SaveDoseGraphSetting")]
        public IWebElement DoseApplyButton;

        [FindsBy(How = How.CssSelector, Using ="#Savesettings")]
        public IWebElement ApplyButton;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(9)")]
        public IWebElement linearScaleTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(10)")]
        public IWebElement logarithmicScaleTag;

        #endregion

        public void GraphSettingsIcon()
        {
            _findElements.ClickElementByJavaScript(GraphSettingIcon, _currentPage, $"Graph settings - Icon");

            _findElements.VerifyElement(GraphSettingsDisplayPopup, _currentPage, $"Graph Settings  Popup");
        }

        public void GraphSettingsApply()
        {
            _findElements.ClickElementByJavaScript(ApplyButton, _currentPage, $"Graph settings - Apply Button");
        }

        public void VerifyGraphSettings()
        {
            try
            {
                GraphSettingsIcon();

                IReadOnlyCollection<IWebElement> graphSettings = _driver.FindElements(By.CssSelector(".row.g-set-form"));

                foreach (IWebElement graphSetting in graphSettings)
                {
                    if (graphSetting.Displayed)
                    {
                        _findElements.ElementTextVerify(graphSetting, graphSetting.Text, _currentPage, $"Graph Setting - {graphSetting.Text}");
                    }
                }

                GraphSettingsApply();

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The graph settings elements text has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in graph settings elements text verification. The error is {e.Message}");
            }
        }

        public void GraphSettingsField(WidgetItems widget)
        {
            try
            {
                if (widget.GraphSettingsVerify)
                {
                    GraphSettingsIcon();

                    bool zeroLineStatus = ZeroLineCheckBox.Selected;
                    if (widget.GraphSettings.Zeroline != zeroLineStatus)
                    {
                        _findElements.ClickElementByJavaScript(ZeroLineCheckBox, _currentPage, $"Graph settings - Zeroline Checkbox");
                    }

                    bool lineMarkerStatus = LineMarkersCheckBox.Selected;
                    if (widget.GraphSettings.Linemarker != lineMarkerStatus)
                    {
                        _findElements.ClickElementByJavaScript(LineMarkersCheckBox, _currentPage, $"Graph settings - Linemarker Checkbox");
                    }

                    bool rateHighLightStatus = LineMarkersCheckBox.Selected;
                    if (widget.GraphSettings.RateHighlight != rateHighLightStatus)
                    {
                        _findElements.ClickElementByJavaScript(RateHightlightCheckBox, _currentPage, $"Graph settings - RateHighLight Checkbox");
                    }

                    bool injectionMarkersStatus = LineMarkersCheckBox.Selected;
                    if (widget.GraphSettings.InjectionMakers != injectionMarkersStatus)
                    {
                        _findElements.ClickElementByJavaScript(InjectionMarkersCheckBox, _currentPage, $"Graph settings - InjectionMarkers Checkbox");
                    }

                    GraphSettingsApply();
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The verification required for graph settings is given in the excel sheet is -{widget.GraphSettingsVerify}");
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in graph settings field verification. The error is {e.Message}");
            }
        }

        public void Zeroline()
        {

        }

        public void DoseZeroline()
        {
            _findElements.ElementTextVerify(DoseZerolineTag, "Zero Line", "Graph Settings - Dose Zeroline Tag", "Graph Settings - Dose Zeroline Tag");
        }

        public void Linemarkers()
        {

        }

        public void DoseLinemarkers()
        {

        }

        public void RateHighlight()
        {

        }

        public void InjectionMarkers()
        {

        }

        public void Zoom()
        {

        }

        public void LinearScale()
        {
            _findElements.ElementTextVerify(linearScaleTag, "Linear Scale", "Graph Settings - Linear Scale", "Graph Settings - Linear Scale");
        }

        public void LogarithmicScale()
        {
            _findElements.ElementTextVerify(logarithmicScaleTag, "Logarithmic Scale", "Graph Settings - Logarithmic Scale", "Graph Settings - Logarithmic Scale");
        }

        //public void YaxisMax()
        //{
        //     _findElements.ElementTextVerify(YaxisMaxTag, "Y Axis Max", "Graph Settings - Y Axis Max", "Graph Settings - Y Axis Max ");
        //}
        //public void YaxisMin()
        //{
        //      _findElements.ElementTextVerify(YaxisMinTag, "Y Axis Min", "Graph Settings -  Y Axis Min", "Graph Settings - Y Axis Min");
        //}

        //public void YInterval()
        //{
        //      _findElements.ElementTextVerify(YintervalTag, "Y  Interval", "Graph Settings -  YInterval", "Graph Settings - YInterval");
        //}

        //public void TimeAxisMax()
        //{
        //     _findElements.ElementTextVerify(TimeaxisMaxTag, "Time Axis Max", "Graph Settings -  Time Axis Max", "Graph Settings -  Time Axis Max");
        //}

        //public void TimeAxisMin()
        //{
        //    _findElements.ElementTextVerify(TimeaxisMinTag, "Time Axis Min", "Graph Settings -  Time Axis Min", "Graph Settings -  Time Axis Min");
        //}

        //public void TimeInterval()
        //{
        //     _findElements.ElementTextVerify(TimeintervalTag, "Time Interval", "Graph Settings -  Time Interval", "Graph Settings -  Time Interval");
        //}

        //public void AutoScale()
        //{
        //      _findElements.ElementTextVerify(AutoscaleTag, "Auto Scale", "Graph Settings - Auto Scale", "Graph Settings - Auto Scale");
        //}
        //public void XaxisMax()
        //{
        //      _findElements.ElementTextVerify(XaxisMaxTag, "X Axis Max", "Graph Settings - X Axis Max", "Graph Settings - X Axis Max");
        //}

        //public void XaxisMin()
        //{
        //      _findElements.ElementTextVerify(XaxisMinTag, "X Axis Min", "Graph Settings - X Axis Min", "Graph Settings - X Axis Min");
        //}

        //public void DoseAxisMax()
        //{
        //    _findElements.ElementTextVerify(DoseaxisMaxTag, "Dose Axis Max", "Graph Settings - Dose Axis Max", "Graph Settings - Dose Axis Max");
        //}

        //public void DoseAxisMin()
        //{
        //    _findElements.ElementTextVerify(DoseaxisMinTag, "Dose Axis Min", "Graph Settings - Dose Axis Min", "Graph Settings - Dose Axis Min");
        //}

        //public void DoseAutoScale()
        //{
        //    _findElements.ElementTextVerify(DoseAutoScaleTag, "Auto Scale", "Graph Settings - Auto Scale", "Graph Settings - Auto Scale");
        //}

    }
}
