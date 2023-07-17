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
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;

        public GraphSettings(string currentPage, IWebDriver driver, FindElements findElements, CommonFunctions commonFunc)
        {
            _driver = driver;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            PageFactory.InitElements(_driver, this);
        }

        #region Graph Settings Elements

        [FindsBy(How = How.XPath, Using = "(//img [@src='/images/svg/Settings.svg'])[1]")]
        public IWebElement GraphSettingIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id='graphSettings']/div/div")]
        public IWebElement GraphSettingsDisplayPopup;

        [FindsBy(How = How.CssSelector, Using = "#Savesettings")]
        public IWebElement ApplyButton;

        [FindsBy(How = How.CssSelector, Using = ".dosegraph-settings-popup")]
        public IWebElement DoseGraphSettingIcon;

        [FindsBy(How = How.XPath, Using = "/html/body/section/div[3]/div/div")]
        public IWebElement DoseGraphSettingsDisplayPopup;

        [FindsBy(How = How.CssSelector, Using = "#SaveDoseGraphSetting")]
        public IWebElement DoseApplyButton;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(3) > button:nth-child(1)")]
        public IWebElement? SyncToView;

        [FindsBy(How = How.CssSelector, Using = "#GraphSettingsSyncView > div:nth-child(1) > div:nth-child(1)")]
        public IWebElement? SyncToViewPopup;

        [FindsBy(How = How.CssSelector, Using = "#GraphSettingsSyncView > div:nth-child(1) > div:nth-child(1) > div:nth-child(3) > button:nth-child(1)")]
        public IWebElement? SyncToViewApplyButton;

        [FindsBy(How = How.CssSelector, Using = ".syncviewresult-tost-success")]
        public IWebElement? SyncToViewToast;

        // CheckBoxs Fields
        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(6)")]
        public IWebElement XAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale-energy']")]
        public IWebElement XAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(3)")]
        public IWebElement YAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale']")]
        public IWebElement YAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(8)")]
        public IWebElement ZeroLineField;

        [FindsBy(How = How.XPath, Using = "//label[@for='zeroline']")]
        public IWebElement ZeroLineCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement LineMarkersField;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers']")]
        public IWebElement LineMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphsettings.hid-rate-highlight")]
        public IWebElement RateHighlightField;

        [FindsBy(How = How.CssSelector, Using = "//label[@for='ratehighlight']")]
        public IWebElement RateHighlightCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(8)")]
        public IWebElement InjectionMarkersField;

        [FindsBy(How = How.CssSelector, Using = "//label[@for='injectionmarkers']")]
        public IWebElement InjectionMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(9)")]
        public IWebElement ZoomField;           

        [FindsBy(How = How.CssSelector, Using = "//label[@for='zoom']")]
        public IWebElement ZoomCheckBox;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettingsenergy")]
        public IWebElement DoseZerolineField;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement DoseLinemarkersField;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(3)")]
        public IWebElement DoseXAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale_dose']")]
        public IWebElement DoseXAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = ".dose-autoscale")]
        public IWebElement DoseYAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='dose-autoscale']")]
        public IWebElement DoseYAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(8)")]
        public IWebElement DoseZeroLineField;

        [FindsBy(How = How.XPath, Using = "//label[@for='zeroline_dose']")]
        public IWebElement DoseZeroLineCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(9)")]
        public IWebElement DoseLineMarkersField;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers_dose']")]
        public IWebElement DoseLineMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(8)")]
        public IWebElement DoseZoomField;

        [FindsBy(How = How.CssSelector, Using = "//label[@for='zoom_dose']")]
        public IWebElement DoseZoomCheckBox;

        // Axis Min Max Fields
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

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(4)")]
        public IWebElement XaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using ="#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(5)")]
        public IWebElement XaxisMinTag;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement DoseaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using ="#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement DoseaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(9)")]
        public IWebElement linearScaleTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(10)")]
        public IWebElement logarithmicScaleTag;

        #endregion

        public void VerifyGraphSettingsIcon()
        {
            _findElements.ClickElementByJavaScript(GraphSettingIcon, _currentPage, $"Graph settings - Icon");

            _findElements.VerifyElement(GraphSettingsDisplayPopup, _currentPage, $"Graph Settings  Popup");
        }

        public void VerifyDoseGraphSettingsIcon()
        {
            _findElements.ClickElementByJavaScript(DoseGraphSettingIcon, _currentPage, $"Dose Graph settings - Icon");

            _findElements.VerifyElement(DoseGraphSettingsDisplayPopup, _currentPage, $"Dose Graph Settings Popup");
        }

        #region CheckBox Fields

        public void XAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(XAutoScaleField, "AutoScale", _currentPage, "Graph Setting - X AutoScale");

            VerifySelectCheckBox(XAutoScaleField, XAutoScaleCheckBox, widget.GraphSettings.RemoveXAutoScale, "X AutoScale");
        }

        public void YAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(YAutoScaleField, "AutoScale", _currentPage, "Graph Setting - Y AutoScale");

            VerifySelectCheckBox(YAutoScaleField, YAutoScaleCheckBox, widget.GraphSettings.RemoveYAutoScale, "Y AutoScale");
        }

        public void ZeroLine(WidgetItems widget)
        {
            _findElements.ElementTextVerify(ZeroLineField, "Zero Line", _currentPage, "Graph Setting - Zero Line");

            VerifySelectCheckBox(ZeroLineField, ZeroLineCheckBox, widget.GraphSettings.RemoveZeroLine, "Zero Line");
        }

        public void LineMarkers(WidgetItems widget)
        {
            _findElements.ElementTextVerify(LineMarkersField, "Line Markers", _currentPage, "Graph Setting - Line Markers");

            VerifySelectCheckBox(LineMarkersField, LineMarkersCheckBox, widget.GraphSettings.RemoveLineMarkers, "Line Markers");
        }

        public void RateHighlight(WidgetItems widget)
        {
            _findElements.ElementTextVerify(RateHighlightField, "Rate Highlight", _currentPage, "Graph Setting - Rate Highlight");

            VerifySelectCheckBox(RateHighlightField, RateHighlightCheckBox, widget.GraphSettings.RemoveRateHighlight, "Rate Highlight");
        }

        public void InjectionMarkers(WidgetItems widget)
        {
            _findElements.ElementTextVerify(InjectionMarkersField, "Injection Markers", _currentPage, "Graph Setting - Injection Markers");

            VerifySelectCheckBox(InjectionMarkersField, InjectionMarkersCheckBox, widget.GraphSettings.RemoveInjectionMarkers, "Injection Markers");
        }

        public void Zoom(WidgetItems widget)
        {
            _findElements.ElementTextVerify(ZoomField, "Zoom", _currentPage, "Graph Setting - Zoom");

            VerifySelectCheckBox(ZoomField, ZoomCheckBox, widget.GraphSettings.RemoveZoom, "Zoom");
        }

        // Dose graph settings
        public void DoseXAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DoseXAutoScaleField, "DoseX AutoScale", _currentPage, "Graph Setting - DoseX AutoScale");

            VerifySelectCheckBox(DoseXAutoScaleField, DoseXAutoScaleCheckBox, widget.GraphSettings.RemoveDoseXAutoScale, "DoseX AutoScale");
        }

        public void DoseYAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DoseYAutoScaleField, "DoseY AutoScale", _currentPage, "Graph Setting - DoseY AutoScale");

            VerifySelectCheckBox(DoseYAutoScaleField, DoseYAutoScaleCheckBox, widget.GraphSettings.RemoveDoseYAutoScale, "DoseY AutoScale");
        }

        public void DoseZeroLine(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DoseZeroLineField, "Dose Zero Line", _currentPage, "Graph Setting - Dose Zero Line");

            VerifySelectCheckBox(DoseZeroLineField, DoseZeroLineCheckBox, widget.GraphSettings.RemoveDoseZeroLine, "Dose Zero Line");
        }

        public void DoseLineMarkers(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DoseLineMarkersField, "Dose Line Markers", _currentPage, "Graph Setting - Dose Line Markers");

            VerifySelectCheckBox(DoseLineMarkersField, DoseLineMarkersCheckBox, widget.GraphSettings.RemoveDoseLineMarkers, "Dose Line Markers");
        }

        public void DoseZoom(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DoseZoomField, "Dose Zoom", _currentPage, "Graph Setting - Dose Zoom");

            VerifySelectCheckBox(DoseZoomField, DoseZoomCheckBox, widget.GraphSettings.RemoveDoseZoom, "Dose Zoom");
        }

        public void VerifySelectCheckBox(IWebElement fieldElement, IWebElement ChkboxElement, bool expectedStatus, string propertyName)
        {
            try
            {
                bool defaultSatus = ChkboxElement.Selected;

                if (expectedStatus == defaultSatus)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedStatus} and actual {defaultSatus} are equal.");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, "", ScreenshotType.Info, fieldElement);
                }
                else if (expectedStatus != defaultSatus)
                {
                    bool status = _findElements.ClickElementByJavaScript(ChkboxElement, _currentPage, $"Graph settings - {propertyName} Checkbox"); ;
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedStatus} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Info, fieldElement);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} - {expectedStatus} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "--", ScreenshotType.Error, fieldElement);
                    }
                }
            }
            catch (NoSuchElementException e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Element - {ChkboxElement} is not found on the page . The error is {e.Message}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} is not verified. The error is {e.Message}");
            }
        }

        #endregion

        public void VerifyGraphSettingsFields()
        {
            try
            {
                IReadOnlyCollection<IWebElement> graphSettings = _driver.FindElements(By.CssSelector(".row.g-set-form"));

                foreach (IWebElement graphSetting in graphSettings)
                {
                    if (graphSetting.Displayed)
                        _findElements.ElementTextVerify(graphSetting, graphSetting.Text, _currentPage, $"Graph Setting - {graphSetting.Text}");
                }

                GraphSettingsApply();

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The graph settings elements text has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in graph settings elements text verification. The error is {e.Message}");
            }
        }

        public void GraphSettingsApply()
        {
            _findElements.ClickElementByJavaScript(ApplyButton, _currentPage, $"Graph settings - Apply Button");
        }

        public void DoseGraphSettingsApply()
        {
            _findElements.ClickElementByJavaScript(DoseApplyButton, _currentPage, $"Dose Graph settings - Apply Button");
        }

        public void GraphSettingSyncToView()
        {
            _findElements.ClickElementByJavaScript(SyncToView, _currentPage, $"Graph Setting - Sync to View");

            _findElements.VerifyElement(SyncToViewPopup, _currentPage, $"Graph Setting - Sync to view Popup");

            _findElements.ClickElementByJavaScript(SyncToViewApplyButton, _currentPage, $"Graph Setting - Sync to view apply button");

            _findElements.VerifyElement(SyncToViewToast, _currentPage, $"Graph Setting - Sync to view Toast Message");
        }
    }
}
