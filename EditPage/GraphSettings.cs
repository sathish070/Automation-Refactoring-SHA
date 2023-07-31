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

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"dosegraph-settings-popup hidsettings\"])[1]")]
        public IWebElement DoseGraphSettingIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id='doseGraphSettings']/div/div")]
        public IWebElement DoseGraphSettingsDisplayPopup;

        [FindsBy(How = How.XPath, Using = "(//img [@src='/images/svg/Settings.svg'])[2]")]
        public IWebElement DosekineticGraphSettingIcon;

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

        #endregion

        #region CheckBox Fields Elements

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(6)")]
        public IWebElement XAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Xautoscale-energy']")]
        public IWebElement XAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(3)")]
        public IWebElement YAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale']")]
        public IWebElement YAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "//div[@id=\"yautoscaleenergy-settings\"]")]
        public IWebElement YEnergyAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale-energy']")]
        public IWebElement YAutoScaleEnergyCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(8)")]
        public IWebElement ZeroLineField;

        [FindsBy(How = How.XPath, Using = "//label[@for='zeroline']")]
        public IWebElement ZeroLineCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement DataPointSymbolsField;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers']")]
        public IWebElement DataPointSymbolsCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphsettings.hid-rate-highlight")]
        public IWebElement RateHighlightField;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"input-group graphsettings\"])[6]/label")]
        public IWebElement RateHighlightCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(8)")]
        public IWebElement InjectionMarkersField;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"input-group graphsettings\"])[7]/label")]
        public IWebElement InjectionMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(9)")]
        public IWebElement ZoomField;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"input-group graphsettings\"])[8]/label")]
        public IWebElement ZoomCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettingsenergy")]
        public IWebElement DoseZerolineField;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement DoseDataPointsSymbolsField;

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
        public IWebElement DoseDataPointSymbolsField;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers_dose']")]
        public IWebElement DoseDataPointSymbolsCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(8)")]
        public IWebElement DoseZoomField;

        [FindsBy(How = How.CssSelector, Using = "//label[@for='zoom_dose']")]
        public IWebElement DoseZoomCheckBox;

        #endregion

        #region Axis Max and Min Fields Elements

        [FindsBy(How = How.XPath, Using = "(//div[@class='row g-set-form yaxis hidyaxis'][1])[1]")]
        public IWebElement YaxisMaxTag;

        [FindsBy(How = How.XPath, Using = "(//div[@class='row g-set-form yaxis hidyaxis'][2])[1]")]
        public IWebElement YaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphATPsettings")]
        public IWebElement YintervalTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement TimeaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement TimeaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphATPsettings.timeIntervalMTI")]
        public IWebElement TimeintervalTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(4)")]
        public IWebElement XaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(5)")]
        public IWebElement XaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement DoseaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement DoseaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(9)")]
        public IWebElement linearScaleTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(10)")]
        public IWebElement logarithmicScaleTag;

        #endregion

        #region Data Table Elements

        [FindsBy(How = How.CssSelector, Using = "[title='Feature Chooser']")]
        public IWebElement DataTableSettingticon;

        [FindsBy(How = How.Id, Using = "ColumnChooser_Modal")]
        public IWebElement DataTableSettingtPopupWindow;

        [FindsBy(How = How.CssSelector, Using = "#ColumnChooser_ModalLabel")]
        public IWebElement ColumnChooserText;

        [FindsBy(How = How.CssSelector, Using = "#ColumnChooser_Modal .row .col-lg-12")]
        public IWebElement SelectAllText;

        [FindsBy(How = How.XPath, Using = "(//[aria-label='Close'][3]")]
        public IWebElement DataTableSettingtPopupWindowCloseIcon;

        [FindsBy(How = How.CssSelector, Using = ".ui-iggrid-header.ui-widget-header.ui-iggrid-multiheader-cell.ui-draggable.ui-iggrid-headercell-featureenabled")]
        public IList<IWebElement> DataTableWidgetList { get; set; }

        #endregion

        public void VerifyGraphSettingsIcon()
        {
            _findElements.ClickElementByJavaScript(GraphSettingIcon, _currentPage, $"Graph settings - Icon");

            _findElements.VerifyElement(GraphSettingsDisplayPopup, _currentPage, $"Graph Settings  Popup");
        }

        public void VerifyDoseGraphSettingsIcon()
        {
            _findElements.ClickElementByJavaScript(GraphSettingIcon, _currentPage, $"Dose Graph settings - Icon");

            _findElements.VerifyElement(DoseGraphSettingsDisplayPopup, _currentPage, $"Dose Graph Settings Popup");
        }

        public void VerifyDoseKineticGraphSettingsIcon()
        {
            _findElements.ClickElementByJavaScript(DosekineticGraphSettingIcon, _currentPage, $"Dose kinetic Graph settings - Icon");

            _findElements.VerifyElement(GraphSettingsDisplayPopup, _currentPage, $"Dose kinetic Graph Settings Popup");
        }


        #region CheckBox Fields

        public void XAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(XAutoScaleField, "Auto Scale", _currentPage, "Graph Setting - X Auto Scale");

            VerifySelectCheckBox(XAutoScaleField, XAutoScaleCheckBox, widget.GraphSettings.RemoveXAutoScale, "X AutoScale");
        }

        public void YAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(YAutoScaleField, "Auto Scale", _currentPage, "Graph Setting - Y Auto Scale");

            VerifySelectCheckBox(YAutoScaleField, YAutoScaleCheckBox, widget.GraphSettings.RemoveYAutoScale, "Y AutoScale");
        }

        public void YEnergyAutoScale(WidgetItems widget)
        {
            _findElements.ElementTextVerify(YEnergyAutoScaleField, "Auto Scale", _currentPage, "Graph Setting - Y Auto Scale");

            VerifySelectCheckBox(YEnergyAutoScaleField, YAutoScaleEnergyCheckBox, widget.GraphSettings.RemoveYAutoScale, "Y AutoScale");
        }

        public void ZeroLine(WidgetItems widget)
        {
            _findElements.ElementTextVerify(ZeroLineField, "Zero Line", _currentPage, "Graph Setting - Zero Line");

            VerifySelectCheckBox(ZeroLineField, ZeroLineCheckBox, widget.GraphSettings.RemoveZeroLine, "Zero Line");
        }

        public void DataPointSymbols(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DataPointSymbolsField, "Data Point Symbols", _currentPage, "Graph Setting - Data Point Symbols");

            VerifySelectCheckBox(DataPointSymbolsField, DataPointSymbolsCheckBox, widget.GraphSettings.RemoveDataPointSymbols, "Data Point Symbols");
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

        public void DoseDataPointsSymbols(WidgetItems widget)
        {
            _findElements.ElementTextVerify(DoseDataPointSymbolsField, "Dose Data Point Symbols", _currentPage, "Graph Setting - Dose Data Point Symbols");

            VerifySelectCheckBox(DoseDataPointSymbolsField, DoseDataPointSymbolsCheckBox, widget.GraphSettings.RemoveDoseDataPointSymbols, "Dose Data Point Symbols");
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
                    ScreenShot.ScreenshotNow(_driver, _currentPage, $"Graph settings - {ChkboxElement}", ScreenshotType.Info, fieldElement);
                }
                else if (expectedStatus != defaultSatus)
                {
                    bool status = _findElements.ClickElementByJavaScript(ChkboxElement, _currentPage, $"Graph settings - {propertyName} Checkbox");
                    if (status)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Expected {propertyName} - {expectedStatus} was selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"Graph settings - {ChkboxElement}", ScreenshotType.Info, fieldElement);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Expected {propertyName} - {expectedStatus} was not selected.");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, $"Graph settings - {ChkboxElement}", ScreenshotType.Error, fieldElement);
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

        public void VerifyDataTableSettings()
        {
            try
            {

                _findElements.ClickElementByJavaScript(DataTableSettingticon, _currentPage, $"DataTable settings - Icon");

                _findElements.VerifyElement(ColumnChooserText, _currentPage, "ColumnChooser Text");

                IWebElement selectAll = _driver.FindElements(By.CssSelector(".grid_groupnames")).First();
                _findElements.VerifyElement(selectAll, _currentPage, "selectAll Text");
                IReadOnlyCollection<IWebElement> DatatableGraphSettings = _driver.FindElements(By.CssSelector(".Group_Names .grid_groupnames"));
                foreach (IWebElement datatableGraphSettings in DatatableGraphSettings)
                {
                    if (datatableGraphSettings.Displayed)
                    {
                        _findElements.ElementTextVerify(datatableGraphSettings, datatableGraphSettings.Text, _currentPage, $"Graph Setting - {datatableGraphSettings.Text.Replace("/", "-")}");
                    }
                }

                var closeElements = _driver.FindElements(By.CssSelector("[aria-label='Close']"));
                closeElements[3].Click();
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in widget elements verification functionality. The error is {e.Message}");
            }
        }
    }
}
