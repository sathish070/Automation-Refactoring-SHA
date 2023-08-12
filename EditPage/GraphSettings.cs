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
        public Graph? graph;

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
        public IWebElement? GraphSettingIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id='graphSettings']/div/div")]
        public IWebElement? GraphSettingsDisplayPopup;

        [FindsBy(How = How.CssSelector, Using = "#Savesettings")]
        public IWebElement? ApplyButton;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"dosegraph-settings-popup hidsettings\"])[1]")]
        public IWebElement? DoseGraphSettingIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id='doseGraphSettings']/div/div")]
        public IWebElement? DoseGraphSettingsDisplayPopup;

        [FindsBy(How = How.XPath, Using = "(//img [@src='/images/svg/Settings.svg'])[2]")]
        public IWebElement? DosekineticGraphSettingIcon;

        [FindsBy(How = How.CssSelector, Using = "#SaveDoseGraphSetting")]
        public IWebElement? DoseApplyButton;

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

        [FindsBy(How = How.Id, Using = "yaxismax-settings")]
        public IWebElement? YAxisMaxField;

        [FindsBy(How = How.Id, Using = "yaxismin-settings")]
        public IWebElement? YAxisMinField;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(6)")]
        public IWebElement? XAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Xautoscale-energy']")]
        public IWebElement? XAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(3)")]
        public IWebElement? YAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale']")]
        public IWebElement? YAutoScaleCheckBox;

        [FindsBy(How = How.Id, Using = "yaxisenergymax-settings")]
        public IWebElement? EnergyYAxisMaxField;

        [FindsBy(How = How.Id, Using = "yaxisenergymin-settings")]
        public IWebElement? EnergyYAxisMinField;

        [FindsBy(How = How.CssSelector, Using = "//div[@id=\"yautoscaleenergy-settings\"]")]
        public IWebElement? YEnergyAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale-energy']")]
        public IWebElement? YAutoScaleEnergyCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(8)")]
        public IWebElement? ZeroLineField;

        [FindsBy(How = How.XPath, Using = "//label[@for='zeroline']")]
        public IWebElement? ZeroLineCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement? DataPointSymbolsField;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers']")]
        public IWebElement? DataPointSymbolsCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphsettings.hid-rate-highlight")]
        public IWebElement? RateHighlightField;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"input-group graphsettings\"])[6]/label")]
        public IWebElement? RateHighlightCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(8)")]
        public IWebElement? InjectionMarkersField;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"input-group graphsettings\"])[7]/label")]
        public IWebElement? InjectionMarkersCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(9)")]
        public IWebElement? ZoomField;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"input-group graphsettings\"])[8]/label")]
        public IWebElement? ZoomCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettingsenergy")]
        public IWebElement? DoseZerolineField;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphsettings")]
        public IWebElement? DoseDataPointsSymbolsField;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(3)")]
        public IWebElement? DoseXAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='Yautoscale_dose']")]
        public IWebElement? DoseXAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = ".dose-autoscale")]
        public IWebElement? DoseYAutoScaleField;

        [FindsBy(How = How.XPath, Using = "//label[@for='dose-autoscale']")]
        public IWebElement? DoseYAutoScaleCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(8)")]
        public IWebElement? DoseZeroLineField;

        [FindsBy(How = How.XPath, Using = "//label[@for='zeroline_dose']")]
        public IWebElement? DoseZeroLineCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(9)")]
        public IWebElement? DoseDataPointSymbolsField;

        [FindsBy(How = How.XPath, Using = "//label[@for='linemarkers_dose']")]
        public IWebElement? DoseDataPointSymbolsCheckBox;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(8)")]
        public IWebElement? DoseZoomField;

        [FindsBy(How = How.XPath, Using = "//label[@for='zoom_dose']")]
        public IWebElement? DoseZoomCheckBox;

        #endregion

        #region Axis Max and Min Fields Elements

        [FindsBy(How = How.XPath, Using = "(//div[@class='row g-set-form yaxis hidyaxis'][1])[1]")]
        public IWebElement? YaxisMaxTag;

        [FindsBy(How = How.XPath, Using = "(//div[@class='row g-set-form yaxis hidyaxis'][2])[1]")]
        public IWebElement? YaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div.row.g-set-form.hidgraphATPsettings")]
        public IWebElement? YintervalTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement? TimeaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement? TimeaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div.row.g-set-form.hidgraphATPsettings.timeIntervalMTI")]
        public IWebElement? TimeintervalTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(4)")]
        public IWebElement? XaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using = "#graphSettings > div > div > div.modal-body > div.graph-setting-options.left > div:nth-child(5)")]
        public IWebElement? XaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(1)")]
        public IWebElement? DoseaxisMaxTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(2)")]
        public IWebElement? DoseaxisMinTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(9)")]
        public IWebElement? linearScaleTag;

        [FindsBy(How = How.CssSelector, Using = "#doseGraphSettings > div > div > div.modal-body > div.graph-setting-options.right > div:nth-child(10)")]
        public IWebElement? logarithmicScaleTag;

        [FindsBy(How = How.Id, Using = "Ymax")]
        public IWebElement? YMaxTextBox;

        [FindsBy(How = How.Id, Using = "Ymin")]
        public IWebElement? YMinTextBox;

        [FindsBy(How = How.Id, Using = "Ymax-energy")]
        public IWebElement? EnergyMapYMaxTextBox;

        [FindsBy(How = How.Id, Using = "Ymin-energy")]
        public IWebElement? EnergyMapYMinTextBox;
        #endregion

        #region Data Table Elements

        [FindsBy(How = How.CssSelector, Using = "[title='Feature Chooser']")]
        public IWebElement? DataTableSettingticon;

        [FindsBy(How = How.Id, Using = "ColumnChooser_Modal")]
        public IWebElement? DataTableSettingtPopupWindow;

        [FindsBy(How = How.CssSelector, Using = "#ColumnChooser_ModalLabel")]
        public IWebElement? ColumnChooserText;

        [FindsBy(How = How.CssSelector, Using = "#ColumnChooser_Modal .row .col-lg-12")]
        public IWebElement? SelectAllText;

        [FindsBy(How = How.XPath, Using = "(//[aria-label='Close'][3]")]
        public IWebElement? DataTableSettingtPopupWindowCloseIcon;

        [FindsBy(How = How.CssSelector, Using = ".ui-iggrid-header.ui-widget-header.ui-iggrid-multiheader-cell.ui-draggable.ui-iggrid-headercell-featureenabled")]
        public IList<IWebElement?> DataTableWidgetList { get; set; }

        #endregion

        public void GraphInitialize()
        {
            graph = new(_currentPage, _driver, _findElements, _commonFunc);
        }

        public void VerifyGraphSettingsIcon()
        {
            GraphInitialize();

            _findElements.ClickElementByJavaScript(GraphSettingIcon, _currentPage, $"Graph settings - Icon");

            _findElements.VerifyElement(GraphSettingsDisplayPopup, _currentPage, $"Graph Settings Popup");
        }

        public void VerifyDoseGraphSettingsIcon()
        {
            GraphInitialize();

            _findElements.ClickElementByJavaScript(GraphSettingIcon, _currentPage, $"Dose Graph settings - Icon");

            _findElements.VerifyElement(DoseGraphSettingsDisplayPopup, _currentPage, $"Dose Graph Settings Popup");
        }

        public void VerifyDoseKineticGraphSettingsIcon()
        {
            GraphInitialize();

            _findElements.ClickElementByJavaScript(DosekineticGraphSettingIcon, _currentPage, $"Dose kinetic Graph settings - Icon");

            _findElements.VerifyElement(GraphSettingsDisplayPopup, _currentPage, $"Dose kinetic Graph Settings Popup");
        }

        #region CheckBox Fields

        public void XAutoScale(WidgetItems widget)
        {
            if (!GraphSettingsDisplayPopup.Displayed)
                VerifyGraphSettingsIcon();

            _findElements.ElementTextVerify(XAutoScaleField, "Auto Scale", _currentPage, "Graph Setting - X Auto Scale");

            VerifySelectCheckBox(XAutoScaleField, XAutoScaleCheckBox, widget.GraphSettings.RemoveXAutoScale, "X AutoScale");

            GraphSettingsApply();
        }

        public void ZeroLine(WidgetItems widget)
        {
            if (!GraphSettingsDisplayPopup.Displayed)
                VerifyGraphSettingsIcon();

            _findElements.ElementTextVerify(ZeroLineField, "Zero Line", _currentPage, "Graph Setting - Zero Line");

            VerifySelectCheckBox(ZeroLineField, ZeroLineCheckBox, widget.GraphSettings.RemoveZeroLine, "Zero Line");

            GraphSettingsApply();
        }

        public void DataPointSymbols(WidgetItems widget)
        {
            if (!GraphSettingsDisplayPopup.Displayed)
                VerifyGraphSettingsIcon();

            _findElements.ElementTextVerify(DataPointSymbolsField, "Data Point Symbols", _currentPage, "Graph Setting - Data Point Symbols");

            VerifySelectCheckBox(DataPointSymbolsField, DataPointSymbolsCheckBox, widget.GraphSettings.RemoveDataPointSymbols, "Data Point Symbols");

            GraphSettingsApply();
        }

        public void RateHighlight(WidgetItems widget)
        {
            if (!GraphSettingsDisplayPopup.Displayed)
                VerifyGraphSettingsIcon();

            _findElements.ElementTextVerify(RateHighlightField, "Rate Highlight", _currentPage, "Graph Setting - Rate Highlight");

            VerifySelectCheckBox(RateHighlightField, RateHighlightCheckBox, widget.GraphSettings.RemoveRateHighlight, "Rate Highlight");

            GraphSettingsApply();
        }

        public void InjectionMarkers(WidgetItems widget)
        {
            if (!GraphSettingsDisplayPopup.Displayed)
                VerifyGraphSettingsIcon();

            _findElements.ElementTextVerify(InjectionMarkersField, "Injection Markers", _currentPage, "Graph Setting - Injection Markers");

            VerifySelectCheckBox(InjectionMarkersField, InjectionMarkersCheckBox, widget.GraphSettings.RemoveInjectionMarkers, "Injection Markers");

            GraphSettingsApply();
        }

        public void Zoom(WidgetItems widget)
        {
            if (!GraphSettingsDisplayPopup.Displayed)
                VerifyGraphSettingsIcon();

            _findElements.ElementTextVerify(ZoomField, "Zoom", _currentPage, "Graph Setting - Zoom");

            VerifySelectCheckBox(ZoomField, ZoomCheckBox, widget.GraphSettings.RemoveZoom, "Zoom");

            GraphSettingsApply();
        }

        // Dose graph settings
        public void DoseXAutoScale(WidgetItems widget)
        {
            if (!DoseGraphSettingsDisplayPopup.Displayed)
                VerifyDoseGraphSettingsIcon();

            _findElements.ElementTextVerify(DoseXAutoScaleField, "DoseX AutoScale", _currentPage, "Graph Setting - DoseX AutoScale");

            VerifySelectCheckBox(DoseXAutoScaleField, DoseXAutoScaleCheckBox, widget.GraphSettings.RemoveDoseXAutoScale, "DoseX AutoScale");

            DoseGraphSettingsApply();
        }

        public void DoseYAutoScale(WidgetItems widget)
        {
            if (!DoseGraphSettingsDisplayPopup.Displayed)
                VerifyDoseGraphSettingsIcon();

            _findElements.ElementTextVerify(DoseYAutoScaleField, "DoseY AutoScale", _currentPage, "Graph Setting - DoseY AutoScale");

            VerifySelectCheckBox(DoseYAutoScaleField, DoseYAutoScaleCheckBox, widget.GraphSettings.RemoveDoseYAutoScale, "DoseY AutoScale");

            DoseGraphSettingsApply();
        }

        public void DoseZeroLine(WidgetItems widget)
        {
            if (!DoseGraphSettingsDisplayPopup.Displayed)
                VerifyDoseGraphSettingsIcon();

            _findElements.ElementTextVerify(DoseZeroLineField, "Dose Zero Line", _currentPage, "Graph Setting - Dose Zero Line");

            VerifySelectCheckBox(DoseZeroLineField, DoseZeroLineCheckBox, widget.GraphSettings.RemoveDoseZeroLine, "Dose Zero Line");

            DoseGraphSettingsApply();
        }

        public void DoseDataPointsSymbols(WidgetItems widget)
        {
            if (!DoseGraphSettingsDisplayPopup.Displayed)
                VerifyDoseGraphSettingsIcon();

            _findElements.ElementTextVerify(DoseDataPointSymbolsField, "Dose Data Point Symbols", _currentPage, "Graph Setting - Dose Data Point Symbols");

            VerifySelectCheckBox(DoseDataPointSymbolsField, DoseDataPointSymbolsCheckBox, widget.GraphSettings.RemoveDoseDataPointSymbols, "Dose Data Point Symbols");

            DoseGraphSettingsApply();
        }

        public void DoseZoom(WidgetItems widget)
        {
            if (!DoseGraphSettingsDisplayPopup.Displayed)
                VerifyDoseGraphSettingsIcon();

            _findElements.ElementTextVerify(DoseZoomField, "Dose Zoom", _currentPage, "Graph Setting - Dose Zoom");

            VerifySelectCheckBox(DoseZoomField, DoseZoomCheckBox, widget.GraphSettings.RemoveDoseZoom, "Dose Zoom");

            DoseGraphSettingsApply();
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

        public void GraphSettingsApply()
        {
            _findElements.ClickElementByJavaScript(ApplyButton, _currentPage, $"Graph settings - Apply Button");

            Thread.Sleep(2000);
        }

        public void DoseGraphSettingsApply()
        {
            _findElements.ClickElementByJavaScript(DoseApplyButton, _currentPage, $"Dose Graph settings - Apply Button");

            Thread.Sleep(2000);
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

        public void YAutoScale(WidgetItems widget)
        {
            try
            {
                GraphInitialize();

                Thread.Sleep(4000);

                 (double maxValue, double minValue, List<double> doubles)= graph.GraphYmaxYminVerification();

                double graphAutoscaleIncValue = maxValue + (maxValue * 20 / 100);
                double graphAutoscaleDecValue = minValue == 0 ? 1 : minValue + (minValue * 20 / 100);

                if (!GraphSettingsDisplayPopup.Displayed)
                    VerifyGraphSettingsIcon();
               
                if (YAxisMaxField.Displayed)
                {
                    _findElements.ElementTextVerify(YAxisMaxField, "Y Axis Max", _currentPage, $"Graph Setting - Y Axis Max");

                    _findElements.ElementTextVerify(YAxisMinField, "Y Axis Min", _currentPage, $"Graph Setting - Y Axis Min");

                    _findElements.ElementTextVerify(YAutoScaleField, "Auto Scale", _currentPage, $"Graph Setting - Y Auto Scale");

                    _findElements.VerifyElement(YMaxTextBox, _currentPage, $"Default Y-Axis Max value in the graph settings");

                    _findElements.VerifyElement(YMinTextBox, _currentPage, $"Default Y-Axis Min value in the graph settings");

                    VerifySelectCheckBox(YAutoScaleField, YAutoScaleCheckBox, widget.GraphSettings.RemoveYAutoScale, $"Energy Map Y AutoScale");

                    _findElements.SendKeys(graphAutoscaleIncValue.ToString(), YMaxTextBox, _currentPage, $"The Given Y-Axis Max Value is {graphAutoscaleIncValue}");

                    _findElements.SendKeys(graphAutoscaleDecValue.ToString(), YMinTextBox, _currentPage, $"The Given Y-Axis Min Value is {graphAutoscaleDecValue}");
                }
                else
                {
                    _findElements.ElementTextVerify(EnergyYAxisMaxField, "Y Axis Max", _currentPage, $"Graph Setting - Energy Map Y Axis Max");

                    _findElements.ElementTextVerify(EnergyYAxisMinField, "Y Axis Min", _currentPage, $"Graph Setting - Energy Map Y Axis Min");

                    _findElements.ElementTextVerify(YEnergyAutoScaleField, "Auto Scale", _currentPage, $"Graph Setting - Energy Y Auto Scale");

                    _findElements.VerifyElement(EnergyMapYMaxTextBox, _currentPage, $"Default Energy Map Y-Axis Max value in the graph settings");

                    _findElements.VerifyElement(EnergyMapYMinTextBox, _currentPage, $"Default Energy Map Y-Axis Min value in the graph settings");

                    VerifySelectCheckBox(YEnergyAutoScaleField, YAutoScaleEnergyCheckBox, widget.GraphSettings.RemoveYAutoScale, $"Energy Map Y AutoScale");

                    _findElements.SendKeys(graphAutoscaleIncValue.ToString(), EnergyMapYMaxTextBox, _currentPage, $"The Given Y-Axis Max Value is {graphAutoscaleIncValue}");

                    _findElements.SendKeys(graphAutoscaleDecValue.ToString(), EnergyMapYMinTextBox, _currentPage, $"The Given Y-Axis Min Value is {graphAutoscaleDecValue}");
                }

                GraphSettingsApply();

                Thread.Sleep(8000);
                graph.GraphYmaxYminVerification();
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error Occurred while verifying the Y- Axis and Auto Scale in the graph settings. The error is {e.Message}");
            }
        }
    }
}
