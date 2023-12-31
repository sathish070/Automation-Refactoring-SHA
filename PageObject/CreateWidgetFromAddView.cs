﻿using System;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;
using SeleniumExtras.PageObjects;
using OpenQA.Selenium.Support.UI;

namespace SHAProject.Create_Widgets
{
    public class CreateWidgetFromAddView
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;


        public CreateWidgetFromAddView(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc)
        {
            _driver = driver;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            PageFactory.InitElements(_driver, this);
        }

        #region Common Element

        [FindsBy(How = How.XPath, Using = "//a[@id='menu-toggle-views']")]
        public IWebElement? SideViewMenuToggleButton;

        [FindsBy(How = How.ClassName, Using = "addnewlist")]
        public IWebElement? AddnewlistViewIcon;

        [FindsBy(How = How.Id, Using = "AddViewsModal")]
        public IWebElement? AddViewPopUp;

        [FindsBy(How = How.XPath, Using = "(//span[@class=\"caret\"])[3]")]
        public IWebElement? AssayKitCompanionViews;

        [FindsBy(How = How.Id, Using = "btnAddview")]
        public IWebElement? AddViewButton;

        [FindsBy(How = How.XPath, Using = "(//span[@class='caret'])[2]")]
        public IWebElement? CustomViewCompanionViews;

        [FindsBy(How = How.XPath, Using = "//div[@class='col-md-2 modal-right-groups']")]
        public IWebElement? AddviewGroups;

        [FindsBy(How = How.XPath, Using = "//span[@id=\"Listhead\"]")]
        public IWebElement? ListHeadName;

        [FindsBy(How = How.XPath, Using = "(//div[@class='group-list-right']/span)[1]")]
        public IWebElement? GroupListElement;

        [FindsBy(How = How.XPath, Using = "(//div[@class='comp-list-right']/span)[1]")]
        public IWebElement? CompundListElement;
        #endregion

        #region StandardView Element

        [FindsBy(How = How.CssSelector, Using = "#default-graphs .caret")]
        private IWebElement? DefaultGraphClick;

        [FindsBy(How = How.CssSelector, Using = "(//*[@class='caret'])[last()]\")")]
        private IWebElement? FilespageAddViewClick;

        [FindsBy(How = How.CssSelector, Using = "#quickView > section.col-md-12.graph-type-sec > div.graph-type-head > div > h5")]
        private IWebElement? quickViewAssayKitValidated;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-QuickView-OCR.svg?v=lWT4LXnVW_aybNcxjHQwbQDAEOCWt7U6kGwqlYCO-V4']")]
        private IWebElement? KineticGraphOCR;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-QuickView-ECAR.svg?v=F8NQ0gVoyYBBhlFsf9IWuzF3XPlZbVYv_Ee5nA0cC6M']")]
        private IWebElement? KineticGraphECAR;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-QuickView-PER.svg?v=-bvzYxhhu3-fy5yQHygYXZCkqTrjKB2-n9e9eGWdE80']")]
        private IWebElement? KineticGraphPER;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View Widget-QuickView-Bar Graph.svg']")]
        private IWebElement? BarChart;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-QuickView-Energetic-Map.svg?v=QYX9ZFC0YDT6YJm5mjKb4XN1BEzbchAc2QecBuJsThI']")]
        private IWebElement? EnergyMap;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Heat-Map.svg']")]
        private IWebElement? HeatMap;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Heat-Map-unavailable.svg']")]
        private IWebElement? UnavailableHeatMap;

        [FindsBy(How = How.XPath, Using = "(//li[@id='msv_my_view'])[1]")]
        private IWebElement? CustomView;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Dose-Response.svg']")]
        private IWebElement? DoseResponse;

        #endregion

        #region Dropdown Elements

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"graphsettingsarea views\"])[3]")]
        public IWebElement? OligoInjectionField;

        [FindsBy(How = How.Id, Using = "dlloligoinjection")]
        public IWebElement? OligoDropdown;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"graphsettingsarea views\"])[4]")]
        public IWebElement? InducedInjectionField;

        [FindsBy(How = How.Id, Using = "inducedinjection")]
        public IWebElement? InducedDropdown;

        #endregion

        #region XF Mito Stress View Elements

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Mitochondrial-Respiration.svg?v=8daQG0X_yW7X--6UYbGipk0mJfMKu9us10MUs0BD9Bo']")]
        private IWebElement MitochondrialRespirationWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Basal.svg?v=YSSMbflXURIa4IajzidnNeyp_l5JEmHfF10GDV_tGTM']")]
        private IWebElement BasalWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Acute-Response.svg?v=2Zurjczl9SPVpnhdFuynySjxIU_QexNEh3iRzTLVqFc']")]
        private IWebElement AcuteResponseWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Proton-Leak.svg?v=4a9FPiOyhA5oCV5o8V0MGZbSpvH2qPvXcp0qmcSGlb0']")]
        private IWebElement ProtonLeakWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Maximal-Respiration.svg?v=KAp2tUFKSUtWh9Hzgq9kCzOUFDmIDziJtvCAaJMOr48']")]
        private IWebElement MaximalRespirationWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Spare-Respiratory-Capacity.svg?v=V59uZVtsaMzkFLZqfXA9urFvONtLJHpWmbsLWGq8Hko']")]
        private IWebElement SpareRespiratoryCapacityWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Nonmito-02-Consumption.svg?v=eQcP8mDlVAXWbx2fsNZUCVYvqCkRJ1W9DJTAZvAtJdQ']")]
        private IWebElement NonMitoO2ConsumptionWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Linked-Respiration.svg?v=yv7NNkIkzpnucYFwCycazjSn6lROGmN2rGBY42v_xrY']")]
        private IWebElement AtpProductionCoupledRespirationWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Coupling-Efficiency.svg?v=merDLEQiSoDYM3Uc_2V-yWj7Q8uDoFekiHzzUoVrWTk']")]
        private IWebElement CouplingEfficiencyPercentwidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Spare-Respiratory-Capacity-pct.svg?v=bwKYF6U0yMnMNhUCiXBFLGj9YBHYH-Hvp-RX_LKQz9M']")]
        private IWebElement SpareRespiratoryCapacityPercentWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-MST-Data-Table.svg?v=xIaVxnW357Hq-YLlqjYEYUYT8KYFGrCLFtPmXwKplwo']")]
        private IWebElement DataTableWidget;

        [FindsBy(How = How.Id, Using = "dllmstoligoinjection")]
        private IWebElement MstOligoinjection;

        #endregion

        #region XF ATP Rate Assay View Elements

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-mitoATP-Production-Rate.svg?v=Y7DhwC72V_x9OIj8tdjX8Cw6EKFsrbu1DYTFqVpm86o']")]
        private IWebElement MitoAtpProductionRatewidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-glycoATP-Production-Rate.svg?v=ez78EL12iA1NDdC5oyEm2riuFKxmGQtrG1pAPh4-DUA']")]
        private IWebElement GlycoAtpProductionRatewidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-ATP-Production-Rate-Data.svg?v=yYPAIeIrsW3jxhEVFWwPsntgv_jmHhzGElzlRxR2ojY']")]
        private IWebElement AtpProductionRateDataWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-ATP-Production-Rate-Basal.svg?v=7vRj_ThXHbYIya_jrlyIbSsJqNV6ExiJnTyBpoDZ2YI']")]
        private IWebElement AtpProductionRateBasalWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-ATP-Production-Rate-Induced.svg?v=wgpT4uLuRBjxt9lVF7niL3gpvUA-9NAuw1qBLRZkql4']")]
        private IWebElement AtpProductionRateInducedWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-Energetic-Map-Basal.svg?v=6uh2nn8Othje2l8F48Rct7N6fCraoW1xy79Mi-8dhfA']")]
        private IWebElement EnergeticMapBasalWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-Energetic-Map-Induced.svg?v=CC3UgTL-evK_WCZIWVC6oOiaUqJ8e3e_REYGEKJlWRY']")]
        private IWebElement EnergeticMapInducedWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-XF-ATP-Rate-Index.svg?v=luG2YnbJhOe7ZP7--a1oh0nTvwNlPt2-F61-9WqAci0']")]
        private IWebElement XfAtpRateIndexWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-ATP-Data-Table-Basal-and-Induced.svg?v=fKWsLf4FHsZwTP11XeiKxcBnLytG8Q7dEUbnX2m0k8g']")]
        private IWebElement DataTableBasalandInduced;

        #endregion

        #region XF Cell Energy Phenotype Elements

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-XF-Cell-Energy-Phenotype.svg?v=YCmC2zY50DgephevUec7MU8pIjFQdoaVkE01W1_LPm0']")]
        private IWebElement CellEnergyPhenotypeWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Metabolic-Potential-OCR.svg?v=HQByQwHT5yS774v8NFo9BLXLZI1P1s9LASoOrHEYj0U']")]
        private IWebElement MetabolicPotentialOCRWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Metabolic-Potential-ECAR.svg?v=IidTPSnVsAE3E12ymARwnigY-A66cTDesDLehOt731M']")]
        private IWebElement MetabolicPotentialECARWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Baseline-OCR.svg?v=GdJJnRGPi7I86NtNemNbSS9YqSZjysxr5YAU9YG50fU']")]
        private IWebElement BaselineOCRWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Baseline-ECAR.svg?v=LqGhpYr3sq0RXTl8_8znZfuRuOD_y7RTzW0NsTGciYg']")]
        private IWebElement BaselineECARWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Stressed-OCR.svg?v=4glCYXFuWG25Hxa86DLpaeiSxnEnIQuzitbd_FJalFs']")]
        private IWebElement StressedOCRWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Stressed-ECAR.svg?v=eQy4t3WZr2QNWf-HOS7SXsKgxSGQWSAmxGcTngh9k18']")]
        private IWebElement StressedECARWidget;

        [FindsBy(How = How.CssSelector, Using = "[src='/images/svg/AddView/View-Widget-Cell-Pheno-Data-Table.svg?v=gfCEv5-VL9_l-yWRkk_ahX2TWgC42fQGyvHyht7kWkg']")]
        private IWebElement CellEnergyDataTableWidget;

        #endregion

        public void CreateWidgets(WidgetCategories wCat, List<WidgetTypes> SelectedWidgets)
        {
            try
            {
                Thread.Sleep(5000);

                bool Isdisplayed = AddViewPopUp.Displayed;
                if (Isdisplayed)
                {
                    AddView(wCat, SelectedWidgets);
                }
                else
                {
                    _commonFunc?.HandleCurrentWindow();

                    _findElements?.ClickElement(SideViewMenuToggleButton, _currentPage, $"Analysis Page - Side View Toggle Button");

                    _findElements?.ScrollIntoViewAndClickElementByJavaScript(AddnewlistViewIcon, _currentPage, $"Analysis Page - Add New List");

                    ScreenShot.ScreenshotNow(_driver, _currentPage, "Add View", ScreenshotType.Info);

                    AddView(wCat, SelectedWidgets);
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Error in AddView functionality {e.Message}");
            }
        }

        public void AddView(WidgetCategories wCat, List<WidgetTypes> SelectedWidgets)
        {
            IWebElement companionView = null;
            try
            {
                if (wCat != WidgetCategories.XfCustomview)
                    companionView = _driver.FindElement(By.XPath("//li[@data-viewname='" + GetCompanionViewName(wCat) + "']"));

                switch (wCat)
                {
                    case WidgetCategories.XfStandard:
                        _findElements?.ClickElement(DefaultGraphClick, _currentPage, $"Add View popup - Standard view");
                        _findElements?.ClickElement(companionView, _currentPage, "Add View popup -Standard view");
                        VerifyAddViewPopupElements();
                        break;

                    case WidgetCategories.XfCustomview:
                        _findElements?.ClickElement(CustomViewCompanionViews, _currentPage, $"Add View popup - Standard view");
                        _findElements?.ClickElement(CustomView, _currentPage, "Add View popup -Standard view");
                        break;

                    case WidgetCategories.XfStandardDose:
                        _findElements?.ClickElement(DefaultGraphClick, _currentPage, $"Add View popup - Standard view");
                        _findElements?.ClickElement(companionView, _currentPage, "Add View popup - Standard Dose view");
                        VerifyAddViewPopupElements();
                        break;

                    case WidgetCategories.XfStandardBlank:
                        _findElements?.ClickElement(DefaultGraphClick, _currentPage, $"Add View popup - Standard view");
                        _findElements?.ClickElement(companionView, _currentPage, "Add View popup - Blank view");
                        break;

                    case WidgetCategories.XfMst:
                        _findElements?.ClickElement(AssayKitCompanionViews, _currentPage, $"Add View popup - Assay Kit Companion Views");
                        _findElements?.ClickElement(companionView, _currentPage, "Add View popup - XF Mito stress View");
                        VerifyAddViewPopupElements();
                        DropdownSelect(_fileUploadOrExistingFileData.OligoInjection, MstOligoinjection, "Add View popup Oligo Dropdown");
                        break;

                    case WidgetCategories.XfAtp:
                        _findElements?.ClickElement(AssayKitCompanionViews, _currentPage, $"Add View popup - Assay Kit Companion Views");
                        _findElements?.ClickElement(companionView, _currentPage, "Add View popup - XF ATP Rate Assay View");
                        VerifyAddViewPopupElements();
                        DropdownVerification();
                        break;

                    case WidgetCategories.XfCellEnergy:
                        _findElements?.ClickElement(AssayKitCompanionViews, _currentPage, $"Add View popup -Assay Kit Companion Views");
                        _findElements?.ClickElement(companionView, _currentPage, "Add View popup - XFCellEnergyPhenotype view");
                        VerifyAddViewPopupElements();
                        break;
                }

                Dictionary<WidgetCategories, List<WidgetTypes>> widgetMappings = GetWidgetMappings();
                if (widgetMappings.ContainsKey(wCat))
                {
                    foreach (var widget in SelectedWidgets)
                    {
                        if(widget == WidgetTypes.HeatMap)
                        {
                            if (HeatMap.Displayed)
                            {
                                ClickWidgetElement(widget, wCat);
                            }
                        }
                        else
                        {
                            if (widgetMappings[wCat].Contains(widget))
                                ClickWidgetElement(widget, wCat);
                            else
                                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Widgets are not selected");
                        }
                    }
                }

                _findElements?.ClickElement(AddViewButton, _currentPage, $"AddView Popup - AddView Button");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Error in AddView functionality {e.Message}");
            }
        }

        public static string GetCompanionViewName(WidgetCategories wCat)
        {
            string? ViewName = wCat switch
            {
                WidgetCategories.XfStandard => "Quick View",
                WidgetCategories.XfStandardBlank => "New View",
                WidgetCategories.XfStandardDose => "Dose View",
                WidgetCategories.XfMst => "XF Cell Mito Stress Test View",
                WidgetCategories.XfSubOx => "XF Sub Ox Stress Test View",
                WidgetCategories.XfAtp => "XF ATP Rate Assay View",
                WidgetCategories.XfCellEnergy => "XF Cell Energy Phenotype View",
                _ => "",
            };
            return ViewName;
        }

        private Dictionary<WidgetCategories, List<WidgetTypes>> GetWidgetMappings()
        {
            Dictionary<WidgetCategories, List<WidgetTypes>> widgetMappings = new Dictionary<WidgetCategories, List<WidgetTypes>>();

            // Added widget mappings here
            widgetMappings.Add(WidgetCategories.XfStandard, new List<WidgetTypes>()
            {
                WidgetTypes.KineticGraph,
                WidgetTypes.KineticGraphEcar,
                WidgetTypes.KineticGraphPer,
                WidgetTypes.BarChart,
                WidgetTypes.EnergyMap,
                WidgetTypes.HeatMap
            });

            widgetMappings.Add(WidgetCategories.XfStandardDose, new List<WidgetTypes>()
            {
                WidgetTypes.DoseResponse,
            });

            widgetMappings.Add(WidgetCategories.XfMst, new List<WidgetTypes>()
            {
                WidgetTypes.MitochondrialRespiration,
                WidgetTypes.Basal,
                WidgetTypes.AcuteResponse,
                WidgetTypes.ProtonLeak,
                WidgetTypes.MaximalRespiration,
                WidgetTypes.SpareRespiratoryCapacity,
                WidgetTypes.NonMitoO2Consumption,
                WidgetTypes.AtpProductionCoupledRespiration,
                WidgetTypes.CouplingEfficiencyPercent,
                WidgetTypes.SpareRespiratoryCapacityPercent,
                WidgetTypes.DataTable
            });

            widgetMappings.Add(WidgetCategories.XfAtp, new List<WidgetTypes>()
            {
                WidgetTypes.MitoAtpProductionRate,
                WidgetTypes.GlycoAtpProductionRate,
                WidgetTypes.AtpProductionRateData,
                WidgetTypes.AtpProductionRateBasal,
                WidgetTypes.AtpProductionRateInduced,
                WidgetTypes.EnergeticMapBasal,
                WidgetTypes.EnergeticMapInduced,
                WidgetTypes.XfAtpRateIndex,
                WidgetTypes.DataTable
            });

            widgetMappings.Add(WidgetCategories.XfCellEnergy, new List<WidgetTypes>()
            {
                WidgetTypes.XfCellEnergyPhenotype,
                WidgetTypes.MetabolicPotentialOcr,
                WidgetTypes.MetabolicPotentialEcar,
                WidgetTypes.BaselineOcr,
                WidgetTypes.BaselineEcar,
                WidgetTypes.StressedOcr,
                WidgetTypes.StressedEcar,
                WidgetTypes.DataTable
            });

            return widgetMappings;
        }

        private void ClickWidgetElement(WidgetTypes wType, WidgetCategories wCat)
        {
            IWebElement widgetElement = GetWidgetElement(wType, wCat);
            if (widgetElement.Displayed)
            {
                string widgetDescription = GetWidgetDescription(wType, wCat);
                _findElements.ClickElement(widgetElement, _currentPage, widgetDescription);
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Info, $"{wCat} -Widgets are not selected");
            }
        }

        private IWebElement? GetWidgetElement(WidgetTypes wType, WidgetCategories wCat)
        {
            // Return the appropriate widget web element based on the widget type
            switch (wCat, wType)
            {
                // Quick View
                case (WidgetCategories.XfStandard, WidgetTypes.KineticGraph):
                    return KineticGraphOCR;
                case (WidgetCategories.XfStandard, WidgetTypes.KineticGraphEcar):
                    return KineticGraphECAR;
                case (WidgetCategories.XfStandard, WidgetTypes.KineticGraphPer):
                    return KineticGraphPER;
                case (WidgetCategories.XfStandard, WidgetTypes.BarChart):
                    return BarChart;
                case (WidgetCategories.XfStandard, WidgetTypes.EnergyMap):
                    return EnergyMap;
                case (WidgetCategories.XfStandard, WidgetTypes.HeatMap):
                    return HeatMap;

                // StandardDose
                case (WidgetCategories.XfStandardDose, WidgetTypes.DoseResponse):
                    return DoseResponse;

                // XF Cell Mito Stress Test View
                case (WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration):
                    return MitochondrialRespirationWidget;
                case (WidgetCategories.XfMst, WidgetTypes.Basal):
                    return BasalWidget;
                case (WidgetCategories.XfMst, WidgetTypes.AcuteResponse):
                    return AcuteResponseWidget;
                case (WidgetCategories.XfMst, WidgetTypes.ProtonLeak):
                    return ProtonLeakWidget;
                case (WidgetCategories.XfMst, WidgetTypes.MaximalRespiration):
                    return MaximalRespirationWidget;
                case (WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity):
                    return SpareRespiratoryCapacityWidget;
                case (WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption):
                    return NonMitoO2ConsumptionWidget;
                case (WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration):
                    return AtpProductionCoupledRespirationWidget;
                case (WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent):
                    return CouplingEfficiencyPercentwidget;
                case (WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent):
                    return SpareRespiratoryCapacityPercentWidget;
                case (WidgetCategories.XfMst, WidgetTypes.DataTable):
                    return DataTableWidget;

                // XF ATP Rate Assay View
                case (WidgetCategories.XfAtp, WidgetTypes.MitoAtpProductionRate):
                    return MitoAtpProductionRatewidget;
                case (WidgetCategories.XfAtp, WidgetTypes.GlycoAtpProductionRate):
                    return GlycoAtpProductionRatewidget;
                case (WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateData):
                    return AtpProductionRateDataWidget;
                case (WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateBasal):
                    return AtpProductionRateBasalWidget;
                case (WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateInduced):
                    return AtpProductionRateInducedWidget;
                case (WidgetCategories.XfAtp, WidgetTypes.EnergeticMapBasal):
                    return EnergeticMapBasalWidget;
                case (WidgetCategories.XfAtp, WidgetTypes.EnergeticMapInduced):
                    return EnergeticMapInducedWidget;
                case (WidgetCategories.XfAtp, WidgetTypes.XfAtpRateIndex):
                    return XfAtpRateIndexWidget;
                case (WidgetCategories.XfAtp, WidgetTypes.DataTable):
                    return DataTableBasalandInduced;

                // XfCellEnergyPhenotyp view
                case (WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype):
                    return CellEnergyPhenotypeWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr):
                    return MetabolicPotentialOCRWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar):
                    return MetabolicPotentialECARWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr):
                    return BaselineOCRWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar):
                    return BaselineECARWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr):
                    return StressedOCRWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar):
                    return StressedECARWidget;
                case (WidgetCategories.XfCellEnergy, WidgetTypes.DataTable):
                    return CellEnergyDataTableWidget;
                default:
                    return null;
            }
        }

        private string GetWidgetDescription(WidgetTypes widget, WidgetCategories wCat)
        {
            return wCat + " - " + widget + " widget";
        }

        private void DropdownSelect(String oligo, IWebElement Dropdown, String Description)
        {
            try
            {
                _findElements?.VerifyElement(Dropdown, _currentPage, Description);

                _findElements.SelectFromDropdown(Dropdown, _currentPage, "text", oligo, $"Oligo dropdown - {oligo}");

                ScreenShot.ScreenshotNow(_driver, _currentPage, Description, ScreenshotType.Info, Dropdown);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error Occured while selecting a {Description} with the Message {e.Message}");
            }
        }

        private void VerifyAddViewPopupElements()
        {
            try
            {
                _findElements?.VerifyElement(AddviewGroups, _currentPage, $"Add View popup Groups Area");

                if(ListHeadName.Text == "Groups")
                {
                    _findElements.ClickElementByJavaScript(GroupListElement, _currentPage, $"Unselecting the First Group");

                    _findElements.ClickElementByJavaScript(GroupListElement, _currentPage, $"Selecting the First Group");
                }
                else
                {
                    _findElements.ClickElementByJavaScript(CompundListElement, _currentPage, $"Unselecting the First Group");

                    _findElements.ClickElementByJavaScript(CompundListElement, _currentPage, $"Selecting the First Group");
                }

                IReadOnlyCollection<IWebElement> selectAllWidgets = _driver.FindElements(By.CssSelector(".allwidgets"));

                foreach (IWebElement selectAllIcon in selectAllWidgets)
                {
                    if (selectAllIcon.Displayed)
                    {
                        _findElements.VerifyElement(selectAllIcon, _currentPage, "Select all Widgets Icom");
                    }
                }

                IReadOnlyCollection<IWebElement> AddViewDetails = _driver.FindElements(By.CssSelector(".col-md-12.graph-details-sec"));

                foreach (IWebElement Details in AddViewDetails)
                {
                    if (Details.Displayed)
                    {
                        _findElements.VerifyElement(Details, _currentPage, $"Add View Popup - Details section");
                    }
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Error occured while verify the add view popup elemnts. The error is {e.Message}");
            }
        }

        private void DropdownVerification()
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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Error occured while verify the oligo and induced dropdown in the add view popup. The error is {e.Message}");
            }
        }
    }
}
