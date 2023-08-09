using System;
using System.Data;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Interactions;

namespace SHAProject.Create_Widgets
{
    public class AnalysisPage
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public AnalysisPage(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc)
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _commonFunc = commonFunc;
            PageFactory.InitElements(_driver, this);
        }

        #region Header Icons
        [FindsBy(How = How.Id, Using = "breadcrumb_assayFile")]
        public IWebElement? breadcrumFile;

        [FindsBy(How = How.Id, Using = "breadcrumb_view")]
        public IWebElement? breadcrumview;

        [FindsBy(How = How.CssSelector, Using = ".data-qc-unavailable [alt='DataQc']")]
        public IWebElement? UnavailableDataQC;

        [FindsBy(How = How.CssSelector, Using = ".data-qc-warning [alt='DataQc']")]
        public IWebElement? WarningDataQC;

        [FindsBy(How = How.CssSelector, Using = "[src=\"/images/svg/Normalize Edit.svg\"]")]
        public IWebElement? NormalizeIcon;

        [FindsBy(How = How.CssSelector, Using = "[src=\"/images/svg/Modify.svg\"]")]
        public IWebElement? ModifyAssayIcon;

        [FindsBy(How = How.Id, Using = "icongraph")]
        public IWebElement? AddWidgetIcon;

        [FindsBy(How = How.Id, Using = "exportview")]
        public IWebElement? ExportView;

        [FindsBy(How = How.Id, Using = "ExportExcel")]
        public IWebElement? ExportExcel;

        [FindsBy(How = How.Id, Using = "ExportPrism")]
        public IWebElement? ExportPrism;

        [FindsBy(How = How.Id, Using = "edit-grids")]
        public IWebElement? EditLayoutBtn;

        [FindsBy(How = How.CssSelector, Using = ".exit-mode")]
        public IWebElement? ExitEditLayout;

        [FindsBy(How = How.CssSelector, Using = ".sidebar-views-sm a")]
        public IWebElement? SideBarOpenBtn;

        [FindsBy(How = How.CssSelector, Using = ".panel-heading.collapsed1 a")]
        public IWebElement? SideBarClosedBtn;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"confirmationdelete\"]/div[1]/div[1]")]
        public IWebElement? deleteWidgetPopup;

        [FindsBy(How = How.XPath, Using = "//button[@onclick=\"DeleteWidget()\"]")]
        public IWebElement? deleteWidgetYesButton;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"graphlst\"]")]
        public IWebElement? GraphListArea;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"sidebar-wrapper\"]")]
        public IWebElement? ViewListArea;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"addNewGraph\"]")]
        public IWebElement? AddNewGraph;

        #endregion

        #region StandardView Widget Elements

        [FindsBy(How = How.XPath, Using = "(//div[@class='list-options'])[last()]")]
        public IWebElement? LastOption;

        [FindsBy(How = How.XPath, Using = "//a[@id='menu-toggle-views']")]
        public IWebElement? SideViewMenuToggleBtn;

        [FindsBy(How = How.XPath, Using = "((//div[@class='popup-options'])[last()]/ul/li)[last()]")]
        public IWebElement? CustomViewOption;

        [FindsBy(How = How.Id, Using = "newmasterviewname")]
        public IWebElement? CustomNameTxtBox;

        [FindsBy(How = How.Id, Using = "txtdescription")]
        public IWebElement? CustomDescription;

        [FindsBy(How = How.Id, Using = "btnsaveasmasterview")]
        public IWebElement? SaveBtn;

        [FindsBy(How = How.ClassName, Using = "addnewlist")]
        private IWebElement? AddNewListViewIcon;

        [FindsBy(How = How.XPath, Using = "(//li[@class='pannel-li'])[last()]")]
        private IWebElement? LastCretedView;

        #endregion

         #region Drag and Drop

        [FindsBy(How = How.XPath, Using = "(//div[@id='graphlst']/div/div)[1]")]
        private IWebElement? DragStart;

        [FindsBy(How = How.XPath, Using = "(//div[@id='graphlst']/div/div)[2]")]
        private IWebElement? DragEnd;

        [FindsBy(How = How.Id, Using = "divwidget1_legend")]
        private IWebElement? LegendsArea;

        [FindsBy(How = How.XPath, Using = "(//div[@id='divwidget1_legend']/span)[1]")]
        private IWebElement? GroupLegend;

        #endregion

        public void AnalysisPageHeaderIcons()
        {
            try
            {
                Thread.Sleep(3000);

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    _commonFunc.HandleCurrentWindow();

                _findElements.VerifyElement(breadcrumFile, _currentPage, $"Analysis Page - Breadcrum File");

                _findElements.VerifyElement(breadcrumview, _currentPage, $"Analysis Page - Breadcrum View");

                if (_fileUploadOrExistingFileData.FileExtension == "asyr")
                    _findElements.VerifyElement(UnavailableDataQC, _currentPage, $"Analysis Page - Unavailable DataQc");
                else
                    _findElements.VerifyElement(WarningDataQC, _currentPage, $"Analysis Page - Warning DataQc");

                _findElements.VerifyElement(NormalizeIcon, _currentPage, $"Analysis Page - Normalization Icon");

                _findElements.VerifyElement(ModifyAssayIcon, _currentPage, $"Analysis Page - Modify Assay Icon");

                _findElements.ClickElement(ExportView, _currentPage, $"Analysis Page - Export View Icon");
                _findElements.ActionsClass(ExportExcel);
                _findElements.VerifyElement(ExportExcel, _currentPage, $"Analysis Page - Export View - Excel");

                _findElements.ClickElement(ExportView, _currentPage, $"Analysis Page - Export View Icon");
                _findElements.ActionsClass(ExportPrism);
                _findElements.VerifyElement(ExportPrism, _currentPage, $"Analysis Page - Export View - Prism");

                _findElements.VerifyElement(EditLayoutBtn, _currentPage, $"Analysis Page - Edit Layout Icon");

                _findElements.VerifyElement(AddWidgetIcon, _currentPage, $"Analysis Page - Add Widget Icon");

                _findElements.VerifyElement(GraphListArea, _currentPage, "Analysis Page - Graph List Area");

                _findElements.ClickElement(SideBarOpenBtn, _currentPage, $"Analysis Page - Sidebar Open Button");

                _findElements.VerifyElement(ViewListArea, _currentPage, "Analysis Page - View List Area");

                _findElements.ClickElement(SideBarClosedBtn, _currentPage, $"Analysis Page - Sidebar Close Button");

                _findElements.ScrollIntoView(AddNewGraph);
                _findElements.VerifyElement(AddNewGraph, _currentPage, "Analysis Page - Add New Graph");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in Analysis page header icon functionality. The error is {e.Message}");
            }
        }

        public void AnalysisPageWidgetElements(WidgetCategories wCat, List<WidgetTypes> SelectedWidgets)
        {
            try
            {
                int count = 0;
                var verifywidgets = _fileUploadOrExistingFileData.SelectedWidgets;
                _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                IReadOnlyCollection<IWebElement> gridStackItems = _driver.FindElements(By.CssSelector(".grid-stack-item"));

                foreach (IWebElement gridStackItem in gridStackItems)
                {
                    _findElements.ScrollIntoView(gridStackItem);

                    // verify the widget title
                    string widgetName = _commonFunc.GetChartTitle(WidgetCategories.XfStandard, verifywidgets[count]);
                    IWebElement widgetTitle = gridStackItem.FindElement(By.CssSelector(".blocklefthead"));
                    _findElements.ElementTextVerify(widgetTitle, widgetName, _currentPage, $"Element name - {widgetName}");

                    if (gridStackItem.FindElements(By.CssSelector(".errorMsg")).Count == 0)
                    {

                        if (gridStackItem.FindElements(By.CssSelector(".heatmap_widget")).Count ==0)
                        {
                            IWebElement panIcon = null;
                            if (gridStackItem.FindElements(By.CssSelector(".Zoom")).Count > 0)
                                panIcon = gridStackItem.FindElement(By.CssSelector(".Zoom")); // canvaschart panIcon icon
                            else if (gridStackItem.FindElements(By.CssSelector(".zoom-btn")).Count > 0)
                                panIcon = gridStackItem.FindElement(By.CssSelector(".zoom-btn")); // amchart panIcon icon

                            ExtentReport.ExtentTest("ExtentTestNode", panIcon.Displayed ? Status.Pass : Status.Fail, panIcon.Displayed ? $"Zoom icon is displayed for {widgetName} " : $"Zoom icon is not displayed {widgetName}");

                            IWebElement resetIcon = null;
                            if (gridStackItem.FindElements(By.CssSelector(".Reset")).Count > 0)
                                resetIcon = gridStackItem.FindElement(By.CssSelector(".Reset")); // canvaschart resetIcon icon
                            else if (gridStackItem.FindElements(By.CssSelector(".reset-btn")).Count > 0)
                                resetIcon = gridStackItem.FindElement(By.CssSelector(".reset-btn")); // amchart resetIcon icon

                            ExtentReport.ExtentTest("ExtentTestNode", resetIcon.Displayed ? Status.Pass : Status.Fail, resetIcon.Displayed ? $"Reset icon is displayed for {widgetName}" : $"Reset icon is not displayed for {widgetName}");
                        }

                        // verify the edit icon
                        IWebElement editIcon = gridStackItem.FindElement(By.CssSelector(".cell-edit-icons"));
                        _findElements.VerifyElement(editIcon, _currentPage, $"{verifywidgets[count]} - Edit Icon");

                        // verify the export icon
                        IWebElement exportIcon = null;
                        if (gridStackItem.FindElements(By.CssSelector(".Export")).Count > 0)
                            exportIcon = gridStackItem.FindElement(By.CssSelector(".Export")); // canvaschart export icon
                        else if (gridStackItem.FindElements(By.CssSelector(".amcharts-amexport-item-level-0")).Count > 0)
                            exportIcon = gridStackItem.FindElement(By.CssSelector(".amcharts-amexport-item-level-0")); // amchart export icon
                        else if (gridStackItem.FindElements(By.CssSelector(".cell-export-heatscrnmap")).Count > 0)
                            exportIcon = gridStackItem.FindElement(By.CssSelector(".cell-export-heatscrnmap")); // heatmap export icon
                        else if (gridStackItem.FindElements(By.CssSelector(".export_table_icon")).Count > 0)
                            exportIcon = gridStackItem.FindElement(By.CssSelector(".export_table_icon")); // datatable export icon

                        _findElements.VerifyElement(exportIcon, _currentPage, $"{verifywidgets[count]} - Export Icon");

                        // verify the measurement element (if present)
                        if (gridStackItem.FindElements(By.CssSelector(".measurement-view")).Count > 0)
                        {
                            IWebElement measurement = gridStackItem.FindElement(By.CssSelector(".measurement-view"));
                            _findElements.ElementTextVerify(measurement, "Measurement 1", _currentPage, $"{widgetName}- measurement text");
                        }

                        if (gridStackItem.FindElements(By.CssSelector(".heatmap_widget")).Count ==0)
                        {
                            if (gridStackItem.FindElements(By.CssSelector(".canvasjs-chart-canvas")).Count > 0)
                            {
                                IWebElement graphArea = gridStackItem.FindElement(By.CssSelector(".canvasjs-chart-canvas"));
                                ExtentReport.ExtentTest("ExtentTestNode", graphArea.Displayed ? Status.Pass : Status.Fail, graphArea.Displayed ? $"Graph area is displayed for {widgetName}" : $"Graph area is not displayed for {widgetName}");
                            }

                            // verify the group legend element (if present)
                            if (gridStackItem.FindElements(By.CssSelector(".platemap-legends")).Count > 0 && gridStackItem.FindElements(By.CssSelector(".cell-export-heatscrnmap")).Count == 0)
                            {
                                IWebElement grouplegend = gridStackItem.FindElement(By.CssSelector(".platemap-legends"));
                                _findElements.VerifyElement(grouplegend, _currentPage, $"{verifywidgets[count]} - Group Legends");
                            }
                        }
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Warning, $"{widgetName} has buffer factor value issue");
                    }
                    count++;
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in standard view widget elements verification functionality. The error is {e.Message}");
            }
        }

        public void ExportViewIconFunctionality()
        {
            try
            {
                bool downloadStatus = false;
                _findElements.ClickElement(ExportView, _currentPage, "Analysis Page - Export View Icon");
                _findElements.ActionsClass(ExportExcel);
                downloadStatus = _findElements.ClickElement(ExportExcel, _currentPage, $"Analysis Page - Export View - Excel");
                ExtentReport.ExtentTest("ExtentTestNode", downloadStatus ? Status.Pass : Status.Fail, downloadStatus ? $"Excel file is download scuccessfully" : $"Excel file is not downloaded");

                _findElements.ClickElement(ExportView, _currentPage, "Analysis Page - Export View Icon");
                _findElements.ActionsClass(ExportPrism);
                downloadStatus = _findElements.ClickElement(ExportPrism, _currentPage, $"Analysis Page - Export View - Prism");
                ExtentReport.ExtentTest("ExtentTestNode", downloadStatus ? Status.Pass : Status.Fail, downloadStatus ? $"Prism file is download scuccessfully" : $"Prism file is not downloaded");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in export view icon functionality. The error is {e.Message}");
            }
        }

        public void EditIconFunctionality(WidgetCategories wCat, WidgetTypes wType)
        {
            try
            {
                _findElements.ClickElement(EditLayoutBtn, _currentPage, $"Analysis Page - Edit Layout Icon");

                int widgetPosition = _commonFunc.GetWidgetPosition(wCat, wType);

                var deleteWidget = _driver.FindElement(By.XPath("//*[@data-widgettype='" + widgetPosition + "']/div[1]/div[2]/a/img"));
                if (deleteWidget != null)
                {
                    _findElements.ClickElementByJavaScript(deleteWidget, _currentPage, $"The Deleted widget is - {wType}");

                    _findElements.VerifyElement(deleteWidgetPopup, _currentPage, $"Analysis Page -Delete widget popup ");

                    _findElements.ClickElementByJavaScript(deleteWidgetYesButton, _currentPage, $"The Deleted widget popup - Yes button");

                    ScreenShot.ScreenshotNow(_driver, _currentPage, $"Deleted Widget is - {wType.ToString()}", ScreenshotType.Info);

                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Deleted Widget in the Analysis page is { wType.ToString()}");

                    _findElements.ClickElement(ExitEditLayout, _currentPage, $"Analysis Page - Exit Edit Layout Icon");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Delete widget  {wType} is not found");
                    ScreenShot.ScreenshotNow(_driver, _currentPage, $"Delete widget  {wType} is not found", ScreenshotType.Error);
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in edit icon functionality. The error is {e.Message}");
            }
        }

        public void CreateCustomView(WorkFlow5Data workFlow5Data)
        {
            try
            {
                _findElements?.ClickElementByJavaScript(SideViewMenuToggleBtn, _currentPage, $"Analysis Page - Side View Toggle Button");

                //_findElements?.ScrollIntoViewAndClickElementByJavaScript(LastOption, _currentPage, $"Analysis Page - Three Dot Option");

                _findElements?.ScrollIntoView(LastOption);

                _findElements.ClickElement(LastOption,_currentPage,"Last Option");

                _findElements?.ActionsClassClick(CustomViewOption, _currentPage, $"Custom view Option");

                _findElements?.VerifyElement(CustomNameTxtBox, _currentPage, $"Analysis Page - Create Custom View Popup Name Text box");

                _findElements?.SendKeys(workFlow5Data.CustomViewName,  CustomNameTxtBox, _currentPage, $"Analysis Page - Create Custom View Popup Name Text box");

                _findElements?.VerifyElement(CustomDescription, _currentPage, $"Analysis Page - Create Custom View Popup Description Text box");

                _findElements?.SendKeys(workFlow5Data.CustomViewDescription, CustomDescription, _currentPage, $"Analysis Page - Create Custom View Popup Description Text box");

                _findElements?.ClickElementByJavaScript(SaveBtn, _currentPage, $"Analysis Page - Create Custom View Popup Save Button");

                _findElements?.ClickElementByJavaScript(AddNewListViewIcon, _currentPage, $"Analysis Page - Add New View");

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, $"Error Occured while trying to create a custom view. The error is {e.Message}");
            }
        }

        public void VerifyCustomview()
        {
            try
            {
                _findElements?.ClickElement(SideViewMenuToggleBtn, _currentPage, $"Analysis Page - Side View Toggle Button");

                _findElements?.ScrollIntoView(AddNewListViewIcon);

                _findElements?.VerifyElement(LastCretedView, _currentPage, $"Analysis Page - Created Custom View");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, $"Error occured while trying to verify the create a custom view. The error is {e.Message}");
            }
        }

        public bool GoToEditWidget(WidgetCategories widgetCategory, WidgetTypes widgetType)
        {
            try
            {
                Thread.Sleep(6000);
                _commonFunc.HandleCurrentWindow();
                int widgetPosition = _commonFunc.GetWidgetPosition(widgetCategory, widgetType);
                IWebElement widgetDiv;
                 if (widgetType == WidgetTypes.KineticGraph)
                {
                    widgetDiv = _driver.FindElement(By.XPath("//*[@data-ratetype='OCR'][@data-widgettype='" + widgetPosition + "']/div[1]/div[1]/a/img"));
                }
                else if (widgetType == WidgetTypes.KineticGraphEcar)
                {
                    widgetDiv = _driver.FindElement(By.XPath("//*[@data-ratetype='ECAR'][@data-widgettype='" + widgetPosition + "']/div[1]/div[1]/a/img"));
                }
                else if (widgetType == WidgetTypes.KineticGraphPer)
                {
                    widgetDiv = _driver.FindElement(By.XPath("//*[@data-ratetype='PER'][@data-widgettype='" + widgetPosition + "']/div[1]/div[1]/a/img"));
                }
                else if (widgetType == WidgetTypes.DataTable)
                {
                    widgetDiv = _driver.FindElement(By.XPath("//*[@id='averagecalculation1']/div[1]/div[1]/a/img"));

                }
                else
                {
                    widgetDiv = _driver.FindElement(By.XPath("//*[@data-widgettype='" + widgetPosition + "']/div[1]/div[1]/a/img"));
                }

                _findElements.ScrollIntoViewAndClickElementByJavaScript(widgetDiv, _currentPage, $"Analysis Page - Edit Icon");

                return true;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, $"Unable to locate the Widget {e.Message}");
                return false;
            }
        }

        public void DragandDrop()
        {
            try
            {
                Actions actions = new Actions(_driver);
                actions.MoveToElement(LegendsArea).Perform();
                Thread.Sleep(2000);

                IWebElement Resize = _driver.FindElement(By.XPath("//div[@class='ui-resizable-handle ui-resizable-se ui-icon ui-icon-gripsmall-diagonal-se'][@style='z-index: 90; display: block;']"));
                actions.MoveToElement(Resize).Perform();
                actions.ClickAndHold(Resize).MoveByOffset(200, 100).Release().Perform();
                _findElements.VerifyElement(DragStart, _currentPage, $"Edit Layout Expanding  Widget");
                Thread.Sleep(2000);

                actions.MoveToElement(LegendsArea).Perform();
                actions.ClickAndHold(Resize).MoveByOffset(-200, -50).Release().Perform();
                _findElements.VerifyElement(DragStart, _currentPage, $"Edit Layout Resized Widget");
                Thread.Sleep(2000);

                _findElements.VerifyElement(GroupLegend, _currentPage, $"Group Legend");

                _findElements.ClickElementByJavaScript(GroupLegend, _currentPage, $"UnSelecting the Group Legend");

                Thread.Sleep(1000);
                _findElements.ClickElementByJavaScript(GroupLegend, _currentPage, $"ReSelecting the Group Legend");

                actions.ClickAndHold(DragStart).Perform();
                _findElements.VerifyElement(DragStart, _currentPage, $"Edit Layout Possition Changing Widget");
                Thread.Sleep(1000);

                actions.MoveToElement(DragEnd).Perform();
                Thread.Sleep(1000);
                actions.Release(DragStart).Perform();

                _findElements.VerifyElement(DragStart, _currentPage, $"Edit Layout Possition Changed Widget");

            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, $"Error Occured While performing Drag and Drop Functionality with the Message: "+ex.Message);
            }


        }
    }
}
