﻿using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Controls;
using SHAProject.Utilities;
using SHAProject.PageObject;
using SHAProject.Page_Object;
using SHAProject.Create_Widgets;
using SHAProject.EditPage;
using AventStack.ExtentReports;
using GraphSettings = SHAProject.EditPage.GraphSettings;
using SHAProject.SeleniumHelpers;
using System.Runtime.InteropServices.JavaScript;
using System.IO;
using OpenQA.Selenium;

namespace SHAProject.Workflows
{
    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("Standard View")]
    public class StandardView : Tests
    {
        public bool loginStatus;
        public Exports? exports;
        public PlateMap? plateMap;
        public HomePage? homePage;
        public FilesPage? filesPage;
        public LoginClass? loginClass;
        public ModifyAssay? modifyAssay;
        public AnalysisPage? analysisPage;
        public GroupLegends? groupLegends;
        public Normalization? normalization;
        public GraphSettings? graphSettings;
        public GraphProperties? graphProperties;
        public CreateWidgetFromAddWidget? addWidgets;
        public CreateWidgetFromAddView? createWidgets;
        public static List<string> testidList = new List<string>();
        public static new readonly string currentPage = "Standard View";

        private static IEnumerable<string> GetTestFixtureBrowsers()
        {
            string buildPath = string.Empty;
            string excelPath = string.Empty;

            /* Determine the correct file paths based on the operating system*/
            if (Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix)
            {
                buildPath = AppDomain.CurrentDomain.BaseDirectory.Replace("bin/Debug/net7.0/", "");
                excelPath = "ExcelTemplate/AutomatedData.xlsx";
            }
            else
            {
                buildPath = AppDomain.CurrentDomain.BaseDirectory.Replace("bin\\Debug\\net7.0\\", "");
                excelPath = "ExcelTemplate\\AutomatedData.xlsx";
            }

            /* Create a FileInfo object using the path to the Excel file*/
            FileInfo fileInfo = new FileInfo(buildPath + excelPath);

            /* Set the ExcelPackage license context to NonCommercial*/
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            /* Create an ArrayList to store the names of the browsers*/
            var browserList = new ArrayList();

            /* Create a DataTable to store the test data*/
            DataTable sheetData = new DataTable();

            /* Use a using statement to create an instance of ExcelPackage to read the Excel file*/
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                /* Read the "Login" worksheet to retrieve the names of the browsers that have been checked*/
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Login"];

                foreach (var drawings in worksheet.Drawings)
                {
                    /* Check if the checkbox has been checked*/
                    var checkbox = drawings as ExcelControlCheckBox;
                    var status = checkbox.Checked;

                    if (status.ToString() == "Checked")
                    {
                        /* Add the name of the checked browser to the browserList*/
                        browserList.Add(checkbox.Text);
                    }
                }

                /* Read the "Workflow-1" worksheet to retrieve the test data*/
                worksheet = package.Workbook.Worksheets["Workflow-5"];

                /* Read the test data from the worksheet and add it to the sheetData DataTable*/
                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    DataRow dataRow = sheetData.NewRow();

                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;

                        /* If this is the first row, create the columns in the DataTable*/
                        if (row == 1)
                        {
                            sheetData.Columns.Add(cellValue != null ? cellValue.ToString() : "");
                        }
                        else
                        {
                            dataRow[col - 1] = cellValue;
                        }
                    }
                    /* Add the row of data to the sheetData DataTable*/
                    sheetData.Rows.Add(dataRow);
                }

                /* Extract the test IDs from the sheetData DataTable and store them in a list named testidList*/
                testidList = sheetData.AsEnumerable().Select(r => r.Field<string>("Run Name")).ToList();
            }

            /* Yield each browser name as an IEnumerable<string>*/
            foreach (var browser in browserList)
            {
                yield return browser.ToString();
            }
        }

        public StandardView(string browser)
        {
            current_browser = browser;
        }

        [OneTimeSetUp]
        public void Setup()
        {
            commonFunc.CreateDirectory(loginFolderPath, currentPage);
            string loginFoldersPath = loginFolderPath + "\\" + currentPage;
            commonFunc.CreateDirectory(loginFoldersPath, "Success");
            commonFunc.CreateDirectory(loginFoldersPath, "Error");
            commonFunc.CreateDirectory(loginFoldersPath, "Downloads");

            setup = new DriverSetup();
            driver = setup.browser(current_browser, loginData.Website, loginFoldersPath + "\\Downloads\\");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(120);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
            commonFunc.SetDriver(driver);

            loginClass = new LoginClass(driver, loginData, commonFunc);
            loginStatus = loginClass.LoginAsExcelUser();

            ExtentReport.CreateExtentTest("WorkFlow -5 : StandardView");
            bool ExcelReadStatus = reader.ReadDataFromExcel("Workflow-5");
            //Thread.sleep(2000);

            if (ExcelReadStatus)
            {
                extentTest.Log(Status.Pass, "Excel read status is true for " + currentPage);
            }
            else
            {
                extentTest.Log(Status.Fail, "Excel read status is false for " + currentPage);
                return;
            }

            ObjectInitalize();
        }

        public void ObjectInitalize()
        {
            graphSettings = new GraphSettings(currentPage, driver, loginClass.findElements, commonFunc);
            graphProperties = new GraphProperties(currentPage, driver, loginClass.findElements, commonFunc);
            homePage = new HomePage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData);
            exports = new Exports(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            filesPage = new FilesPage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, FilesTabData);
            modifyAssay = new ModifyAssay(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            analysisPage = new AnalysisPage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            groupLegends = new GroupLegends(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            addWidgets = new CreateWidgetFromAddWidget(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            createWidgets = new CreateWidgetFromAddView(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            normalization = new Normalization(currentPage, driver, loginClass.findElements, normalizationData, fileUploadOrExistingFileData, commonFunc);
            plateMap = new PlateMap(currentPage, driver, loginClass.findElements, commonFunc, fileUploadOrExistingFileData, fileUploadOrExistingFileData.FileType, normalizationData);
        }

        [Test, Order(1)]
        public void CreateQuickView()
        {
            ExtentReport.CreateExtentTestNode("CreateQuickView");

            if (loginStatus)
            {
                bool FileStatus = false;
                bool Searchedfile = false;
                if (fileUploadOrExistingFileData.IsFileUploadRequired)
                {
                    FileStatus = homePage.HomePageFileUpload();
                }
                else if (fileUploadOrExistingFileData.OpenExistingFile)
                {
                    Searchedfile = filesPage.SearchFilesInFileTab(currentPage);
                }
                else
                {
                    Assert.Ignore("Both FileUpload status and Open existing file status is false");
                }

                if (FileStatus || Searchedfile)
                {
                    createWidgets?.CreateWidgets(WidgetCategories.XfStandard, fileUploadOrExistingFileData.SelectedWidgets);
                }
            }
            else
            {
                Assert.Fail();
            }
        }

        [Test, Order(2)]
        public void CheckQuickViewLayout()
        {
            if (WorkFlow5Data.AnalysisLayoutVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                ExtentReport.CreateExtentTestNode("CreateQuickViewLayout");

                analysisPage.AnalysisPageHeaderIcons();

                analysisPage.AnalysisPageWidgetElements(WidgetCategories.XfStandard, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The Layout Verification for Standard view is given in the excel sheet is selected as No");
        }

        [Test, Order(4)]
        public void CheckAnalysisPageFunctionality()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Check Analysis Page Functionality");

            Thread.Sleep(3000);
            if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                commonFunc.HandleCurrentWindow();

            analysisPage.ExportViewIconFunctionality();

            if (WorkFlow5Data.DeleteWidgetRequired)
                analysisPage.EditIconFunctionality(WidgetCategories.XfStandard, WorkFlow5Data.DeleteWidgetName);

            if (WorkFlow5Data.AddWidgetRequired)
                addWidgets.AddWidgets(WidgetCategories.XfStandard, WorkFlow5Data.AddWidgetName);

            commonFunc.MoveBackToAnalysisPage();
        }

        [Test, Order(5)]
        public void NormalizationConcept()
        {
            if (WorkFlow5Data.NormalizationVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                ExtentReport.CreateExtentTestNode("Normalization Concept");

                Thread.Sleep(3000);
                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                normalization.ApplyNormalizationValues(WorkFlow5Data.ApplyToAllWidgets);
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Verification for normalization concept is given in the excel sheet is selected as No.");
        }

        [Test, Order(6)]
        public void ModifyAssay()
        {
            if (WorkFlow5Data.ModifyAssay)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                ExtentReport.CreateExtentTestNode("Modify Assay");

                modifyAssay.ModifyAssayHeaderTabs();

                modifyAssay.GroupTabElements(WorkFlow5Data.AddGroupName);

                modifyAssay.PlateMapElements(WorkFlow5Data.SelecttheControls);

                modifyAssay.AssayMediaElements();

                modifyAssay.BackgroundBufferElements();

                modifyAssay.InjectionNamesElements(WorkFlow5Data.InjectionName);
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The verification for modify assay is given in the excel sheet is selected as No.");
        }

        [Test, Order(7)]

        public void CheckEditWidgetPage()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Check Edit Widget Page Functionality");

            switch (WorkFlow5Data.SelectWidgetName)
            {
                case WidgetTypes.KineticGraph:
                    ExtentReport.CreateExtentTestNode("Kinetic Graph OCR");
                    KineticGraphs(WorkFlow5Data.KineticGraphOcr, WidgetTypes.KineticGraph);
                    break;
                case WidgetTypes.KineticGraphEcar:
                    ExtentReport.CreateExtentTestNode("Kinetic Graph ECAR");
                    KineticGraphs(WorkFlow5Data.KineticGraphEcar, WidgetTypes.KineticGraphEcar);
                    break;
                case WidgetTypes.KineticGraphPer:
                    ExtentReport.CreateExtentTestNode("Kinetic Graph PER");
                    KineticGraphs(WorkFlow5Data.KineticGraphPer, WidgetTypes.KineticGraphPer);
                    break;
                case WidgetTypes.BarChart:
                    BarChart();
                    break;
                case WidgetTypes.EnergyMap:
                    EnergyMap();
                    break;
                case WidgetTypes.HeatMap:
                    HeatMap();
                    break;
                default:
                    ExtentReport.CreateExtentTestNode("Kinetic Graph OCR");
                    KineticGraphs(WorkFlow5Data.KineticGraphOcr, WidgetTypes.KineticGraph);
                    break;
            }
        }

        [Test, Order(8)]
        public void KineticGraph()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            foreach (WidgetTypes widget in fileUploadOrExistingFileData.SelectedWidgets)
            {
                if (widget == WidgetTypes.KineticGraph)
                {
                    ExtentReport.CreateExtentTestNode("Kinetic Graph OCR");
                    KineticGraphs(WorkFlow5Data.KineticGraphOcr, WidgetTypes.KineticGraph);
                }

                if (widget == WidgetTypes.KineticGraphEcar)
                {
                    ExtentReport.CreateExtentTestNode("Kinetic Graph ECAR");
                    KineticGraphs(WorkFlow5Data.KineticGraphEcar, WidgetTypes.KineticGraphEcar);
                }

                if (widget == WidgetTypes.KineticGraphPer)
                {
                    ExtentReport.CreateExtentTestNode("Kinetic Graph PER");
                    KineticGraphs(WorkFlow5Data.KineticGraphPer, WidgetTypes.KineticGraphPer);
                }
            }
        }

        [Test, Order(9)]
        public void BarChart()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.BarChart))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                Thread.Sleep(5000);

                ExtentReport.CreateExtentTestNode("Bar Chart");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.BarChart);
                if (hasEditWidgetPageGone)
                {
                    graphProperties.Measurement(WorkFlow5Data.Barchart);

                    graphProperties.Rate(WorkFlow5Data.Barchart);

                    graphProperties.Display(WorkFlow5Data.Barchart);

                    graphProperties.Normalization(WorkFlow5Data.Barchart);

                    graphProperties.ErrorFormat(WorkFlow5Data.Barchart, WidgetCategories.XfStandard, WidgetTypes.BarChart);

                    graphProperties.BackgroundCorrection(WorkFlow5Data.Barchart);

                    graphProperties.Baseline(WorkFlow5Data.Barchart);

                    graphProperties.SortBy(WorkFlow5Data.Barchart);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.VerifyExpectedGraphUnits(WorkFlow5Data.Barchart.ExpectedGraphUnits, WidgetTypes.BarChart, false);

                    if (WorkFlow5Data.Barchart.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.YAutoScale(WorkFlow5Data.Barchart);

                        graphSettings.ZeroLine(WorkFlow5Data.Barchart);

                        graphSettings.Zoom(WorkFlow5Data.Barchart);

                        graphSettings.GraphSettingsApply();
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow5Data.Barchart.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfStandard, WidgetTypes.BarChart, WorkFlow5Data.Barchart);

                    if (WorkFlow5Data.Barchart.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfStandard, WidgetTypes.BarChart, WorkFlow5Data.Barchart);
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Bar Chart widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(10)]
        public void EnergyMap()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.EnergyMap))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                ExtentReport.CreateExtentTestNode("Energy Map");

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.EnergyMap);
                if (hasEditWidgetPageGone)
                {
                    graphProperties.Measurement(WorkFlow5Data.EnergyMap);

                    graphProperties.Rate(WorkFlow5Data.EnergyMap);

                    graphProperties.Display(WorkFlow5Data.EnergyMap);

                    graphProperties.Normalization(WorkFlow5Data.EnergyMap);

                    graphProperties.ErrorFormat(WorkFlow5Data.EnergyMap, WidgetCategories.XfStandard, WidgetTypes.EnergyMap);

                    graphProperties.BackgroundCorrection(WorkFlow5Data.EnergyMap);

                    graphProperties.Baseline(WorkFlow5Data.EnergyMap);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.VerifyExpectedGraphUnits(WorkFlow5Data.EnergyMap.ExpectedGraphUnits, WidgetTypes.EnergyMap, false);

                    if (WorkFlow5Data.EnergyMap.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow5Data.EnergyMap);

                        graphSettings.YAutoScale(WorkFlow5Data.EnergyMap);

                        graphSettings.Zoom(WorkFlow5Data.EnergyMap);

                        graphSettings.GraphSettingsApply();
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow5Data.EnergyMap.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfStandard, WidgetTypes.EnergyMap, WorkFlow5Data.EnergyMap);

                    if (WorkFlow5Data.EnergyMap.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfStandard, WidgetTypes.EnergyMap, WorkFlow5Data.EnergyMap);
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Energy Map widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(11)]
        public void HeatMap()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.HeatMap))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                ExtentReport.CreateExtentTestNode("Heat Map");

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.HeatMap);
                if (hasEditWidgetPageGone)
                {
                    graphProperties.Measurement(WorkFlow5Data.HeatMap);

                    graphProperties.Rate(WorkFlow5Data.HeatMap);

                    graphProperties.Normalization(WorkFlow5Data.HeatMap);

                    graphProperties.BackgroundCorrection(WorkFlow5Data.HeatMap);

                    graphProperties.Baseline(WorkFlow5Data.HeatMap);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.VerifyExpectedGraphUnits(WorkFlow5Data.HeatMap.ExpectedGraphUnits, WidgetTypes.HeatMap, false);

                    if (WorkFlow5Data.HeatMap.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.YAutoScale(WorkFlow5Data.HeatMap);

                        graphSettings.ZeroLine(WorkFlow5Data.HeatMap);

                        graphSettings.LineMarkers(WorkFlow5Data.HeatMap);

                        graphSettings.RateHighlight(WorkFlow5Data.HeatMap);

                        graphSettings.InjectionMarkers(WorkFlow5Data.HeatMap);

                        graphSettings.Zoom(WorkFlow5Data.HeatMap);

                        graphSettings.GraphSettingsApply();
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow5Data.HeatMap.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfStandard, WidgetTypes.HeatMap, WorkFlow5Data.HeatMap);

                    if (WorkFlow5Data.HeatMap.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfStandard, WidgetTypes.HeatMap, WorkFlow5Data.HeatMap);
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Heat Map widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(12)]
        public void DoseResponseAddWidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            //ExtentReport.CreateExtentTestNode("Dosr Response in Add Widget");

            Thread.Sleep(10000);

            if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                commonFunc.HandleCurrentWindow();

            addWidgets.AddWidgets(WidgetCategories.XfStandard, WidgetTypes.DoseResponse);

        }

        [Test, Order(13)]
        public void DoseResponseAddView()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Dose Response in Add View");

            createWidgets?.CreateWidgets(WidgetCategories.XfStandardDose, WorkFlow5Data.AddDoseWidget);
        }

        [Test, Order(14)]
        public void Dose_Response()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                DoseResponseAddView();

            ExtentReport.CreateExtentTestNode("Dose Response Widget");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandardDose, WidgetTypes.DoseResponse);
            if (hasEditWidgetPageGone)
            {
                graphProperties.Measurement(WorkFlow5Data.DoseResponse);

                graphProperties.Rate(WorkFlow5Data.DoseResponse);

                graphProperties.Normalization(WorkFlow5Data.DoseResponse);

                graphProperties.ErrorFormat(WorkFlow5Data.DoseResponse, WidgetCategories.XfStandardDose, WidgetTypes.DoseResponse);

                graphProperties.BackgroundCorrection(WorkFlow5Data.DoseResponse);

                if (WorkFlow5Data.DoseResponse.CheckNormalizationWithPlateMap)
                    graphProperties.VerifyExpectedGraphUnits(WorkFlow5Data.DoseResponse.ExpectedGraphUnits, WidgetTypes.DoseResponse, false);

                if (WorkFlow5Data.DoseResponse.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow5Data.DoseResponse);

                    graphSettings.ZeroLine(WorkFlow5Data.DoseResponse);

                    graphSettings.LineMarkers(WorkFlow5Data.DoseResponse);

                    graphSettings.RateHighlight(WorkFlow5Data.DoseResponse);

                    graphSettings.InjectionMarkers(WorkFlow5Data.DoseResponse);

                    graphSettings.Zoom(WorkFlow5Data.DoseResponse);

                    graphSettings.GraphSettingsApply();

                    graphSettings.VerifyDoseGraphSettingsIcon();

                    // Dose graph settings
                    graphSettings.DoseYAutoScale(WorkFlow5Data.DoseResponse);

                    graphSettings.DoseXAutoScale(WorkFlow5Data.DoseResponse);

                    graphSettings.DoseZeroLine(WorkFlow5Data.DoseResponse);

                    graphSettings.DoseLineMarkers(WorkFlow5Data.DoseResponse);

                    graphSettings.DoseZoom(WorkFlow5Data.DoseResponse);

                    graphSettings.DoseGraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow5Data.DoseResponse.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("A05", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfStandard, WidgetTypes.DoseResponse, WorkFlow5Data.DoseResponse);

                if (WorkFlow5Data.DoseResponse.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfStandard, WidgetTypes.DoseResponse, WorkFlow5Data.DoseResponse);
            }
        }

        [Test, Order(15)]
        public void BlankView()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                homePage.HomePageFileUpload();

            ExtentReport.CreateExtentTestNode("Create Blank View");

            List<WidgetTypes> widgets = new List<WidgetTypes>();

            createWidgets.CreateWidgets(WidgetCategories.XfStandardBlank, widgets);

            addWidgets.AddWidgets(WidgetCategories.XfStandard, WidgetTypes.DoseResponse);
        }

        [Test, Order(16)]
        public void CustomView()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Create Custom View");

            analysisPage.CreateCustomView(WorkFlow5Data);

            List<WidgetTypes> widgets = new List<WidgetTypes>();

            createWidgets.AddView(WidgetCategories.XfCustomview, widgets);

            analysisPage.VerifyCustomview();

        }

        public void KineticGraphs(WidgetItems widget, WidgetTypes wType)
        {
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, wType);
            if (hasEditWidgetPageGone)
            {
                graphProperties.Measurement(widget);

                graphProperties.Rate(widget);

                graphProperties.Display(widget);

                graphProperties.Y(widget);

                graphProperties.Normalization(widget);

                graphProperties.ErrorFormat(widget, WidgetCategories.XfStandard, wType);

                graphProperties.BackgroundCorrection(widget);

                graphProperties.Baseline(widget);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.VerifyExpectedGraphUnits(widget.ExpectedGraphUnits, wType, false);

                if (widget.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(widget);

                    graphSettings.ZeroLine(widget);

                    graphSettings.LineMarkers(widget);

                    graphSettings.RateHighlight(widget);

                    graphSettings.InjectionMarkers(widget);

                    graphSettings.Zoom(widget);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (widget.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("A05", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfStandard, wType, widget);

                if (widget.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfStandard, wType, widget);
            }
        }
    }
}
