using System;
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
using System.Security.Cryptography.X509Certificates;

namespace SHAProject.Workflows
{
    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("Standard View")]
    public class StandardView : Tests
    {
        public bool loginStatus;
        public Exports exports;
        public PlateMap plateMap;
        public LoginClass loginClass;
        public UploadFile uploadFile;
        public ModifyAssay modifyAssay;
        public AnalysisPage analysisPage;
        public GroupLegends groupLegends;
        public Normalization normalization;
        public GraphSettings graphSettings;
        public GraphProperties? graphProperties;
        public CreateWidgetFromAddWidget addWidgets;
        public CreateWidgetFromAddView createWidgets;
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
                    var status = checkbox?.Checked;

                    if (status.ToString() == "Checked")
                    {
                        /* Add the name of the checked browser to the browserList*/
                        browserList.Add(checkbox.Text);
                    }
                }

                /* Read the "Workflow-1" worksheet to retrieve the test data*/
                worksheet = package.Workbook.Worksheets["Workflow-1"];

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

            ExtentReport.CreateExtentTest("WorkFlow -1 : StandardView");
            bool ExcelReadStatus = reader.ReadDataFromExcel("Workflow-1");
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
            uploadFile = new UploadFile(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData);
            exports = new Exports(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            modifyAssay = new ModifyAssay(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            analysisPage = new AnalysisPage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            groupLegends = new GroupLegends(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            addWidgets = new CreateWidgetFromAddWidget(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            createWidgets = new CreateWidgetFromAddView(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            normalization = new Normalization(currentPage, driver, loginClass.findElements, normalizationData, fileUploadOrExistingFileData, commonFunc);
            plateMap = new PlateMap(currentPage, driver, loginClass.findElements, commonFunc, fileUploadOrExistingFileData, fileUploadOrExistingFileData.FileType);
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
                    FileStatus = uploadFile.HomePageFileUpload();
                }
                else if (fileUploadOrExistingFileData.OpenExistingFile)
                {
                    Searchedfile = uploadFile.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);
                }
                else
                {
                    Assert.Ignore("Both FileUpload status and Open existing file status is false");
                }

                if (!FileStatus && Searchedfile)
                {
                    Thread.Sleep(5000);
                    createWidgets?.CreateWidgets(WidgetCategories.XfStandard, fileUploadOrExistingFileData.SelectedWidgets);
                }
                else
                {
                    Thread.Sleep(3000);
                    createWidgets?.AddView(WidgetCategories.XfStandard, fileUploadOrExistingFileData.SelectedWidgets);
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

            if (WorkFlow1Data.AnalysisLayoutVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                ExtentReport.CreateExtentTestNode("CreateQuickViewLayout");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

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

            if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                commonFunc.HandleCurrentWindow();

            ExtentReport.CreateExtentTestNode("Check Analysis Page Functionality");

            analysisPage.ExportViewIconFunctionality();

            if (WorkFlow1Data.DeleteWidgetRequired)
                analysisPage.EditIconFunctionality(WidgetCategories.XfStandard, WorkFlow1Data.DeleteWidgetName);

            if (WorkFlow1Data.AddWidgetRequired)
                addWidgets.AddWidgets(WidgetCategories.XfStandard, WorkFlow1Data.AddWidgetName);

            commonFunc.MoveBackToAnalysisPage();
        }

        [Test, Order(5)]
        public void NormalizationConcept()
        {

            if (WorkFlow1Data.NormalizationVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateQuickView();

                ExtentReport.CreateExtentTestNode("Normalization Concept");

                normalization.ApplyNormalizationValues(WorkFlow1Data.ApplyToAllWidgets);

                analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.BarChart);

                normalization.NormalizationToggle();

                commonFunc.MoveBackToAnalysisPage();

                uploadFile.SearchFilesInFileTab(WorkFlow1Data.NormalizedFileName);

                normalization.NormalizationElements();

            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Verification for normalization concept is given in the excel sheet is selected as No.");
        }

        [Test, Order(6)]
        public void ModifyAssay()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Modify Assay");

            if (WorkFlow1Data.ModifyAssay)
            {
                modifyAssay.ModifyAssayHeaderTabs();

                modifyAssay.GroupTabElements(WorkFlow1Data.AddGroupName);

                modifyAssay.PlateMapElements(WorkFlow1Data.SelecttheControls);

                modifyAssay.AssayMediaElements();

                modifyAssay.BackgroundBufferElements();

                modifyAssay.InjectionNamesElements(WorkFlow1Data.InjectionName);

            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Verification for modify assy is given in the excel sheet is selected as No.");
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

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WorkFlow1Data.SelectWidgetName);
            if (hasEditWidgetPageGone)
            {
                graphProperties.Graphproprties();

                graphSettings.VerifyGraphSettings();

                WidgetItems widget = null;
                switch (WorkFlow1Data.SelectWidgetName)
                {
                    case WidgetTypes.KineticGraph:
                        widget = WorkFlow1Data.KineticGraphOcr;
                        break;
                    case WidgetTypes.KineticGraphEcar:
                        widget = WorkFlow1Data.KineticGraphEcar;
                        break;
                    case WidgetTypes.KineticGraphPer:
                        widget = WorkFlow1Data.KineticGraphPer;
                        break;
                    case WidgetTypes.BarChart:
                        widget = WorkFlow1Data.Barchart;
                        break;
                    case WidgetTypes.EnergyMap:
                        widget = WorkFlow1Data.EnergyMap;
                        break;
                    case WidgetTypes.HeatMap:
                        widget = WorkFlow1Data.HeatMap;
                        break;
                    default:
                        widget = WorkFlow1Data.KineticGraphOcr;
                        break;
                }
                graphSettings.GraphSettingsField(widget);

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionality();

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfStandard, WidgetTypes.BarChart, WorkFlow1Data.Barchart);

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
                    ExtentReport.CreateExtentTestNode("Kinetic Graph Ocr");
                    KineticGraphs(WorkFlow1Data.KineticGraphOcr,widget);
                }
                if(widget == WidgetTypes.KineticGraphEcar) 
                {
                    ExtentReport.CreateExtentTestNode("Kinetic Graph Ecar");
                    KineticGraphs(WorkFlow1Data.KineticGraphEcar, widget);
                }
                if(widget == WidgetTypes.KineticGraphPer)
                {
                    ExtentReport.CreateExtentTestNode("Kinetic Graph Per");
                    KineticGraphs(WorkFlow1Data.KineticGraphPer, widget);
                }
            }
        }

        [Test, Order(9)]
        public void BarChart()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Bar Chart");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.BarChart);
            if (hasEditWidgetPageGone)
            {
                graphProperties.Measurement(WorkFlow1Data.Barchart);

                graphProperties.Rate(WorkFlow1Data.Barchart);

                graphProperties.Display(WorkFlow1Data.Barchart);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow1Data.Barchart);

                graphProperties.ErrorFormat(WorkFlow1Data.Barchart);

                graphProperties.BackgroundCorrection(WorkFlow1Data.Barchart);

                graphProperties.Baseline(WorkFlow1Data.Barchart);

                graphSettings.VerifyGraphSettings();

                graphSettings.GraphSettingsField(WorkFlow1Data.Barchart);

                //ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.GlycoAtpProductionRate);

                //if (platemapWellCountResult.Status)
                //    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.GlycoAtpProductionRate}");
                //else
                //    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.GlycoAtpProductionRate}");

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionality();

                plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("A05", "Included in the current calculation");

                groupLegends?.EditWidgetGroupLegends(WidgetCategories.XfStandard, WidgetTypes.BarChart, WorkFlow1Data.Barchart);

                graphProperties.VerifyNormalizationUnits(WorkFlow1Data.Barchart.GraphUnits, WidgetTypes.BarChart, false);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.VerifyNormalizationUnits(WorkFlow1Data.Barchart.GraphUnits, WidgetTypes.BarChart, true);

                if (WorkFlow1Data.Barchart.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfStandard, WidgetTypes.BarChart, WorkFlow1Data.Barchart);
            }
        }

        [Test, Order(10)]
        public void EnergyMap()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Energy Map");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.BarChart);
            if (hasEditWidgetPageGone)
            {
            }
        }

        [Test, Order(11)]
        public void HeatMap()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Heat Map");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, WidgetTypes.BarChart);
            if (hasEditWidgetPageGone)
            {
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

            createWidgets?.CreateWidgets(WidgetCategories.XfStandardDose, WorkFlow1Data.AddDoseWidget);
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
                graphProperties.Measurement(WorkFlow1Data.DoseResponse);

                graphProperties.Rate(WorkFlow1Data.DoseResponse);

                graphProperties.Normalization(WorkFlow1Data.DoseResponse);

                graphProperties.ErrorFormat(WorkFlow1Data.DoseResponse);

                graphProperties.BackgroundCorrection(WorkFlow1Data.DoseResponse);

                if(WorkFlow1Data.DoseResponse.GraphSettingsVerify)
                    graphSettings.VerifyGraphSettings();
                if (WorkFlow1Data.DoseResponse.CheckNormalizationWithPlateMap)
                    graphProperties.VerifyNormalizationUnits(WorkFlow1Data.DoseResponse.GraphUnits, WidgetTypes.DoseResponse, false);

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionality();

                if (WorkFlow1Data.DoseResponse.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfStandard, WidgetTypes.DoseResponse, WorkFlow1Data.DoseResponse);

            }
        }

        [Test,Order(15)]
        public void BlankView()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                uploadFile.HomePageFileUpload();

            //ExtentReport.CreateExtentTestNode("Create Blank View");

            List<WidgetTypes> widgets = new List<WidgetTypes>();

            createWidgets.CreateWidgets(WidgetCategories.XfStandardBlank, widgets);

            addWidgets.AddWidgets(WidgetCategories.XfStandard, WidgetTypes.DoseResponse);
        }

        [Test,Order(16)]
        public void CustomView()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateQuickView();

            ExtentReport.CreateExtentTestNode("Create Custom View");

            analysisPage.CreatecustomView(WorkFlow1Data);

            List<WidgetTypes> widgets = new List<WidgetTypes>();

            createWidgets.AddView(WidgetCategories.XfCustomview, widgets);

            analysisPage.VerifyCustomview();

        }

        public void KineticGraphs(WidgetItems Graph,WidgetTypes widget)
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfStandard, widget);
            if (hasEditWidgetPageGone)
            {
                graphProperties.Measurement(Graph);

                graphProperties.Rate(Graph);

                graphProperties.Display(Graph);

                graphProperties.Y(Graph);

                graphProperties.Normalization(Graph);

                graphProperties.ErrorFormat(Graph);

                graphProperties.BackgroundCorrection(Graph);

                graphProperties.Baseline(Graph);

                graphProperties.VerifyNormalizationUnits(Graph.GraphUnits, widget, false);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.VerifyNormalizationUnits(Graph.GraphUnits, widget, false);

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionality();

                if (Graph.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("A05", "Included in the current calculation");

                graphSettings.VerifyGraphSettings();

                if (Graph.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfStandard, widget, Graph);

            }
        }
    }
}