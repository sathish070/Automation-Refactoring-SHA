using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using AventStack.ExtentReports;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml;
using OpenQA.Selenium;
using SHAProject.Create_Widgets;
using SHAProject.Utilities;
using SHAProject.PageObject;
using SHAProject.EditPage;
using GraphSettings = SHAProject.EditPage.GraphSettings;
using SHAProject.Page_Object;
using System.IO;

namespace SHAProject.Workflows
{
    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("XF Cell Mito Stress View")]
    public class XFCellMitoStressView : Tests
    {
        public static readonly string currentPage = "XF Cell Mito Stress View";
        public bool loginStatus;
        public Exports? exports;
        public Graph? graph;
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
        public CreateWidgetFromAddView? createWidgets;

        public static List<string> testidList = new List<string>();
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
                        browserList.Add(checkbox?.Text);
                    }
                }

                /* Read the "Workflow-1" worksheet to retrieve the test data*/
                worksheet = package.Workbook.Worksheets["Workflow-6"];

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
        public XFCellMitoStressView(String browser)
        {
            current_browser = browser;
        }

        [OneTimeSetUp]
        public new void Setup()
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

            ExtentReport.CreateExtentTest("XF Cell Mito Stress View");
            bool ExcelReadStatus = reader.ReadDataFromExcel("Workflow-6");

            if (ExcelReadStatus)
            {
                extentTest?.Log(Status.Pass, "Excel read status is true for " + currentPage);
            }
            else
            {
                extentTest?.Log(Status.Fail, "Excel read status is false for " + currentPage);
                return;
            }

            ObjectInitalized();

            ExtentReport.CreateExtentTestNode("XF Cell Mito Stress View");
        }
        public void ObjectInitalized()
        {
            graphSettings = new GraphSettings(currentPage, driver, loginClass.findElements, commonFunc);
            graphProperties = new GraphProperties(currentPage, driver, loginClass.findElements, commonFunc);
            homePage = new HomePage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData);
            exports = new Exports(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            filesPage = new FilesPage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, FilesTabData);
            modifyAssay = new ModifyAssay(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            analysisPage = new AnalysisPage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            groupLegends = new GroupLegends(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            normalization = new Normalization(currentPage, driver, loginClass.findElements, normalizationData, fileUploadOrExistingFileData, commonFunc);
            plateMap = new PlateMap(currentPage, driver, loginClass.findElements, commonFunc, fileUploadOrExistingFileData, fileUploadOrExistingFileData.FileType, normalizationData);
            createWidgets = new CreateWidgetFromAddView(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            graph = new Graph(currentPage, driver, loginClass.findElements, commonFunc);
        }

        [Test, Order(1)]
        public void CreateXFCellMitoStressView()
        {
            if (loginStatus)
            {
                bool FileStatus = false;
                bool Searchedfile = false;

                if (fileUploadOrExistingFileData.IsFileUploadRequired)
                    FileStatus = homePage.HomePageFileUpload();

                else if (fileUploadOrExistingFileData.OpenExistingFile)
                    Searchedfile = filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                else
                    Assert.Ignore("Both FileUpload status and Open existing file status is false");

                if (FileStatus || Searchedfile)
                    createWidgets?.CreateWidgets(WidgetCategories.XfMst, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
                Assert.Ignore();
        }

        [Test, Order(2)]
        public void CheckXFCellMitoStressViewLayout()
        {
            ExtentReport.CreateExtentTestNode("XF Cell Mito Stress View Layout");

            if (WorkFlow6Data.AnalysisLayoutVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                analysisPage.AnalysisPageHeaderIcons();

                analysisPage.AnalysisPageWidgetElements(WidgetCategories.XfMst, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The Layout Verification for Standard view is given in the excel sheet is selected as No");
        }

        [Test, Order(3)]
        public void NormalizationConcept()
        {

            ExtentReport.CreateExtentTestNode("Normalization Concept");

            if (WorkFlow6Data.NormalizationVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                Thread.Sleep(3000);

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                normalization.ApplyNormalizationValues(WorkFlow6Data.ApplyToAllWidgets);

                analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration);

                normalization.NormalizationToggle();

                commonFunc.MoveBackToAnalysisPage();

                filesPage.SearchFilesInFileTab(WorkFlow6Data.NormalizedFileName);

                normalization.NormalizationElements();
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Verification for normalization concept is given in the excel sheet is selected as No.");
        }

        [Test, Order(4)]
        public void ModifyAssay()
        {
            ExtentReport.CreateExtentTestNode("Modify Assay");

            if (WorkFlow6Data.ModifyAssay)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                modifyAssay.ModifyAssayHeaderTabs();

                modifyAssay.GroupTabElements(WorkFlow6Data.AddGroupName);

                modifyAssay.PlateMapElements(WorkFlow6Data.SelecttheControls);

                modifyAssay.AssayMediaElements();

                modifyAssay.BackgroundBufferElements();

                modifyAssay.InjectionNamesElements(WorkFlow6Data.InjectionName);

                modifyAssay.GeneralInfoElements();
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The verification for modify assay is given in the excel sheet is selected as No.");
        }

        [Test, Order(5)]
        public void MitochondrialRespiration()
        {
            ExtentReport.CreateExtentTestNode("Mitochondrial Respiration");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.MitochondrialRespiration))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.MitochondrialRespiration;
                    widgetName = widgetType.ToString();

                    graphProperties.Measurement(WorkFlow6Data.MitochondrialRespiration);

                    graphProperties.Rate(WorkFlow6Data.MitochondrialRespiration);

                    graphProperties.Display(WorkFlow6Data.MitochondrialRespiration);

                    graphProperties.Y(WorkFlow6Data.MitochondrialRespiration);

                    graphProperties.Normalization(WorkFlow6Data.MitochondrialRespiration);

                    graphProperties.ErrorFormat(WorkFlow6Data.MitochondrialRespiration, WidgetCategories.XfMst, WidgetTypes.MitoAtpProductionRate);

                    graphProperties.BackgroundCorrection(WorkFlow6Data.MitochondrialRespiration);

                    graphProperties.Baseline(WorkFlow6Data.MitochondrialRespiration);

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.MitochondrialRespiration.ExpectedGraphUnits, WidgetTypes.MitochondrialRespiration);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.MitochondrialRespiration.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.MitochondrialRespiration);

                        graphSettings.ZeroLine(WorkFlow6Data.MitochondrialRespiration);

                        graphSettings.DataPointSymbols(WorkFlow6Data.MitochondrialRespiration);

                        graphSettings.RateHighlight(WorkFlow6Data.MitochondrialRespiration);

                        graphSettings.InjectionMarkers(WorkFlow6Data.MitochondrialRespiration);

                        graphSettings.Zoom(WorkFlow6Data.MitochondrialRespiration);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.MitochondrialRespiration.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration, WorkFlow6Data.MitochondrialRespiration);

                    if (WorkFlow6Data.MitochondrialRespiration.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration, WorkFlow6Data.MitochondrialRespiration);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Mitochondrial Respiration widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(6)]
        public void BasalRespiration()
        {
            ExtentReport.CreateExtentTestNode("Basal Respiration");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.Basal))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.Basal);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.Basal;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.BasalRespiration);

                    graphProperties.Display(WorkFlow6Data.BasalRespiration);

                    graphProperties.Normalization(WorkFlow6Data.BasalRespiration);

                    graphProperties.ErrorFormat(WorkFlow6Data.BasalRespiration, WidgetCategories.XfMst, WidgetTypes.Basal);

                    graphProperties.SortBy(WorkFlow6Data.BasalRespiration);

                    //Need to work on Graph Proiperty - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.BasalRespiration.ExpectedGraphUnits, WidgetTypes.Basal);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.BasalRespiration.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.BasalRespiration);

                        graphSettings.ZeroLine(WorkFlow6Data.BasalRespiration);

                        graphSettings.Zoom(WorkFlow6Data.BasalRespiration);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.BasalRespiration.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.Basal, WorkFlow6Data.BasalRespiration);

                    if (WorkFlow6Data.BasalRespiration.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.Basal, WorkFlow6Data.BasalRespiration);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Basal Respiration widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(7)]
        public void AcuteResponse()
        {
            ExtentReport.CreateExtentTestNode("Acute Response");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.AcuteResponse))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.AcuteResponse);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.AcuteResponse;
                    widgetName = widgetType.ToString();

                    graphProperties.Display(WorkFlow6Data.AcuteResponse);

                    graphProperties.Normalization(WorkFlow6Data.AcuteResponse);

                    graphProperties.ErrorFormat(WorkFlow6Data.AcuteResponse, WidgetCategories.XfMst, WidgetTypes.AcuteResponse);

                    graphProperties.SortBy(WorkFlow6Data.AcuteResponse);

                    //Need to work on Graph Proiperty - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.AcuteResponse.ExpectedGraphUnits, WidgetTypes.AcuteResponse);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.AcuteResponse.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.AcuteResponse);

                        graphSettings.ZeroLine(WorkFlow6Data.AcuteResponse);

                        graphSettings.Zoom(WorkFlow6Data.AcuteResponse);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.AcuteResponse.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.AcuteResponse, WorkFlow6Data.AcuteResponse);

                    if (WorkFlow6Data.AcuteResponse.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.AcuteResponse, WorkFlow6Data.AcuteResponse);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Acute Response widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(8)]
        public void ProtonLeak()
        {
            ExtentReport.CreateExtentTestNode("Proton Leak");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.ProtonLeak))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.ProtonLeak);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.ProtonLeak;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.ProtonLeak);

                    graphProperties.Display(WorkFlow6Data.ProtonLeak);

                    graphProperties.Normalization(WorkFlow6Data.ProtonLeak);

                    graphProperties.ErrorFormat(WorkFlow6Data.ProtonLeak, WidgetCategories.XfMst, WidgetTypes.ProtonLeak);

                    graphProperties.SortBy(WorkFlow6Data.ProtonLeak);

                    //Need to work on Graph Property - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.ProtonLeak.ExpectedGraphUnits, WidgetTypes.ProtonLeak);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.ProtonLeak.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.ProtonLeak);

                        graphSettings.ZeroLine(WorkFlow6Data.ProtonLeak);

                        graphSettings.Zoom(WorkFlow6Data.ProtonLeak);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.ProtonLeak.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.ProtonLeak, WorkFlow6Data.ProtonLeak);

                    if (WorkFlow6Data.ProtonLeak.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.ProtonLeak, WorkFlow6Data.ProtonLeak);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Proton Leak widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(9)]
        public void MaximalRespiration()
        {
            ExtentReport.CreateExtentTestNode("Maximal Respiration");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.MaximalRespiration))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.MaximalRespiration);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.MaximalRespiration;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.MaximalRespiration);

                    graphProperties.Display(WorkFlow6Data.MaximalRespiration);

                    graphProperties.Normalization(WorkFlow6Data.MaximalRespiration);

                    graphProperties.ErrorFormat(WorkFlow6Data.MaximalRespiration, WidgetCategories.XfMst, WidgetTypes.MaximalRespiration);

                    graphProperties.SortBy(WorkFlow6Data.MaximalRespiration);

                    //Need to work on Graph Property - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.MaximalRespiration.ExpectedGraphUnits, WidgetTypes.MaximalRespiration);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.MaximalRespiration.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.MaximalRespiration);

                        graphSettings.ZeroLine(WorkFlow6Data.MaximalRespiration);

                        graphSettings.Zoom(WorkFlow6Data.MaximalRespiration);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.MaximalRespiration.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.MaximalRespiration, WorkFlow6Data.MaximalRespiration);

                    if (WorkFlow6Data.MaximalRespiration.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.MaximalRespiration, WorkFlow6Data.MaximalRespiration);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Maximal Respiration widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(10)]
        public void SpareRespiratory()
        {
            ExtentReport.CreateExtentTestNode("Spare Respiratory Capacity");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.SpareRespiratoryCapacity))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.SpareRespiratoryCapacity;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.SpareRespiratoryCapacity);

                    graphProperties.Display(WorkFlow6Data.SpareRespiratoryCapacity);

                    graphProperties.Normalization(WorkFlow6Data.SpareRespiratoryCapacity);

                    graphProperties.ErrorFormat(WorkFlow6Data.SpareRespiratoryCapacity, WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity);

                    graphProperties.SortBy(WorkFlow6Data.SpareRespiratoryCapacity);

                    //Need to work on Graph Property - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.SpareRespiratoryCapacity.ExpectedGraphUnits, WidgetTypes.SpareRespiratoryCapacity);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.SpareRespiratoryCapacity.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.SpareRespiratoryCapacity);

                        graphSettings.ZeroLine(WorkFlow6Data.SpareRespiratoryCapacity);

                        graphSettings.Zoom(WorkFlow6Data.SpareRespiratoryCapacity);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.SpareRespiratoryCapacity.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity, WorkFlow6Data.SpareRespiratoryCapacity);

                    if (WorkFlow6Data.SpareRespiratoryCapacity.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity, WorkFlow6Data.SpareRespiratoryCapacity);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Spare Respiratory Capacity widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(11)]
        public void NonMitochondrialRespiration()
        {
                            ExtentReport.CreateExtentTestNode("Non Mitochondrial Respiration");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.NonMitoO2Consumption))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.NonMitoO2Consumption;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.NonmitoO2Consumption);

                    graphProperties.Display(WorkFlow6Data.NonmitoO2Consumption);

                    graphProperties.Normalization(WorkFlow6Data.NonmitoO2Consumption);

                    graphProperties.ErrorFormat(WorkFlow6Data.NonmitoO2Consumption, WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption);

                    graphProperties.SortBy(WorkFlow6Data.NonmitoO2Consumption);

                    //Need to work on Graph Property - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.NonmitoO2Consumption.ExpectedGraphUnits, WidgetTypes.NonMitoO2Consumption);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.NonmitoO2Consumption.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.NonmitoO2Consumption);

                        graphSettings.ZeroLine(WorkFlow6Data.NonmitoO2Consumption);

                        graphSettings.Zoom(WorkFlow6Data.NonmitoO2Consumption);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.NonmitoO2Consumption.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption, WorkFlow6Data.NonmitoO2Consumption);

                    if (WorkFlow6Data.NonmitoO2Consumption.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption, WorkFlow6Data.NonmitoO2Consumption);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Non mito O2 Consumption widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(12)]
        public void ATPProduction()
        {
            ExtentReport.CreateExtentTestNode("ATP Production Coupled Respiration");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.AtpProductionCoupledRespiration))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.AtpProductionCoupledRespiration;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.ATPProductionCoupledRespiration);

                    graphProperties.Display(WorkFlow6Data.ATPProductionCoupledRespiration);

                    graphProperties.Normalization(WorkFlow6Data.ATPProductionCoupledRespiration);

                    graphProperties.ErrorFormat(WorkFlow6Data.ATPProductionCoupledRespiration, WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration);

                    graphProperties.SortBy(WorkFlow6Data.ATPProductionCoupledRespiration);

                    //Need to work on Graph Property - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.ATPProductionCoupledRespiration.ExpectedGraphUnits, WidgetTypes.AtpProductionCoupledRespiration);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.ATPProductionCoupledRespiration.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.ATPProductionCoupledRespiration);

                        graphSettings.ZeroLine(WorkFlow6Data.ATPProductionCoupledRespiration);

                        graphSettings.Zoom(WorkFlow6Data.ATPProductionCoupledRespiration);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.ATPProductionCoupledRespiration.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration, WorkFlow6Data.ATPProductionCoupledRespiration);

                    if (WorkFlow6Data.ATPProductionCoupledRespiration.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration, WorkFlow6Data.ATPProductionCoupledRespiration);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "ATP Production Coupled Respiration widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(13)]
        public void CouplingEfficiency()
        {
            ExtentReport.CreateExtentTestNode("Coupling Efficiency % ");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.CouplingEfficiencyPercent))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.CouplingEfficiencyPercent;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.CouplingEfficiency);

                    graphProperties.Display(WorkFlow6Data.CouplingEfficiency);

                    graphProperties.ErrorFormat(WorkFlow6Data.CouplingEfficiency, WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent);

                    graphProperties.SortBy(WorkFlow6Data.CouplingEfficiency);

                    //Need to work on Graph Property - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.CouplingEfficiency.ExpectedGraphUnits, WidgetTypes.CouplingEfficiencyPercent);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.CouplingEfficiency.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.CouplingEfficiency);

                        graphSettings.ZeroLine(WorkFlow6Data.CouplingEfficiency);

                        graphSettings.Zoom(WorkFlow6Data.CouplingEfficiency);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.CouplingEfficiency.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent, WorkFlow6Data.CouplingEfficiency);

                    if (WorkFlow6Data.CouplingEfficiency.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent, WorkFlow6Data.CouplingEfficiency);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Coupling Efficiency widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(14)]
        public void SpareRespiratoryCapacity()
        {
            ExtentReport.CreateExtentTestNode("Spare Respiratory Capacity");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.SpareRespiratoryCapacityPercent))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.SpareRespiratoryCapacityPercent;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                    graphProperties.Display(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                    graphProperties.ErrorFormat(WorkFlow6Data.SpareRespiratoryCapacityPercentage, WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent);

                    graphProperties.SortBy(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                    //Need to work on Graph Proiperty - Chart

                    graph.VerifyExpectedGraphUnits(WorkFlow6Data.SpareRespiratoryCapacityPercentage.ExpectedGraphUnits, WidgetTypes.SpareRespiratoryCapacityPercent);

                    graph.PanZoom(ChartType.Amchart);

                    //graph.GraphTootipVerificationWithRadius();

                    if (WorkFlow6Data.SpareRespiratoryCapacityPercentage.GraphSettingsVerify)
                    {
                        graphSettings.YAutoScale(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                        graphSettings.ZeroLine(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                        graphSettings.Zoom(WorkFlow6Data.SpareRespiratoryCapacityPercentage);
                    }

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow6Data.SpareRespiratoryCapacityPercentage.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent, WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                    if (WorkFlow6Data.SpareRespiratoryCapacityPercentage.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent, WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Spare Respiratory Capacity % widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(15)]
        public void DataTable()
        {
            ExtentReport.CreateExtentTestNode("Data Table");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.DataTable))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFCellMitoStressView();

                if (!driver.Title.Contains(fileUploadOrExistingFileData.FileName))
                    filesPage.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.DataTable);
                if (hasEditWidgetPageGone)
                {
                    WidgetTypes widgetType = WidgetTypes.DataTable;
                    widgetName = widgetType.ToString();

                    graphProperties.Oligo(WorkFlow6Data.DataTable);

                    graphProperties.Normalization(WorkFlow6Data.DataTable);

                    graphProperties.ErrorFormat(WorkFlow6Data.DataTable, WidgetCategories.XfMst, WidgetTypes.DataTable);

                    graphSettings.VerifyDataTableSettings();

                    plateMap.DataTableVerification();

                    //Unselect the Groups
                    groupLegends.EditWidgetDataTableGroupLegends(WidgetCategories.XfMst, WidgetTypes.DataTable, WorkFlow6Data.DataTable);

                    //Select the Groups
                    groupLegends.EditWidgetDataTableGroupLegends(WidgetCategories.XfMst, WidgetTypes.DataTable, WorkFlow6Data.DataTable);

                    if (WorkFlow6Data.DataTable.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.DataTable, WorkFlow6Data.DataTable);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Data Table widget is not required in Excel sheet selected as No.");
            }
        }
    }
}
