

using AventStack.ExtentReports;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml;
using SHAProject.Create_Widgets;
using SHAProject.PageObject;
using SHAProject.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
//using SeaHorseAutomation.EditPage;
//using SHAProject.Edit_Page;
using SHAProject.EditPage;
using SHAProject.Page_Object;
using static System.Net.WebRequestMethods;
using static System.Runtime.InteropServices.JavaScript.JSType;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using System.Globalization;
using SHAProject.SeleniumHelpers;
using GraphSettings = SHAProject.EditPage.GraphSettings;
using System.IO;

namespace SHAProject.Workflows
{
    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("XF ATP Rate Assay View")]
    public class XFATPRateAssayView : Tests
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
        public CreateWidgetFromAddView? createWidgets;
        public static new readonly string currentPage = "XF ATP Rate AssayView";

        public FindElements? findElements;
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
                    var status = checkbox.Checked;

                    if (status.ToString() == "Checked")
                    {
                        /* Add the name of the checked browser to the browserList*/
                        browserList.Add(checkbox.Text);
                    }
                }

                /* Read the "Workflow-1" worksheet to retrieve the test data*/
                worksheet = package.Workbook.Worksheets["Workflow-7"];

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

        public XFATPRateAssayView(string browser)
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

            ExtentReport.CreateExtentTest("WorkFlow -7 : XF ATP Rate Assay View");
            bool ExcelReadStatus = reader.ReadDataFromExcel("Workflow-7");

            if (ExcelReadStatus)
            {
                extentTest.Log(Status.Pass, "Excel read status is true for " + currentPage);
            }
            else
            {
                extentTest.Log(Status.Fail, "Excel read status is false for " + currentPage);
                return;
            }

            ObjectInitalized();
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

        }

        [Test, Order(1)]
        public void CreateXFATPRateAssayView()
        {
            ExtentReport.CreateExtentTestNode("Create XF ATP Rate Assay View");
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
                    createWidgets?.CreateWidgets(WidgetCategories.XfAtp, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
                Assert.Ignore();
        }

        [Test, Order(2)]
        public void CreateXFATPRateAssayViewLayout()
        {
            if (WorkFlow7Data.AnalysisLayoutVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateXFATPRateAssayView();

                ExtentReport.CreateExtentTestNode("CreateXFATPRateAssayViewLayout");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                analysisPage.AnalysisPageHeaderIcons();

                analysisPage.AnalysisPageWidgetElements(WidgetCategories.XfAtp, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The Layout Verification for Standard view is given in the excel sheet is selected as No");
        }

        [Test, Order(3)]
        public void MitoAtpProductionRatewidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("MitoATPProductionRateWidget");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.MitoAtpProductionRate);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.MitoAtpProductionRate;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.MitoATPProductionRate);

                graphProperties.Induced(WorkFlow7Data.MitoATPProductionRate);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.MitoATPProductionRate);

                graphProperties.ErrorFormat(WorkFlow7Data.MitoATPProductionRate, WidgetCategories.XfAtp, WidgetTypes.MitoAtpProductionRate);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.MitoATPProductionRate.ExpectedGraphUnits, WidgetTypes.MitoAtpProductionRate);

                if (WorkFlow7Data.MitoATPProductionRate.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.ZeroLine(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.Zoom(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow7Data.MitoATPProductionRate.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("B01", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.MitoAtpProductionRate, WorkFlow7Data.MitoATPProductionRate);

                if (WorkFlow7Data.MitoATPProductionRate.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.MitoAtpProductionRate, WorkFlow7Data.MitoATPProductionRate);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.MitoATPProductionRate);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.MitoAtpProductionRate);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.MitoATPProductionRate.ExpectedGraphUnits, WidgetTypes.MitoAtpProductionRate);


                //graphProperties.GraphTootipVerificatioWithRadius();
            }
        }

        [Test, Order(4)]
        public void GlycoAtpProductionRatewidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("Glyco ATP Production Rate");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.GlycoAtpProductionRate);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.GlycoAtpProductionRate;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.GlycoATPProductionRate);
                graphProperties.Induced(WorkFlow7Data.GlycoATPProductionRate);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.GlycoATPProductionRate);

                graphProperties.ErrorFormat(WorkFlow7Data.GlycoATPProductionRate, WidgetCategories.XfAtp, WidgetTypes.GlycoAtpProductionRate);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.GlycoATPProductionRate.ExpectedGraphUnits, WidgetTypes.GlycoAtpProductionRate);

                if (WorkFlow7Data.GlycoATPProductionRate.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.GlycoATPProductionRate);

                    graphSettings.ZeroLine(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.Zoom(WorkFlow7Data.GlycoATPProductionRate);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow7Data.GlycoATPProductionRate.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("B01", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.GlycoAtpProductionRate, WorkFlow7Data.GlycoATPProductionRate);

                if (WorkFlow7Data.GlycoATPProductionRate.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.GlycoAtpProductionRate, WorkFlow7Data.GlycoATPProductionRate);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.GlycoATPProductionRate);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.GlycoAtpProductionRate);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.GlycoATPProductionRate.ExpectedGraphUnits, WidgetTypes.GlycoAtpProductionRate);

                //graphProperties.GraphTootipVerificatioWithRadius();
            }
        }

        [Test, Order(5)]
        public void ATPProductionRateDataWidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("ATP Production Rate Data");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateData);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.AtpProductionRateData;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.ATPProductionRateData);

                graphProperties.Induced(WorkFlow7Data.ATPProductionRateData);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.ATPProductionRateData);

                graphProperties.ErrorFormat(WorkFlow7Data.ATPProductionRateData, WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateData);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.ATPProductionRateData.ExpectedGraphUnits, WidgetTypes.AtpProductionRateData);

                if (WorkFlow7Data.ATPProductionRateData.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.ATPProductionRateData);

                    graphSettings.ZeroLine(WorkFlow7Data.ATPProductionRateData);

                    graphSettings.Zoom(WorkFlow7Data.ATPProductionRateData);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow7Data.ATPProductionRateData.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("B01", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateData, WorkFlow7Data.ATPProductionRateData);

                if (WorkFlow7Data.ATPProductionRateData.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateData, WorkFlow7Data.ATPProductionRateData);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.ATPProductionRateData);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.AtpProductionRateData);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.ATPProductionRateData.ExpectedGraphUnits, WidgetTypes.AtpProductionRateData);

                //graphProperties.GraphTootipVerificatioWithRadius();
            }
        }

        [Test, Order(6)]
        public void ATPProductionRateBasalWidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("ATP Production Rate - Basal");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateBasal);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.AtpProductionRateBasal;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.ATPProductionRate_Basal);

                graphProperties.Induced(WorkFlow7Data.ATPProductionRate_Basal);

                graphProperties.Display(WorkFlow7Data.ATPProductionRate_Basal);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.ATPProductionRate_Basal);

                graphProperties.ErrorFormat(WorkFlow7Data.ATPProductionRate_Basal, WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateBasal);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.ATPProductionRate_Basal.ExpectedGraphUnits, WidgetTypes.AtpProductionRateBasal);

                if (WorkFlow7Data.ATPProductionRate_Basal.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.ATPProductionRate_Basal);

                    graphSettings.ZeroLine(WorkFlow7Data.ATPProductionRate_Basal);

                    graphSettings.Zoom(WorkFlow7Data.ATPProductionRate_Basal);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow7Data.GlycoATPProductionRate.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.VerifyNormalizationVal(); //Need to Verify

                plateMap.WellDataPopup("B01", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateBasal, WorkFlow7Data.ATPProductionRate_Basal);

                if (WorkFlow7Data.ATPProductionRate_Basal.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateBasal, WorkFlow7Data.ATPProductionRate_Basal);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.ATPProductionRate_Basal);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.AtpProductionRateBasal);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.ATPProductionRate_Basal.ExpectedGraphUnits, WidgetTypes.AtpProductionRateBasal);

                //graphProperties.BarGraphVerification();
            }
        }

        [Test, Order(7)]
        public void ATPProductionRateInducedWidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("ATP Production Rate - Induced");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateInduced);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.AtpProductionRateInduced;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.ATPproductionRate_Induced);

                graphProperties.Induced(WorkFlow7Data.ATPproductionRate_Induced);

                graphProperties.Display(WorkFlow7Data.ATPproductionRate_Induced);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.ATPproductionRate_Induced);

                graphProperties.ErrorFormat(WorkFlow7Data.ATPproductionRate_Induced, WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateInduced);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.ATPproductionRate_Induced.ExpectedGraphUnits, WidgetTypes.AtpProductionRateInduced);

                if (WorkFlow7Data.ATPproductionRate_Induced.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.ATPproductionRate_Induced);

                    graphSettings.ZeroLine(WorkFlow7Data.ATPproductionRate_Induced);

                    graphSettings.Zoom(WorkFlow7Data.ATPproductionRate_Induced);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow7Data.GlycoATPProductionRate.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("B01", "Included in  current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateInduced, WorkFlow7Data.ATPproductionRate_Induced);

                if (WorkFlow7Data.ATPproductionRate_Induced.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateInduced, WorkFlow7Data.ATPproductionRate_Induced);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.ATPproductionRate_Induced);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.ATPproductionRate_Induced.ExpectedGraphUnits, WidgetTypes.AtpProductionRateInduced);

                //graphProperties.BarGraphVerification();
            }
        }

        [Test, Order(8)]
        public void EnergeticMapBasalWidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("EnergeticMap - Basal");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.EnergeticMapBasal);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.AtpProductionRateInduced;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.EnergeticMap_Basal);

                graphProperties.Induced(WorkFlow7Data.EnergeticMap_Basal);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.EnergeticMap_Basal);

                graphProperties.ErrorFormat(WorkFlow7Data.EnergeticMap_Basal, WidgetCategories.XfAtp, WidgetTypes.EnergeticMapBasal);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.EnergeticMap_Basal.ExpectedGraphUnits, WidgetTypes.EnergeticMapBasal);

                if (WorkFlow7Data.EnergeticMap_Basal.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.EnergeticMap_Basal);

                    graphSettings.ZeroLine(WorkFlow7Data.EnergeticMap_Basal);

                    graphSettings.Zoom(WorkFlow7Data.EnergeticMap_Basal);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                if (WorkFlow7Data.EnergeticMap_Basal.CheckNormalizationWithPlateMap)
                    plateMap.VerifyNormalizationVal();

                plateMap.WellDataPopup("B01", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.EnergeticMapBasal, WorkFlow7Data.EnergeticMap_Basal);

                if (WorkFlow7Data.EnergeticMap_Basal.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.EnergeticMapBasal, WorkFlow7Data.EnergeticMap_Basal);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.EnergeticMap_Basal);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.EnergeticMapBasal);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.EnergeticMap_Basal.ExpectedGraphUnits, WidgetTypes.EnergeticMapBasal);

                //graphProperties.GraphTootipVerificatioWithRadius();
            }
        }

        [Test, Order(9)]
        public void EnergeticMapInducedWidget()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("EnergeticMap - Induced");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.EnergeticMapInduced);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.EnergeticMapInduced;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.EnergeticMap_Induced);

                graphProperties.Induced(WorkFlow7Data.EnergeticMap_Induced);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.EnergeticMap_Induced);

                graphProperties.ErrorFormat(WorkFlow7Data.EnergeticMap_Induced, WidgetCategories.XfAtp, WidgetTypes.EnergeticMapInduced);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.EnergeticMap_Induced.ExpectedGraphUnits, WidgetTypes.EnergeticMapInduced);

                if (WorkFlow7Data.EnergeticMap_Induced.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.EnergeticMap_Induced);

                    graphSettings.ZeroLine(WorkFlow7Data.EnergeticMap_Induced);

                    graphSettings.Zoom(WorkFlow7Data.EnergeticMap_Induced);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                plateMap.VerifyNormalizationVal(); //Need to Verify

                plateMap.WellDataPopup("B01", "Included in  current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.EnergeticMapInduced, WorkFlow7Data.EnergeticMap_Induced);

                if (WorkFlow7Data.EnergeticMap_Induced.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.EnergeticMapInduced, WorkFlow7Data.EnergeticMap_Induced);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.EnergeticMapInduced);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.EnergeticMap_Induced.ExpectedGraphUnits, WidgetTypes.EnergeticMapInduced);

                //graphProperties.EnegryMapGraphVerification();

                //graphProperties.GraphTootipVerificatioWithRadius();

            }
        }

        [Test, Order(10)]
        public void ATPRateIndexWidget()
        {

            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFATPRateAssayView();

            ExtentReport.CreateExtentTestNode("Workflow-7: ATPRateIndexWidget");

            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.XfAtpRateIndex);
            if (hasEditWidgetPageGone)
            {
                WidgetTypes widgetType = WidgetTypes.XfAtpRateIndex;
                widgetName = widgetType.ToString();

                graphProperties.Oligo(WorkFlow7Data.XFATPRateIndex);

                graphProperties.Induced(WorkFlow7Data.XFATPRateIndex);

                if (fileUploadOrExistingFileData.IsNormalized)
                    graphProperties.Normalization(WorkFlow7Data.XFATPRateIndex);

                graphProperties.ErrorFormat(WorkFlow7Data.XFATPRateIndex, WidgetCategories.XfAtp, WidgetTypes.XfAtpRateIndex);

                graphProperties.VerifyExpectedGraphUnits(WorkFlow7Data.XFATPRateIndex.ExpectedGraphUnits, WidgetTypes.XfAtpRateIndex);

                if (WorkFlow7Data.XFATPRateIndex.GraphSettingsVerify)
                {
                    graphSettings.VerifyGraphSettingsIcon();

                    graphSettings.YAutoScale(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.ZeroLine(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.Zoom(WorkFlow7Data.MitoATPProductionRate);

                    graphSettings.GraphSettingsApply();
                }

                plateMap.PlateMapIcons();

                plateMap.PlateMapFunctionalities();

                plateMap.VerifyNormalizationVal(); //Need to Verify

                plateMap.WellDataPopup("B01", "Included in current calculation");

                groupLegends.EditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.XfAtpRateIndex, WorkFlow7Data.XFATPRateIndex);

                if (WorkFlow7Data.XFATPRateIndex.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.XfAtpRateIndex, WorkFlow7Data.XFATPRateIndex);

                commonFunc.MoveBackToAnalysisPage();

                //graphSettings.VerifyGraphSettings();

                //graphSettings.GraphSettingsField(WorkFlow7Data.XFATPRateIndex);

                //graphSettings.VerifyGraphSettingsYmaxandYminValue();

                //plateMap.VerifyPlateMapRowAndColumnWell(WidgetTypes.XfAtpRateIndex);

                //graphProperties.VerifyNormalizationUnits(WorkFlow7Data.XFATPRateIndex.ExpectedGraphUnits, WidgetTypes.XfAtpRateIndex);

                //graphProperties.BarGraphVerification();

                //graphProperties.GraphTootipVerificatioWithRadius();
            }
        }

        //[Test, Order(11)]
        //public void DataTableWidget()
        //{
        //    string currentPath = commonFunc.GetCurrentPath();

        //    if (currentPath.Contains("Widget/Edit"))
        //        commonFunc.MoveBackToAnalysisPage();

        //    if (!currentPath.Contains("Analysis"))
        //        CreateXFATPRateAssayView();

        //    ExtentReport.CreateExtentTestNode("DataTable");

        //    bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfAtp, WidgetTypes.DataTable);
        //    if (hasEditWidgetPageGone)
        //    {
        //        WidgetTypes widgetType = WidgetTypes.DataTable;
        //        widgetName = widgetType.ToString();

        //        graphProperties.Oligo(WorkFlow7Data.DataTable);

        //        graphProperties.Induced(WorkFlow7Data.DataTable);

        //        if (fileUploadOrExistingFileData.IsNormalized)
        //            graphProperties.Normalization(WorkFlow7Data.DataTable);

        //        graphProperties.ErrorFormat(WorkFlow7Data.DataTable, WidgetCategories.XfAtp, WidgetTypes.DataTable);

        //        graphSettings.DataTableVerifyGraphSettings();

        //        plateMap.DataTableVerification();

        //        //Unselect the Groups
        //        groupLegends.DataTableEditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.DataTable, WorkFlow7Data.DataTable);

        //        //Select the Groups
        //        groupLegends.DataTableEditWidgetGroupLegends(WidgetCategories.XfAtp, WidgetTypes.DataTable, WorkFlow7Data.DataTable);

        //        if (WorkFlow7Data.DataTable.IsExportRequired)
        //            exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.DataTable, WorkFlow7Data.DataTable);

        //        commonFunc.MoveBackToAnalysisPage();
        //    }
        //}

        //[Test, Order(12)]
        //public void InvalidBufferFactorValue()
        //{
        //    string currentPath = commonFunc.GetCurrentPath();

        //    if (currentPath.Contains("Widget/Edit"))
        //        commonFunc.MoveBackToAnalysisPage();

        //    if (!currentPath.Contains("Analysis"))
        //        CreateXFATPRateAssayView();

        //    if (loginStatus)
        //    {
        //        ExtentReport.CreateExtentTestNode("Invalid Buffer Factor Value ");

        //        bool fileStatus = false;

        //        fileStatus = filesPage.FilesTabSelection();

        //        if (fileStatus)
        //        {
        //            createWidgets.AddView(WidgetCategories.XfAtp, fileUploadOrExistingFileData.SelectedWidgets);

        //        }
        //    }

        //    analysisPage.InvalidBufferFactorVerification();

        //}
    }
}
