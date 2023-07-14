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

namespace SHAProject.Workflows
{
    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("XF Cell Mito Stress View")]
    public class XFCellMitoStressView: Tests
    {
        public LoginClass? loginClass;
        public static readonly string currentPage = "XF Cell Mito Stress View";
        public bool loginStatus;
        public UploadFile? Upload;
        public CreateWidgetFromAddView? createView;
        public bool uploadStatus = false;
        public GraphProperties properties;
        public EditPage.GraphSettings graphSettings;
        public Exports exports;
        public AnalysisPage analysisPage;
        public PlateMap plateMap;

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

            ExtentReport.CreateExtentTest("XF Cell Mito Stress View");
            bool ExcelReadStatus = reader.ReadDataFromExcel("Workflow-6");
            Objectinitialized();
            //Thread.sleep(2000);

            if (ExcelReadStatus)
            {
                extentTest?.Log(Status.Pass, "Excel read status is true for " + currentPage);
            }
            else
            {
                extentTest?.Log(Status.Fail, "Excel read status is false for " + currentPage);
                return;
            }

        }

        public void Objectinitialized()
        {
            Upload = new UploadFile(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData);
            createView = new CreateWidgetFromAddView(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, commonFunc);
            properties = new GraphProperties(currentPage, driver,loginClass.findElements,commonFunc);
            graphSettings = new EditPage.GraphSettings(currentPage,driver,loginClass.findElements, commonFunc);
            exports = new Exports(currentPage, driver, loginClass.findElements,fileUploadOrExistingFileData,commonFunc);
            analysisPage = new AnalysisPage(currentPage,driver, loginClass.findElements,fileUploadOrExistingFileData,commonFunc);
            plateMap = new PlateMap(currentPage,driver, loginClass.findElements,commonFunc,fileUploadOrExistingFileData, fileUploadOrExistingFileData.FileType);
        }


        [Test, Order(1)]
        public void CreateXFCellMitoStressViewWidgets()
        {
            ExtentReport.CreateExtentTestNode("Upload File and Creating a MST View");
            if (loginStatus)
            {
                bool FileStatus = false;
                bool Searchedfile = false;
                if (fileUploadOrExistingFileData.IsFileUploadRequired)
                {
                    FileStatus = Upload.HomePageFileUpload();
                }
                else if (fileUploadOrExistingFileData.OpenExistingFile)
                {
                    Searchedfile = Upload.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);
                }
                else
                {
                    Assert.Ignore("Both FileUpload status and Open existing file status is false");
                }

                if (!FileStatus && Searchedfile)
                {
                    Thread.Sleep(5000);
                    createView?.CreateWidgets(WidgetCategories.XfMst, fileUploadOrExistingFileData.SelectedWidgets);
                }
                else
                {
                    Thread.Sleep(3000);
                    createView?.AddView(WidgetCategories.XfMst, fileUploadOrExistingFileData.SelectedWidgets);
                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(2)]
        public void Mitochondrial_Respiration()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Mitochondrial Respiration Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration);
            if (hasEditWidgetPageGone)
            {

                properties.Measurement(WorkFlow6Data.MitochondrialRespiration);

                properties.Rate(WorkFlow6Data.MitochondrialRespiration);

                properties.Display(WorkFlow6Data.MitochondrialRespiration);

                properties.Y(WorkFlow6Data.MitochondrialRespiration);

                properties.Normalization(WorkFlow6Data.MitochondrialRespiration);

                properties.ErrorFormat(WorkFlow6Data.MitochondrialRespiration);

                properties.BackgroundCorrection(WorkFlow6Data.MitochondrialRespiration);

                properties.Baseline(WorkFlow6Data.MitochondrialRespiration);

                properties.VerifyNormalizationUnits(WorkFlow6Data.MitochondrialRespiration.GraphUnits, WidgetTypes.MitochondrialRespiration, false);

                graphSettings.VerifyGraphSettings();

                graphSettings.GraphSettingsField(WorkFlow6Data.MitochondrialRespiration);

                if (WorkFlow6Data.MitochondrialRespiration.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.MitochondrialRespiration.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration, WorkFlow6Data.MitochondrialRespiration);
            }

        }

        [Test, Order(3)]
        public void Basal_Respiration()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Basal Respiration Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.Basal);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.BasalRespiration);

                properties.Display(WorkFlow6Data.BasalRespiration);

                properties.Normalization(WorkFlow6Data.BasalRespiration);

                properties.ErrorFormat(WorkFlow6Data.BasalRespiration);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.BasalRespiration.GraphUnits, WidgetTypes.Basal, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.BasalRespiration);

                if (WorkFlow6Data.BasalRespiration.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.BasalRespiration.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.Basal, WorkFlow6Data.BasalRespiration);

            }
        }

        [Test,Order(4)]
        public void Acute_Response()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Acute Response Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.AcuteResponse);
            if (hasEditWidgetPageGone)
            {
                properties.Display(WorkFlow6Data.AcuteResponse);

                properties.Normalization(WorkFlow6Data.AcuteResponse);

                properties.ErrorFormat(WorkFlow6Data.AcuteResponse);

                //SortBy

                properties.VerifyNormalizationUnits(WorkFlow6Data.AcuteResponse.GraphUnits, WidgetTypes.AcuteResponse, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.AcuteResponse);

                if (WorkFlow6Data.AcuteResponse.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.AcuteResponse.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.AcuteResponse, WorkFlow6Data.AcuteResponse);
            }
        }

        [Test, Order(5)]
        public void Proton_Leak()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Proton Leak Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.ProtonLeak);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.ProtonLeak);

                properties.Display(WorkFlow6Data.ProtonLeak);

                properties.Normalization(WorkFlow6Data.ProtonLeak);

                properties.ErrorFormat(WorkFlow6Data.ProtonLeak);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.ProtonLeak.GraphUnits, WidgetTypes.ProtonLeak, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.ProtonLeak);

                if (WorkFlow6Data.ProtonLeak.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.ProtonLeak.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.ProtonLeak, WorkFlow6Data.ProtonLeak);

            }
        }

        [Test, Order(6)]
        public void Maximal_Respiration()
        {

            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Maximal Respiration Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.MaximalRespiration);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.MaximalRespiration);

                properties.Display(WorkFlow6Data.MaximalRespiration);

                properties.Normalization(WorkFlow6Data.MaximalRespiration);

                properties.ErrorFormat(WorkFlow6Data.MaximalRespiration);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.MaximalRespiration.GraphUnits, WidgetTypes.MaximalRespiration, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.MaximalRespiration);

                if (WorkFlow6Data.MaximalRespiration.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.MaximalRespiration.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.MaximalRespiration, WorkFlow6Data.MaximalRespiration);
            }
        }

        [Test, Order(7)]
        public void Spare_Respiratory()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Spare Respiratory Capacity Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.SpareRespiratoryCapacity);

                properties.Display(WorkFlow6Data.SpareRespiratoryCapacity);

                properties.Normalization(WorkFlow6Data.SpareRespiratoryCapacity);

                properties.ErrorFormat(WorkFlow6Data.SpareRespiratoryCapacity);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.SpareRespiratoryCapacity.GraphUnits, WidgetTypes.SpareRespiratoryCapacity, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.SpareRespiratoryCapacity);

                if (WorkFlow6Data.SpareRespiratoryCapacity.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.SpareRespiratoryCapacity.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity, WorkFlow6Data.SpareRespiratoryCapacity);
            }
        }

        [Test, Order(8)]
        public void Non_Mitochondrial_Respiration()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Non Mitochondrial Respiration Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.NonmitoO2Consumption);

                properties.Display(WorkFlow6Data.NonmitoO2Consumption);

                properties.Normalization(WorkFlow6Data.NonmitoO2Consumption);

                properties.ErrorFormat(WorkFlow6Data.NonmitoO2Consumption);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.NonmitoO2Consumption.GraphUnits, WidgetTypes.NonMitoO2Consumption, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.NonmitoO2Consumption);

                if (WorkFlow6Data.NonmitoO2Consumption.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.NonmitoO2Consumption.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption, WorkFlow6Data.NonmitoO2Consumption);
            }
        }

        [Test,Order(9)]
        public void ATP_Production()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("ATP Production Coupled Respiration Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.ATPProductionCoupledRespiration);

                properties.Display(WorkFlow6Data.ATPProductionCoupledRespiration);

                properties.Normalization(WorkFlow6Data.ATPProductionCoupledRespiration);

                properties.ErrorFormat(WorkFlow6Data.ATPProductionCoupledRespiration);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.ATPProductionCoupledRespiration.GraphUnits, WidgetTypes.AtpProductionCoupledRespiration, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.ATPProductionCoupledRespiration);

                if (WorkFlow6Data.ATPProductionCoupledRespiration.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.ATPProductionCoupledRespiration.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration, WorkFlow6Data.ATPProductionCoupledRespiration);
            }
        }

        [Test,Order(10)]
        public void Coupling_Efficiency()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Coupling Efficiency (%) Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.CouplingEfficiency);

                properties.Display(WorkFlow6Data.CouplingEfficiency);

                properties.Normalization(WorkFlow6Data.CouplingEfficiency);

                properties.ErrorFormat(WorkFlow6Data.CouplingEfficiency);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.CouplingEfficiency.GraphUnits, WidgetTypes.CouplingEfficiencyPercent, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.CouplingEfficiency);

                if (WorkFlow6Data.CouplingEfficiency.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");
                }

                if (WorkFlow6Data.CouplingEfficiency.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent, WorkFlow6Data.CouplingEfficiency);
            }
        }

        [Test,Order(11)]
        public void Spare_Respiratory_Capacity()
        {
            string currentPath = commonFunc.GetCurrentPath();

            if (currentPath.Contains("Widget/Edit"))
                commonFunc.MoveBackToAnalysisPage();

            if (!currentPath.Contains("Analysis"))
                CreateXFCellMitoStressViewWidgets();

            ExtentReport.CreateExtentTestNode("Spare Respiratory Capacity Widget");
            bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent);
            if (hasEditWidgetPageGone)
            {
                properties.Oligo(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                properties.Display(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                properties.Normalization(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                properties.ErrorFormat(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                //sortby

                properties.VerifyNormalizationUnits(WorkFlow6Data.SpareRespiratoryCapacityPercentage.GraphUnits, WidgetTypes.SpareRespiratoryCapacityPercent, false);

                graphSettings.GraphSettingsField(WorkFlow6Data.SpareRespiratoryCapacityPercentage);

                if (WorkFlow6Data.SpareRespiratoryCapacityPercentage.CheckNormalizationWithPlateMap)
                {
                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionality();

                    plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in the current calculation");
                }

                if (WorkFlow6Data.SpareRespiratoryCapacityPercentage.IsExportRequired)
                    exports?.EditWidgetExports(WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent, WorkFlow6Data.SpareRespiratoryCapacityPercentage);
            }
        }

        [Test,Order(12)]
        public void Data_Table()
        {

        }
    }
}
