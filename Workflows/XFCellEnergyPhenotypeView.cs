using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Controls;
using SHAProject.EditPage;
using SHAProject.Utilities;
using SHAProject.PageObject;
using SHAProject.Page_Object;
using SHAProject.Create_Widgets;
using AventStack.ExtentReports;
using GraphSettings = SHAProject.EditPage.GraphSettings;

namespace SHAProject.Workflows
{

    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("XF Cell Energy Phenotype View")]
    public class XFCellEnergyPhenotypeView : Tests
    {
        public bool loginStatus;
        public Exports? exports;
        public PlateMap? plateMap;
        public LoginClass? loginClass;
        public UploadFile? uploadFile;
        public ModifyAssay? modifyAssay;
        public AnalysisPage? analysisPage;
        public GroupLegends? groupLegends;
        public Normalization? normalization;
        public GraphSettings? graphSettings;
        public GraphProperties? graphProperties;
        public CreateWidgetFromAddWidget? addWidgets;
        public CreateWidgetFromAddView? createWidgets;
        public static List<string> testidList = new List<string>();
        public static new readonly string currentPage = "XF Cell Energy Phenotype View";

        private static IEnumerable<string> GetTestFixtureBrowsers()
        {
            string buildPath = string.Empty;
            string excelPath = string.Empty;

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

            FileInfo fileInfo = new FileInfo(buildPath + excelPath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var browserList = new ArrayList();
            DataTable sheetData = new DataTable();

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Login"];

                foreach (var drawings in worksheet.Drawings)
                {
                    var checkbox = drawings as ExcelControlCheckBox;
                    var status = checkbox.Checked;

                    if (status.ToString() == "Checked")
                    {
                        browserList.Add(checkbox.Text);
                    }
                }

                worksheet = package.Workbook.Worksheets["Workflow-8"];

                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    DataRow dataRow = sheetData.NewRow();

                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;

                        if (row == 1)
                        {
                            sheetData.Columns.Add(cellValue != null ? cellValue.ToString() : "");
                        }
                        else
                        {
                            dataRow[col - 1] = cellValue;
                        }
                    }

                    sheetData.Rows.Add(dataRow);
                }

                testidList = sheetData.AsEnumerable().Select(r => r.Field<string>("Run Name")).ToList();
            }

            foreach (var browser in browserList)
            {
                yield return browser.ToString();
            }
        }

        public XFCellEnergyPhenotypeView(string browser)
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
            commonFunc.SetDriver(driver);

            loginClass = new LoginClass(driver, loginData, commonFunc);
            loginStatus = loginClass.LoginAsExcelUser();

            ExtentReport.CreateExtentTest("WorkFlow -8 : XFCellEnergyPhenotype View");
            bool ExcelReadStatus = reader.ReadDataFromExcel("Workflow-8");

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
        public void CreateCellEnergyView()
        {
            ExtentReport.CreateExtentTestNode("CreateCellEnergyView");

            if (loginStatus)
            {
                bool FileStatus = false;
                if (fileUploadOrExistingFileData.IsFileUploadRequired)
                {
                    FileStatus = uploadFile.HomePageFileUpload();
                }
                else if (fileUploadOrExistingFileData.OpenExistingFile)
                {
                    FileStatus = uploadFile.SearchFilesInFileTab(fileUploadOrExistingFileData.FileName);
                }
                else
                {
                    Assert.Ignore("Both FileUpload status and Open existing file status is false");
                }

                if (FileStatus)
                {
                    createWidgets.CreateWidgets(WidgetCategories.XfCellEnergy, fileUploadOrExistingFileData.SelectedWidgets);
                }
                else
                {
                    Assert.Fail();
                }
            }
            else
            {
                Assert.Fail();
            }
        }

        [Test, Order(2)]
        public void CheckCellEnergyViewLayout()
        {
            ExtentReport.CreateExtentTestNode("CheckCellEnergyViewLayout");

            if (WorkFlow8Data.LayoutVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                analysisPage.AnalysisPageHeaderIcons();

                analysisPage.AnalysisPageWidgetElements(WidgetCategories.XfStandard, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "XF Cell Energy Phenotype View Layout Verification is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(3)]
        public void XFCellEnergyPhenotype()
        {
            ExtentReport.CreateExtentTestNode("XFCellEnergyPhenotype");

            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.XfCellEnergyPhenotype))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype);

                if (hasEditWidgetPageGone)
                {

                    graphProperties.Normalization(WorkFlow8Data.CellEnergyPhenotype);
                    graphProperties.ErrorFormat(WorkFlow8Data.CellEnergyPhenotype);

                    if (WorkFlow8Data.CellEnergyPhenotype.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype, WorkFlow8Data.CellEnergyPhenotype);

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype, WorkFlow8Data.CellEnergyPhenotype);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "XF Cell Energy Phenotype widget is not required in Excel sheet selected as No.");
            }
        }
    }
}
