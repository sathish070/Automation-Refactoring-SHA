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

            ExtentReport.CreateExtentTest("WorkFlow -8 : XF Cell Energy Phenotype View");
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
        public void CreateCellEnergyView()
        {
            ExtentReport.CreateExtentTestNode("Create XF Cell Energy Phenotype View");

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
                    createWidgets?.CreateWidgets(WidgetCategories.XfCellEnergy, fileUploadOrExistingFileData.SelectedWidgets);
            }
            else
                Assert.Fail();
        }

        [Test, Order(2)]
        public void CheckCellEnergyViewLayout()
        {
            if (WorkFlow8Data.LayoutVerification)
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("CheckCellEnergyViewLayout");

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
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.XfCellEnergyPhenotype))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("XF Cell Energy Phenotype");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype);

                if (hasEditWidgetPageGone)
                {
                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.Normalization(WorkFlow8Data.CellEnergyPhenotype);

                    graphProperties.ErrorFormat(WorkFlow8Data.CellEnergyPhenotype, WidgetCategories.XfStandard, WidgetTypes.XfCellEnergyPhenotype);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.CellEnergyPhenotype.ExpectedGraphUnits, WidgetTypes.XfCellEnergyPhenotype);

                    if (WorkFlow8Data.CellEnergyPhenotype.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.CellEnergyPhenotype);

                        graphSettings.YAutoScale(WorkFlow8Data.CellEnergyPhenotype);

                        graphSettings.Zoom(WorkFlow8Data.CellEnergyPhenotype);

                        graphSettings.GraphSettingsApply();
                    }

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype, WorkFlow8Data.CellEnergyPhenotype);

                    if (WorkFlow8Data.CellEnergyPhenotype.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype, WorkFlow8Data.CellEnergyPhenotype);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "XF Cell Energy Phenotype widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(4)]
        public void MetabolicPotentialOCR()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.MetabolicPotentialOcr))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Metabolic Potential OCR");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr);
                if (hasEditWidgetPageGone)
                {
                    graphProperties.Display(WorkFlow8Data.MetabolicPotentialOCR);

                    graphProperties.ErrorFormat(WorkFlow8Data.MetabolicPotentialOCR, WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr);

                    graphProperties.SortBy(WorkFlow8Data.MetabolicPotentialOCR);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.MetabolicPotentialOCR.ExpectedGraphUnits, WidgetTypes.MetabolicPotentialOcr);

                    if (WorkFlow8Data.MetabolicPotentialOCR.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.MetabolicPotentialOCR);

                        graphSettings.ZeroLine(WorkFlow8Data.MetabolicPotentialOCR);

                        graphSettings.Zoom(WorkFlow8Data.MetabolicPotentialOCR);

                        graphSettings.GraphSettingsApply();
                    }

                    ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.MetabolicPotentialOcr);

                    if (platemapWellCountResult.Status)
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.MetabolicPotentialOcr}");
                    else
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.MetabolicPotentialOcr}");

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow8Data.MetabolicPotentialOCR.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr, WorkFlow8Data.MetabolicPotentialOCR);

                    if (WorkFlow8Data.MetabolicPotentialOCR.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr, WorkFlow8Data.MetabolicPotentialOCR);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Metabolic Potential OCR widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(5)]
        public void MetabolicPotentialECAR()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.MetabolicPotentialEcar))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Metabolic Potential ECAR");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar);
                if (hasEditWidgetPageGone)
                {
                    graphProperties.Display(WorkFlow8Data.MetabolicPotentialECAR);

                    graphProperties.ErrorFormat(WorkFlow8Data.MetabolicPotentialECAR, WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar);

                    graphProperties.SortBy(WorkFlow8Data.MetabolicPotentialECAR);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.MetabolicPotentialECAR.ExpectedGraphUnits, WidgetTypes.MetabolicPotentialEcar);

                    if (WorkFlow8Data.MetabolicPotentialECAR.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.MetabolicPotentialECAR);

                        graphSettings.ZeroLine(WorkFlow8Data.MetabolicPotentialECAR);

                        graphSettings.Zoom(WorkFlow8Data.MetabolicPotentialECAR);

                        graphSettings.GraphSettingsApply();
                    }

                    ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.MetabolicPotentialEcar);

                    if (platemapWellCountResult.Status)
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.MetabolicPotentialEcar}");
                    else
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.MetabolicPotentialEcar}");

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow8Data.MetabolicPotentialECAR.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar, WorkFlow8Data.MetabolicPotentialECAR);

                    if (WorkFlow8Data.MetabolicPotentialECAR.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar, WorkFlow8Data.MetabolicPotentialECAR);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Metabolic Potential ECAR widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(6)]
        public void BaselineOCR()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.BaselineOcr))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Baseline OCR");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr);

                if (hasEditWidgetPageGone)
                {
                    graphProperties.Display(WorkFlow8Data.BaselineOCR);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.Normalization(WorkFlow8Data.BaselineOCR);

                    graphProperties.ErrorFormat(WorkFlow8Data.BaselineOCR, WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr);

                    graphProperties.SortBy(WorkFlow8Data.BaselineOCR);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.BaselineOCR.ExpectedGraphUnits, WidgetTypes.BaselineOcr);

                    if (WorkFlow8Data.BaselineOCR.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.BaselineOCR);

                        graphSettings.ZeroLine(WorkFlow8Data.BaselineOCR);

                        graphSettings.Zoom(WorkFlow8Data.BaselineOCR);

                        graphSettings.GraphSettingsApply();
                    }

                    ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.BaselineOcr);

                    if (platemapWellCountResult.Status)
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.BaselineOcr}");
                    else
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.BaselineOcr}");

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow8Data.BaselineOCR.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr, WorkFlow8Data.BaselineOCR);

                    if (WorkFlow8Data.BaselineOCR.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr, WorkFlow8Data.BaselineOCR);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Baseline OCR widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(7)]
        public void BaselineECAR()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.BaselineEcar))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Baseline ECAR");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar);

                if (hasEditWidgetPageGone)
                {
                    graphProperties.Display(WorkFlow8Data.BaselineECAR);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.Normalization(WorkFlow8Data.BaselineECAR);

                    graphProperties.ErrorFormat(WorkFlow8Data.BaselineECAR, WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar);

                    graphProperties.SortBy(WorkFlow8Data.BaselineECAR);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.BaselineECAR.ExpectedGraphUnits, WidgetTypes.BaselineEcar);

                    if (WorkFlow8Data.BaselineECAR.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.BaselineECAR);

                        graphSettings.ZeroLine(WorkFlow8Data.BaselineECAR);

                        graphSettings.Zoom(WorkFlow8Data.BaselineECAR);

                        graphSettings.GraphSettingsApply();
                    }

                    ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.BaselineEcar);

                    if (platemapWellCountResult.Status)
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.BaselineEcar}");
                    else
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.BaselineEcar}");

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow8Data.BaselineECAR.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar, WorkFlow8Data.BaselineECAR);

                    if (WorkFlow8Data.BaselineECAR.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar, WorkFlow8Data.BaselineECAR);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Baseline ECAR widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(8)]
        public void StressedOCR()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.StressedOcr))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Stressed OCR");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr);

                if (hasEditWidgetPageGone)
                {
                    graphProperties.Display(WorkFlow8Data.StressedOCR);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.Normalization(WorkFlow8Data.StressedOCR);

                    graphProperties.ErrorFormat(WorkFlow8Data.StressedOCR, WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr);

                    graphProperties.SortBy(WorkFlow8Data.StressedOCR);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.StressedOCR.ExpectedGraphUnits, WidgetTypes.StressedOcr);

                    if (WorkFlow8Data.StressedOCR.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.StressedOCR);

                        graphSettings.ZeroLine(WorkFlow8Data.StressedOCR);

                        graphSettings.Zoom(WorkFlow8Data.StressedOCR);

                        graphSettings.GraphSettingsApply();
                    }

                    ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.StressedOcr);

                    if (platemapWellCountResult.Status)
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.StressedOcr}");
                    else
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.StressedOcr}");

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow8Data.StressedOCR.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr, WorkFlow8Data.StressedOCR);

                    if (WorkFlow8Data.StressedOCR.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr, WorkFlow8Data.StressedOCR);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Stressed OCR widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(9)]
        public void StressedECAR()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.StressedEcar))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Stressed ECAR");

                if (RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                    commonFunc.HandleCurrentWindow();

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar);

                if (hasEditWidgetPageGone)
                {
                    graphProperties.Display(WorkFlow8Data.StressedECAR);

                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.Normalization(WorkFlow8Data.StressedECAR);

                    graphProperties.ErrorFormat(WorkFlow8Data.StressedECAR, WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar);

                    graphProperties.SortBy(WorkFlow8Data.StressedECAR);

                    graphProperties.VerifyExpectedGraphUnits(WorkFlow8Data.StressedECAR.ExpectedGraphUnits, WidgetTypes.StressedEcar);

                    if (WorkFlow8Data.StressedECAR.GraphSettingsVerify)
                    {
                        graphSettings.VerifyGraphSettingsIcon();

                        graphSettings.XAutoScale(WorkFlow8Data.StressedECAR);

                        graphSettings.ZeroLine(WorkFlow8Data.StressedECAR);

                        graphSettings.Zoom(WorkFlow8Data.StressedECAR);

                        graphSettings.GraphSettingsApply();
                    }

                    ResultStatus platemapWellCountResult = plateMap.VerifyPlateMapRowandCloumnWell(WidgetTypes.StressedEcar);

                    if (platemapWellCountResult.Status)
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"{platemapWellCountResult.Message}{WidgetTypes.StressedEcar}");
                    else
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"{platemapWellCountResult.Message} {WidgetTypes.StressedEcar}");

                    plateMap.PlateMapIcons();

                    plateMap.PlateMapFunctionalities();

                    if (WorkFlow8Data.StressedECAR.CheckNormalizationWithPlateMap)
                        plateMap.VerifyNormalizationVal();

                    plateMap.WellDataPopup("A05", "Included in current calculation");

                    groupLegends.EditWidgetGroupLegends(WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar, WorkFlow8Data.StressedECAR);

                    if (WorkFlow8Data.StressedECAR.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar, WorkFlow8Data.StressedECAR);

                    commonFunc.MoveBackToAnalysisPage();
                }
            }
            else
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Stressed ECAR widget is not required in Excel sheet selected as No.");
            }
        }

        [Test, Order(10)]
        public void DataTable()
        {
            if (fileUploadOrExistingFileData.SelectedWidgets.Contains(WidgetTypes.DataTable))
            {
                string currentPath = commonFunc.GetCurrentPath();

                if (currentPath.Contains("Widget/Edit"))
                    commonFunc.MoveBackToAnalysisPage();

                if (!currentPath.Contains("Analysis"))
                    CreateCellEnergyView();

                ExtentReport.CreateExtentTestNode("Data Table");

                bool hasEditWidgetPageGone = analysisPage.GoToEditWidget(WidgetCategories.XfCellEnergy, WidgetTypes.DataTable);

                if (hasEditWidgetPageGone)
                {
                    if (fileUploadOrExistingFileData.IsNormalized)
                        graphProperties.Normalization(WorkFlow8Data.DataTable);

                    graphProperties.ErrorFormat(WorkFlow8Data.DataTable, WidgetCategories.XfCellEnergy, WidgetTypes.DataTable);

                    if (WorkFlow8Data.DataTable.DataTableSettingsVerify)
                        graphSettings.VerifyDataTableSettings();

                    plateMap.DataTableVerification();

                    groupLegends.EditWidgetDataTableGroupLegends(WidgetCategories.XfAtp, WidgetTypes.DataTable, WorkFlow8Data.DataTable);

                    if (WorkFlow8Data.DataTable.IsExportRequired)
                        exports?.EditWidgetExports(WidgetCategories.XfAtp, WidgetTypes.DataTable, WorkFlow8Data.DataTable);

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
