using AventStack.ExtentReports;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml;
using SHAProject.PageObject;
using SHAProject.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using SHAProject.Create_Widgets;
using System.IO;

namespace SHAProject.Workflows
{
    [TestFixtureSource(nameof(GetTestFixtureBrowsers))]
    [Category("Files Tab")]
    public class FilesTab :Tests
    {
        public LoginClass? loginClass;
        public FilesPage? filesPage;
        public static readonly string currentPage = "Files Tab";
        public bool loginStatus;
        public bool files = false;

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
                worksheet = package.Workbook.Worksheets["FilesTab"];

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

        public FilesTab(String browser)
        {
            current_browser = browser;
        }

        [OneTimeSetUp]
        public void Setup()
        {
            setup = new DriverSetup();
            driver = setup.browser(current_browser, loginData.Website, pathToBeCreated);
            loginClass = new LoginClass(driver, loginData, commonFunc);
            loginStatus = loginClass.LoginAsExcelUser();
            commonFunc.CreateDirectory(loginFolderPath, currentPage);
            string loginFoldersPath = loginFolderPath + "\\" + currentPage;
            commonFunc.CreateDirectory(loginFoldersPath, "Success");
            commonFunc.CreateDirectory(loginFoldersPath, "Error");
            commonFunc.SetDriver(driver);

            ExtentReport.CreateExtentTest("FilesTab");
            bool ExcelReadStatus = reader.ReadDataFromExcel("FilesTab");
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

            ObjectInitalize();

        }

        public void ObjectInitalize()
        {
            filesPage = new FilesPage(currentPage, driver, loginClass.findElements, fileUploadOrExistingFileData, FilesTabData);
        }

        [Test,Order(1)]
        public void LayoutVerification() //Test ID - 1 Layout_Verification
        {
            ExtentReport.CreateExtentTestNode("LayoutVerification");
            if (loginStatus)
            {
                files = filesPage.FilesPageRedirect();
                if (files)
                {
                    filesPage.LayoutIconsVerification();  
                }
            }
            else
            {
             Assert.Ignore();
            }
        }

        [Test,Order(2)]
        public void Pagination()  //Test ID - 2 Pagenation
        {
            ExtentReport.CreateExtentTestNode("PaginationVerification");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.PagenationVerificattion(); 
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.PagenationVerificattion();
                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(3)]
        public void SearchBox() //Test ID - 3 Search_box
        {
            ExtentReport.CreateExtentTestNode("SearchBoxVerification");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.SearchboxVerification();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.SearchboxVerification();
                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(4)]
        public void New_AssayandProject() //Test ID - 4&5 New_Assay&Project
        {
            ExtentReport.CreateExtentTestNode("Create New Assay and Project");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.CreateNewAssayandProject();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.CreateNewAssayandProject();
                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(5)]
        public void New_Folder() //Test ID - 4&5 New_Assay&Project
        {
            ExtentReport.CreateExtentTestNode("Create New Folder");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.CreateNewFolder();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.CreateNewFolder();
                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(6)]
        public void FileUpload() {
            ExtentReport.CreateExtentTestNode("Upload a file functionality");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.fileUpload();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.fileUpload();

                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(7)]
        public void FolderandSubfolder()
        {
            ExtentReport.CreateExtentTestNode("Create Folder and Sub Folder Functionality");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.folderFunctions();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.folderFunctions();

                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test,Order(8)]
        public void RenameandDelete()
        {
            ExtentReport.CreateExtentTestNode("Rename and Delete Functionality");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.renameAnddelete();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.renameAnddelete();

                }
            }
            else
            {
                Assert.Ignore();
            }
        }

        [Test, Order(9)]
        public void HaederIconFunctionality()
        {
            ExtentReport.CreateExtentTestNode("Header Icon Functionality");
            if (loginStatus)
            {
                if (files)
                {
                    filesPage.HeaderIconFunction();
                }
                else
                {
                    files = filesPage.FilesPageRedirect();
                    filesPage.HeaderIconFunction();

                }
            }
            else
            {
                Assert.Ignore();
            }
        }
    }
}
