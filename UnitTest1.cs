using AventStack.ExtentReports;
using NUnit.Framework;
using OfficeOpenXml;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.Workflows;
using SharpCompress.Common;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace SHAProject
{
    public class Tests
    {
        public static ExcelReader? reader;
        public string currentBuildPath = string.Empty;
        public String current_browser;
        public LoginData? loginData { get; set; }
        protected CurrentBrowser? currentBrowser { get; set; }
        protected NormalizationData? normalizationData { get; set; }
        protected FileUploadOrExistingFileData? fileUploadOrExistingFileData { get; set; }
        protected WidgetItems? widgetItems { get; set; }
        protected WorkFlow5Data? WorkFlow5Data { get; set; }
        protected WorkFlow6Data? WorkFlow6Data { get; set; }
        protected WorkFlow7Data? WorkFlow7Data { get; set; }
        protected WorkFlow8Data? WorkFlow8Data { get; set; }
        public FilesTabData? FilesTabData { get; set; }
        public ExtentTest? extentTest;
        public ExtentTest? extentTestNode;
        public ExtentReports? extentReport;
        public ExcelPackage? excelPackage;
        public ExcelWorksheet? worksheet;
        public DataTable? dtExecutionStatus;
        public DataRow? dtExecutionRow;
        private readonly Process caffeineProcess = new();
        public CommonFunctions commonFunc;
        public static string loginFolderPath;
        public static string reportFolderName;
        public IWebDriver? driver;
        public DriverSetup? setup;
        public static string widgetName;

        [OneTimeSetUp]
        public void Setup()
        {
            commonFunc = new CommonFunctions();
            currentBuildPath = commonFunc.LogPath();
            string timestamp = commonFunc.GetTimestamp();

            reportFolderName = "Logs\\SHA_" + "Report" + "-" + timestamp.ToString();

            extentReport = ExtentReport.ExtentStart(currentBuildPath, reportFolderName, timestamp);
            RunBeforeAnyTests();

            loginFolderPath = currentBuildPath + reportFolderName;
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            ExtentReport.ExtentClose();
            ShutDownScreenAlwaysOn();
            FolderNames();
            setup.driver?.Close();
            setup.driver?.Quit();
        }

        public void RunBeforeAnyTests()
        {
            InitiateScreenAlwaysOn();
            extentTest = extentReport.CreateTest("Excel Reader");
            try
            {
                loginData = new LoginData();
                currentBrowser = new CurrentBrowser();
                normalizationData = new NormalizationData();
                fileUploadOrExistingFileData = new FileUploadOrExistingFileData();
                WorkFlow5Data = new WorkFlow5Data();
                WorkFlow6Data = new WorkFlow6Data();
                WorkFlow7Data = new WorkFlow7Data();
                WorkFlow8Data = new WorkFlow8Data();
                FilesTabData = new FilesTabData();

                currentBrowser.BrowserName = "Chrome";

                reader = new ExcelReader(loginData, fileUploadOrExistingFileData, normalizationData, WorkFlow5Data, WorkFlow6Data, WorkFlow7Data, WorkFlow8Data, currentBuildPath, currentBrowser, extentTest, FilesTabData);

                bool excelReadStatus = reader.ReadDataFromExcel("Login");
                if (excelReadStatus)
                {
                    extentTest.Log(Status.Pass, "Excel read status success for login page");
                }
                else
                {
                    extentTest.Log(Status.Fail, "Excel read status failed");
                    return;
                }
            }
            catch (Exception e)
            {
                extentTest.Log(Status.Fail, "Excel read status failed. The error is " + e.Message);
                return;
            }
        }

        private void InitiateScreenAlwaysOn()
        {
            try
            {
                caffeineProcess.StartInfo.FileName = Path.Combine(currentBuildPath + "Caffeine\\caffeine64.exe");
                caffeineProcess.Start();
            }
            catch (Exception ex)
            {
                extentTest.Log(Status.Fail, "Some error has occured in initiating Caffeine process for always screen On. The error is " + ex.Message);
            }
        }

        private void ShutDownScreenAlwaysOn()
        {
            try
            {
                caffeineProcess?.Kill();
            }
            catch (Exception ex)
            {
                extentTest.Log(Status.Fail, "Some error has occured in shutting down of caffeine process for always screen On. The error is " + ex.Message);
            }
        }

        public void FolderNames()
        {
            if (Directory.Exists(loginFolderPath))
            {
                reportFolderName = reportFolderName.Replace("Report", current_browser);
                string newFolderPath = Path.Combine(currentBuildPath, reportFolderName);
                Directory.Move(loginFolderPath, newFolderPath);

                var directoryInfo = new DirectoryInfo(newFolderPath);
                var htmlFile = directoryInfo.GetFiles("*.html").OrderByDescending(f => f.LastWriteTime).FirstOrDefault();
                string htmlContent = File.ReadAllText(htmlFile.FullName);
                htmlContent = htmlContent.Replace("SHA_Report", "SHA_" + current_browser);

                string newFile = newFolderPath+"\\"+htmlFile.Name;
                File.WriteAllText(newFile, htmlContent);
            }
        }
    }
}