using AventStack.ExtentReports;
using NUnit.Framework;
using OfficeOpenXml;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.Workflows;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace SHAProject
{
    public class Tests
    {
        public static ExcelReader? reader;
        public string CURRENT_BUILD_PATH = string.Empty;
        public String current_browser;
        public LoginData? loginData { get; set; }
        protected CurrentBrowser? currentBrowser { get; set; }
        protected NormalizationData? normalizationData { get; set; }
        protected FileUploadOrExistingFileData? fileUploadOrExistingFileData { get; set; }
        protected WidgetItems? widgetItems { get; set; }
        protected WorkFlow1Data WorkFlow1Data { get; set; }
        protected WorkFlow6Data WorkFlow6Data { get; set; }
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
        public CommonFunctions? commonFunc;
        public static string loginFolderPath;
        public static string pathToBeCreated;
        public IWebDriver? driver;
        public DriverSetup? setup;

        [OneTimeSetUp]
        public void Setup()
        {
            commonFunc = new CommonFunctions();
            string currentBuildPath = commonFunc.LogPath();
            string timestamp = commonFunc.GetTimestamp();

            pathToBeCreated = "Logs\\SA_" + "Chrome" + "-" + timestamp.ToString();

            extentReport = ExtentReport.ExtentStart(currentBuildPath, pathToBeCreated, timestamp);
            RunBeforeAnyTests(currentBuildPath);

            loginFolderPath = currentBuildPath + pathToBeCreated;
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            ExtentReport.ExtentClose();
            ShutDownScreenAlwaysOn();
            setup.driver?.Close();
            setup.driver?.Quit();
        }

        public void RunBeforeAnyTests(string currentBulidPath)
        {
            //InitiateScreenAlwaysOn();
            extentTest = extentReport.CreateTest("Excel Reader");
            try
            {
                loginData = new LoginData();
                currentBrowser = new CurrentBrowser();
                normalizationData = new NormalizationData();
                fileUploadOrExistingFileData = new FileUploadOrExistingFileData();
                WorkFlow1Data = new WorkFlow1Data();
                WorkFlow6Data = new WorkFlow6Data();
                WorkFlow8Data = new WorkFlow8Data();
                FilesTabData = new FilesTabData();
                string CURRENT_BUILD_PATH = currentBulidPath;

                currentBrowser.BrowserName = "Chrome";

                reader = new ExcelReader(loginData, fileUploadOrExistingFileData, normalizationData, WorkFlow1Data, WorkFlow6Data,WorkFlow8Data, CURRENT_BUILD_PATH, currentBrowser, extentTest, FilesTabData);

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
                caffeineProcess.StartInfo.FileName = Path.Combine(CURRENT_BUILD_PATH + "Caffeine\\caffeine64.exe");
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
                extentTest.Log(Status.Fail, "Some error has occured in shutting down of Caffeine process for always screen On. The error is " + ex.Message);
            }
        }
    }
}