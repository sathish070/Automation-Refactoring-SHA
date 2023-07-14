using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AngleSharp.Dom;

namespace SHAProject.Utilities
{
    public static class ExtentReport 
    {
        public static ExtentReports? extentReport;
        public static ExtentTest? extentTest;
        public static ExtentTest? extentTestNode;

        public static ExtentReports ExtentStart(string screenshotPath, string pathToBeCreated, string timeStamp)
        {
            string dateTime = timeStamp;
            extentReport = new ExtentReports();

            string extentReportFile = @"QA Report - " + "Chrome" + " - " + dateTime + ".html";

            var htmlReporter = new ExtentV3HtmlReporter(screenshotPath + pathToBeCreated +"\\" + extentReportFile);
            htmlReporter.Config.ReportName = "<b>Seahorse Automation Testing</b>";
            htmlReporter.Config.DocumentTitle = "Seahorse Automation ";

            extentReport.AttachReporter(htmlReporter);
            extentReport.AddSystemInfo("Execution Time", dateTime);
            extentReport.AddSystemInfo("Environment", "QA");

            return extentReport;
        }

        public static void ExtentClose()
        {
            extentReport.Flush();
        }

        public static void CreateExtentTest(string createTest) 
        {
            extentTest = extentReport.CreateTest(createTest);
        }

        public static void CreateExtentTestNode(string createTest) 
        {
            extentTestNode = extentTest.CreateNode(createTest);
        }

        public static void ExtentTest(string extent,Status status, string info)
        {
            if (extent == "ExtentTest")
            {
                extentTest.Log(status, info);
            }
            else if (extent == "ExtentTestNode")
            {
                extentTestNode.Log(status, info);
            }
        }
        public static void ExtentScreenshot(string extent, Status status, string name, string reportScreenshotPath)
        {
            if (extent == "ExtentTest")
            {
                extentTest.Log(status, "Screenshot  - " + name, MediaEntityBuilder.CreateScreenCaptureFromPath(reportScreenshotPath).Build());
            }
            else
            {
                extentTestNode.Log(status, "Screenshot  - " + name, MediaEntityBuilder.CreateScreenCaptureFromPath(reportScreenshotPath).Build());
            }
        }
    }
}
