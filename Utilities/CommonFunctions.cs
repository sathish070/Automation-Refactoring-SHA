using AventStack.ExtentReports;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using SeleniumExtras.PageObjects;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.Utilities
{
    public class CommonFunctions: Tests
    {
        //public IJavaScriptExecutor? jScript;
        public IWebDriver _driver;
        public void SetDriver(IWebDriver driver)
        {
            _driver = driver;
        }

        public string LogPath()
        {
            string CURRENT_BUILD_PATH = string.Empty;
            string logPath = string.Empty; 
            /* Check if OS platform is either MacOSX or Unix*/
            if (Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix)
            {
                /*Set log path for Mac OS*/
                CURRENT_BUILD_PATH = AppDomain.CurrentDomain.BaseDirectory.Replace("bin/Debug/net7.0/", "");
                logPath = CURRENT_BUILD_PATH + "/Logs/";
            }
            else
            {
                /* Set log path for Windows OS*/
                CURRENT_BUILD_PATH = AppDomain.CurrentDomain.BaseDirectory.Replace("bin\\Debug\\net7.0\\", "");
                logPath = CURRENT_BUILD_PATH + "\\Logs\\";
            }
            return CURRENT_BUILD_PATH;
        }
        public string GetTimestamp()
        {
            string dateTime = DateTime.Now.ToString("dd-MM-yyyy-h-mm-ss");
            return dateTime;
        }

        public void CreateDirectory(string currentBuildPath, string reportFolderName)
        {
            try
            {
                /* Check if the directory exists and create it if not*/
                if (!Directory.Exists(currentBuildPath + "\\" + reportFolderName) || !Directory.Exists(currentBuildPath + "/" + reportFolderName))
                {
                    /* Create directory and get directory info*/
                    string folderPath = Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix ? currentBuildPath + "/" + reportFolderName : currentBuildPath + "\\" + reportFolderName;
                    DirectoryInfo di = Directory.CreateDirectory(folderPath);
                    //ExtentReport.ExtentTest("ExtentTest", Status.Pass, "The Directory has been created successfully. The path is " + di.FullName);
                }
                else
                {
                    /* If directory already exists, log it and set status to false*/
                    string folderPath = Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix ? currentBuildPath + "/" + reportFolderName : currentBuildPath + "\\" + reportFolderName;
                    //ExtentReport.ExtentTest("ExtentTest", Status.Fail, "The Directory has been already created. The path is " + folderPath);
                }
            }
            catch (IOException ioex)
            {
                    extentTestNode.Log(Status.Fail, "An error occured in creating directory. The error is " + ioex.Message);
            }
        }

        public void HandleCurrentWindow()
        {
            string newWindowHandle = _driver.WindowHandles.Last();
            _driver.SwitchTo().Window(newWindowHandle);
        }

        public string GetCurrentPath()
        {
            string currentPath = new Uri(_driver.Url).AbsolutePath;
            return currentPath;
        }

        public string GetAddWidgetCatgName(WidgetCategories wCat)
        {
            string? widgetName = wCat switch
            {
                WidgetCategories.XfStandard => "standardgraphs",
                WidgetCategories.XfMst => "mstgraph",
                WidgetCategories.XfAtp => "atpinducedgraph",
                WidgetCategories.XfAtpScreening => "atpscreeninginducedgraph",
                WidgetCategories.XfAtpDose => "atpdoseinducedgraph",
                WidgetCategories.XfCellEnergy => "cellenergydgraph",
                WidgetCategories.XfSubOx => "suboxgraph",
                WidgetCategories.XfGra => "gragraph",
                WidgetCategories.XfTCell => "tcellgraph",
                WidgetCategories.XfTCellPersistence => "tcellpersistence",
                WidgetCategories.XfTCellFitness => "tcellfitness",
                WidgetCategories.XfMitoDose => "mitodosegraph",
                WidgetCategories.XfMitoScreening => "mitodosescreening",
                _ => "standardgraphs",
            };
            return widgetName;
        }

        public string GetChartTitle(WidgetCategories wCat, WidgetTypes wType)
        {
            /*Get the chat title name by selecting the widget category and widget types*/
            string? chartName = (wCat, wType) switch
            {
                // Quick View
                (WidgetCategories.XfStandard, WidgetTypes.BarChart) => "Bar Graph",
                (WidgetCategories.XfStandard, WidgetTypes.KineticGraph) => "Kinetic Graph",
                (WidgetCategories.XfStandard, WidgetTypes.KineticGraphEcar) => "Kinetic Graph",
                (WidgetCategories.XfStandard, WidgetTypes.KineticGraphPer) => "Kinetic Graph",
                (WidgetCategories.XfStandard, WidgetTypes.EnergyMap) => "Energy Map",
                (WidgetCategories.XfStandard, WidgetTypes.HeatMap) => "Heat Map",

                // Dose Response View
                (WidgetCategories.XfStandard, WidgetTypes.DoseResponse) => "Dose-Response",

                // XfMitochondrialRespiration View
                (WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration) => "Mitochondrial Respiration",
                (WidgetCategories.XfMst, WidgetTypes.Basal) => "Basal Respiration",
                (WidgetCategories.XfMst, WidgetTypes.AcuteResponse) => "Acute Response",
                (WidgetCategories.XfMst, WidgetTypes.ProtonLeak) => "Proton Leak",
                (WidgetCategories.XfMst, WidgetTypes.MaximalRespiration) => "Maximal Respiration",
                (WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity) => "Spare Respiratory Capacity",
                (WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption) => "Non-mitochondrial Oxygen Consu",
                (WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration) => "ATP-Production Coupled Respira",
                (WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent) => "Coupling Efficiency (%)",
                (WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent) => "Spare Respiratory Capacity (%)",

                // XfCellEnergyPhenotype View
                (WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype) => "XF Cell Energy Phenotype",
                (WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr) => "Metabolic Potential OCR",
                (WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar) => "Metabolic Potential ECAR",
                (WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr) => "Baseline OCR",
                (WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar) => "Baseline ECAR",
                (WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr) => "Stressed OCR",
                (WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar) => "Stressed ECAR",
                (WidgetCategories.XfCellEnergy, WidgetTypes.DataTable) => "Average Assay Parameter Calculations",
                _ => "",
            };
            return chartName;
        }

        public int GetWidgetPosition(WidgetCategories wCat, WidgetTypes wType)
        {
            int? widgetPosition = (wCat, wType) switch
            {
                // Quick View
                (WidgetCategories.XfStandard, WidgetTypes.KineticGraph) => 19,
                (WidgetCategories.XfStandard, WidgetTypes.KineticGraphEcar) => 19,
                (WidgetCategories.XfStandard, WidgetTypes.KineticGraphPer) => 19,
                (WidgetCategories.XfStandard, WidgetTypes.BarChart) => 6,
                (WidgetCategories.XfStandard, WidgetTypes.EnergyMap) => 16,
                (WidgetCategories.XfStandard, WidgetTypes.HeatMap) => 42,

                // Dose Response View
                (WidgetCategories.XfStandardDose, WidgetTypes.DoseResponse) => 43,

                // XfMitochondrialRespiration View
                (WidgetCategories.XfMst, WidgetTypes.MitochondrialRespiration) => 24,
                (WidgetCategories.XfMst, WidgetTypes.Basal) => 8,
                (WidgetCategories.XfMst, WidgetTypes.AcuteResponse) => 1,
                (WidgetCategories.XfMst, WidgetTypes.ProtonLeak) => 28,
                (WidgetCategories.XfMst, WidgetTypes.MaximalRespiration) => 20,
                (WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacity) => 29,
                (WidgetCategories.XfMst, WidgetTypes.NonMitoO2Consumption) => 25,
                (WidgetCategories.XfMst, WidgetTypes.AtpProductionCoupledRespiration) => 2,
                (WidgetCategories.XfMst, WidgetTypes.CouplingEfficiencyPercent) => 7,
                (WidgetCategories.XfMst, WidgetTypes.SpareRespiratoryCapacityPercent) => 30,
                (WidgetCategories.XfMst, WidgetTypes.DataTable) => 13,


                //XfAtpRateAssayView
                (WidgetCategories.XfAtp, WidgetTypes.MitoAtpProductionRate) => 23,
                (WidgetCategories.XfAtp, WidgetTypes.GlycoAtpProductionRate) => 17,
                (WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateData) => 4,
                (WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateBasal) => 3,
                (WidgetCategories.XfAtp, WidgetTypes.AtpProductionRateInduced) => 5,
                (WidgetCategories.XfAtp, WidgetTypes.EnergeticMapBasal) => 14,
                (WidgetCategories.XfAtp, WidgetTypes.EnergeticMapInduced) => 15,
                (WidgetCategories.XfAtp, WidgetTypes.XfAtpRateIndex) => 33,
                (WidgetCategories.XfAtp, WidgetTypes.DataTable) => 13,

                // XfCellEnergyPhenotype View
                (WidgetCategories.XfCellEnergy, WidgetTypes.XfCellEnergyPhenotype) => 34,
                (WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialOcr) => 21,
                (WidgetCategories.XfCellEnergy, WidgetTypes.MetabolicPotentialEcar) => 22,
                (WidgetCategories.XfCellEnergy, WidgetTypes.BaselineOcr) => 10,
                (WidgetCategories.XfCellEnergy, WidgetTypes.BaselineEcar) => 11,
                (WidgetCategories.XfCellEnergy, WidgetTypes.StressedOcr) => 31,
                (WidgetCategories.XfCellEnergy, WidgetTypes.StressedEcar) => 32,
                (WidgetCategories.XfCellEnergy, WidgetTypes.DataTable) => 13
            };
            return (int)widgetPosition;
        }

        public void MoveBackToAnalysisPage()
        {
            try
            {
                var backToAnalysis = _driver.FindElement(By.XPath("//a[@class='nav-link-back']"));
                backToAnalysis.Click();
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Unable to move back to analysis page.The error is {e.Message}");
            }
        }

        public int GetWellIndexFromLabel(FileType fileType, string label)
        {
            List<string> completeWellNames;

            if (fileType == FileType.Xfp)
            {
                completeWellNames = PlateMapName.GetXfpWellName();
            }
            else if (fileType == FileType.Xfe24)
            {
                completeWellNames = PlateMapName.GetXfe24WellName();
            }
            else
            {
                completeWellNames = PlateMapName.GetXfe96WellName();
            }

            int index = completeWellNames.IndexOf(label.ToUpper());
            if (index == -1)
                throw new Exception("Well label not found");

            return index;
        }

        public ChartType GetChartType(WidgetTypes wType)
        {
            if (WidgetTypes.KineticGraph == wType || WidgetTypes.KineticGraphEcar == wType || WidgetTypes.KineticGraphPer == wType || WidgetTypes.HeatMap == wType || WidgetTypes.DoseResponse == wType
                || WidgetTypes.MitochondrialRespiration == wType)
                return ChartType.CanvasJS;
            else
                return ChartType.Amchart;
        }

        public string GetGraphUnits(ChartType type)
        {
            Thread.Sleep(2000);
            if (type == ChartType.CanvasJS)
                return _driver.ExecuteJavaScript<string>("return currentchartsrc.options.axisY[0].title");
            else
                return _driver.ExecuteJavaScript<string>("return currentchartsrc._yAxes._values[0].title.currentText");
            //return driver.FindElement(By.CssSelector("[role='widget'] g[font-size] text tspan")).Text;
        }
    }
}
