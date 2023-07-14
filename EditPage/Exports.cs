using System;
using System.Xml;
using System.Data;
using System.Linq;
using System.Text;
using System.Collections;
using SkiaSharp;
using OfficeOpenXml;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;

namespace SHAProject.EditPage
{
    public class Exports : Tests
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public Exports(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc)
        {
            _driver = driver;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
        }

        public void EditWidgetExports(WidgetCategories wCat, WidgetTypes wType, WidgetItems widget)
        {
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
            try
            {
                if (_driver.FindElements(By.CssSelector(".Export")).Count > 0)
                {
                    CanvasChartExports();
                }
                else if (_driver.FindElements(By.CssSelector(".amcharts-amexport-item-level-0")).Count > 0)
                {
                    AmChartExports();
                }
                else if (_driver.FindElements(By.CssSelector(".export_Heatmap_icon")).Count > 0)
                {
                    HeatmapExports();
                }

                Thread.Sleep(5000);
                VerifyExports(wCat, wType, widget);
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page exports functionality. The error is {e.Message}");
            }
        }

        private void AmChartExports()
        {
            try
            {
                // export icon 
                IWebElement exportIcon = _driver.FindElement(By.CssSelector(".amcharts-amexport-item-level-0"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // image option
                IWebElement exportOption = _driver.FindElement(By.XPath("//a[contains(text(),'Image')]"));
                _findElements.ActionsClass(exportOption);

                // export png
                IWebElement exportPNG = _driver.FindElement(By.CssSelector(".amcharts-amexport-item-png a"));
                _findElements.ClickElement(exportPNG, _currentPage, "Editwidget Page - Export PNG");

                // export icon
                Thread.Sleep(5000);
                exportIcon = _driver.FindElement(By.CssSelector(".amcharts-amexport-item-level-0"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // image option
                Thread.Sleep(5000);
                exportOption = _driver.FindElement(By.XPath("//a[contains(text(),'Image')]"));
                _findElements.ActionsClass(exportOption);

                // export jpg
                IWebElement exportJPG = _driver.FindElement(By.CssSelector(".amcharts-amexport-item-jpg a"));
                _findElements.ClickElement(exportJPG, _currentPage, "Editwidget Page - Export JPG");

                // export icon 
                Thread.Sleep(5000);
                exportIcon = _driver.FindElement(By.CssSelector(".amcharts-amexport-item-level-0"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // data option
                Thread.Sleep(5000);
                exportOption = _driver.FindElement(By.XPath("//a[contains(text(),'Data')]"));
                _findElements.ActionsClass(exportOption);

                // export data
                var elements = _driver.FindElements(By.CssSelector(".amcharts-amexport-item-custom a"));
                if (elements.Count > 0)
                {
                    _findElements.ClickElement(elements[0], _currentPage, "Editwidget Page - Export Excel"); // data excel
                    _findElements.ClickElement(elements[1], _currentPage, "Editwidget Page - Export Prism"); // data prism
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page amchart exports functionality. The error is {e.Message}");
            }
        }

        private void CanvasChartExports()
        {
            try
            {
                // export icon 
                IWebElement exportIcon = _driver.FindElement(By.CssSelector(".Export"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // image option
                Thread.Sleep(5000);
                IWebElement exportOption = _driver.FindElement(By.CssSelector(".export-image-menu-list"));
                _findElements.ActionsClass(exportOption);

                // export jpg
                Thread.Sleep(5000);
                IWebElement exportJPG = _driver.FindElement(By.CssSelector("li.export-image-menu.jpg"));
                _findElements.ClickElement(exportJPG, _currentPage, "Editwidget Page - Export JPG");

                // export icon 
                Thread.Sleep(5000);
                exportIcon = _driver.FindElement(By.CssSelector(".Export"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // image option
                Thread.Sleep(5000);
                exportOption = _driver.FindElement(By.CssSelector(".export-image-menu-list"));
                _findElements.ActionsClass(exportOption);

                // export png
                Thread.Sleep(5000);
                IWebElement exportPNG = _driver.FindElement(By.CssSelector("li.export-image-menu.png"));
                _findElements.ClickElement(exportPNG, _currentPage, "Editwidget Page - Export PNG");

                // export icon 
                Thread.Sleep(5000);
                exportIcon = _driver.FindElement(By.CssSelector(".Export"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // data option
                Thread.Sleep(5000);
                exportOption = _driver.FindElement(By.CssSelector(".export-data-menu-list"));
                _findElements.ActionsClass(exportOption);

                // export excel
                Thread.Sleep(5000);
                IWebElement exportExcel = _driver.FindElement(By.CssSelector("li.export-data-menu.Excel"));
                _findElements.ClickElement(exportExcel, _currentPage, "Editwidget Page - Export Excel");

                // export icon 
                Thread.Sleep(5000);
                exportIcon = _driver.FindElement(By.CssSelector(".Export"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // data option
                Thread.Sleep(5000);
                exportOption = _driver.FindElement(By.CssSelector(".export-data-menu-list"));
                _findElements.ActionsClass(exportOption);

                // export prism
                Thread.Sleep(5000);
                IWebElement exportPrism = _driver.FindElement(By.CssSelector("li.export-data-menu.Prism"));
                _findElements.ClickElement(exportPrism, _currentPage, "Editwidget Page - Export Prism");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page canvas chart exports functionality. The error is {e.Message}");
            }
        }

        private void HeatmapExports()
        {
            try
            {
                // export icon 
                IWebElement exportIcon = _driver.FindElement(By.CssSelector(".export_Heatmap_icon"));
                _findElements.ClickElement(exportIcon, _currentPage, "Editwidget Page - Export View Icon");

                // export excel
                IWebElement exportExcel = _driver.FindElement(By.CssSelector("#expexcel"));
                _findElements.ClickElement(exportExcel, _currentPage, "Editwidget Page - Export Excel");

                // export prism
                Thread.Sleep(5000);
                IWebElement exportPrism = _driver.FindElement(By.CssSelector("#expprism"));
                _findElements.ClickElement(exportPrism, _currentPage, "Editwidget Page - Export Prism");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page heatmap exports functionality. The error is {e.Message}");
            }
        }

        private void VerifyExports(WidgetCategories wCat, WidgetTypes wType, WidgetItems widget)
        {
            try
            {
                string DownloadPath = Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix ? loginFolderPath + "/" + _currentPage + "/Downloads/" : loginFolderPath + "\\" + _currentPage + "\\Downloads\\";
                var directoryInfo = new DirectoryInfo(DownloadPath);

                if (_driver.FindElements(By.CssSelector(".export_Heatmap_icon")).Count == 0)
                {
                    #region verify-images

                    var pngImage = directoryInfo.GetFiles("*.png").OrderByDescending(f => f.LastWriteTime).FirstOrDefault();
                    var jpgImage = directoryInfo.GetFiles("*.jpg").OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                    if (pngImage != null)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Exported Image is in PNG format.");
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "The Exported Image is not in PNG format.");
                    }

                    if (jpgImage != null)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Exported Image is in JPG format.");
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "The Exported Image is not in JPG format.");
                    }

                    if (IsValidImage(pngImage.FullName))
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Exported PNG Image successfully.");
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Exported PNG Image is corrupted.");
                    }

                    if (IsValidImage(jpgImage.FullName))
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Exported JPG Image successfully.");
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Exported JPG Image is corrupted.");
                    }

                    #endregion
                }
                #region verify-datas

                var excelFile = directoryInfo.GetFiles("*.xlsx").OrderByDescending(f => f.LastWriteTime).FirstOrDefault();
                var prismFile = directoryInfo.GetFiles("*.pzfx").OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                if (excelFile != null)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Exported Data is in Excel format.");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "The Exported Data is not in Excel format.");
                }

                if (prismFile != null)
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Exported Data is in Prism format.");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The Exported Data is not in Prism format.");
                }

                if (IsValidFile(prismFile.FullName))
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Exported Prism data successfully.");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "The Exported Data is not in Prism format.");
                }

                var chartName = _commonFunc.GetChartTitle(wCat, wType);
                if (wType == WidgetTypes.DataTable)
                {
                    chartName = "Average Assay Parameter Calculations";
                }

                if (wType == WidgetTypes.DataTable)
                {
                    ArrayList dataTable = new ArrayList() { "Basal Rates (Average)(1)", "Well Data Basal Rates (2)", };
                    ArrayList arrayListsheetNames = new ArrayList();


                    string filePath = excelFile.FullName;
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        // Get all the sheet names
                        foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                        {
                            arrayListsheetNames.Add(worksheet.Name);
                        }
                    }

                    List<string> list1 = new List<string>(dataTable.Cast<string>());
                    List<string> list2 = new List<string>(arrayListsheetNames.Cast<string>());
                    bool areEqual = list1.SequenceEqual(list2);
                    if (areEqual)
                    {
                        string sheetNames = string.Join(", ", arrayListsheetNames.Cast<string>());
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file names are: " + sheetNames);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Downloaded excel file is Invalid");
                    }
                    chartName = "Average Assay Parameter Calculations";
                }

                DataTable? downloadedExcelData = null;
                if (wType != WidgetTypes.DataTable)
                {
                    downloadedExcelData = GetExcelData(chartName, excelFile.FullName);
                }

                if (downloadedExcelData != null && wType != WidgetTypes.DataTable)
                {
                    var normalized = downloadedExcelData.Rows[0]["Normalized"].ToString();
                    var normalizedvalue = widget.Normalization ? "Yes" : "No";
                    if (normalized == normalizedvalue)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file Normalized value for " + chartName + " is " + normalized);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "The Exported Data is not in Prism format.");
                    }
                }

                if (wType != WidgetTypes.DoseResponse)
                {
                    IWebElement baseLine = _driver.FindElement(By.Id("baselineselection"));
                    if (baseLine.Displayed)
                    {
                        var baseline = downloadedExcelData.Rows[0]["Baseline"].ToString();
                        var baselinevalue = widget.Baseline == "OFF" ? "No" : "Yes";
                        if (baseline == baselinevalue)
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file Baseline value for " + chartName + " is " + baselinevalue);
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Downloaded excel file Baseline value is invaild  for " + chartName + " is " + baselinevalue);
                        }
                    }
                }

                IWebElement backgroundCorrection = _driver.FindElement(By.CssSelector(".graph-ms.bg-correction.hideprop"));
                if (backgroundCorrection.Displayed)
                {
                    var backgroundcorrection = downloadedExcelData.Rows[0]["Background Correction"].ToString();
                    var backgroundcorrectionvalue = widget.BackgroundCorrection.ToString();
                    if (backgroundcorrection == backgroundcorrectionvalue)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file Background Correction value for " + chartName + " is " + backgroundcorrection);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Downloaded excel file Background Correction value is invalid for " + chartName + " is " + backgroundcorrectionvalue);
                    }
                }

                if (downloadedExcelData != null && wType != WidgetTypes.DataTable)
                {
                    var widgettitle = downloadedExcelData.Rows[0]["Widget Title"].ToString();
                    if (widgettitle == chartName)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file name is " + widgettitle);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Downloaded excel file is Invalid");
                    }
                }

                IWebElement rateType = _driver.FindElement(By.CssSelector(".graph-ms.select-measurement.rate.hiderate"));
                if (rateType.Displayed)
                {
                    var datatype = downloadedExcelData.Rows[0]["Data Type"].ToString();
                    var datatypevalue = widget.Rate.ToString();
                    if (datatype == datatypevalue)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file Datatype for " + chartName + " is " + datatype);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Downloaded excel file Datatype is invalid for " + chartName + " is " + datatypevalue);
                    }
                }

                if (downloadedExcelData != null && wType != WidgetTypes.DataTable)
                {

                    var units = downloadedExcelData.Rows[0]["Units"].ToString();
                    var unitsvalue = widget.NormalizedGraphUnits.ToString();
                    if (unitsvalue != "")
                    {
                        string graphunits = string.Empty;
                        if (unitsvalue != "N/A")
                        {
                            string[] unitsList = unitsvalue.Split("(");
                            graphunits = unitsList[1].Replace(")", "");
                        }
                        if (unitsvalue == "N/A")
                        {
                            graphunits = "N/A";
                        }
                        if (units == graphunits)
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Downloaded excel file Units values for " + chartName + " is " + units);
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Downloaded excel file Units values is invalid for " + chartName + " is " + unitsvalue);
                        }
                    }
                    downloadedExcelData?.Dispose();
                }

                #endregion
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page verify exports functionality. The error is {e.Message}");
            }
        }

        protected bool IsValidImage(string imageName)
        {
            bool status = false;

            using (var stream = new MemoryStream(File.ReadAllBytes(imageName)))
            {
                using (var bitmap = SKBitmap.Decode(stream))
                {
                    if (bitmap != null)
                    {
                        status = true;
                    }
                    else
                    {
                        status = false;
                    }
                }
            }
            return status;
        }

        protected bool IsValidFile(string fileName)
        {
            bool status = false;

            try
            {
                XmlReader reader = XmlReader.Create(fileName);
                {
                    while (reader.Read())
                    {
                        if (reader != null)
                            status = true;
                    }
                }
            }
            catch (Exception)
            {
                status = false;
            }

            return status;
        }

        public DataTable GetExcelData(string sheetName, string excelPath)
        {
            if (sheetName == "")
            {
                return null;
            }
            FileInfo fileInfo = new FileInfo(excelPath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
            DataTable dtExcel = new DataTable();

            // Loop through each row and column
            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                // Create a new row in the DataTable
                DataRow dataRow = dtExcel.NewRow();

                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    // Get the cell value
                    var cellValue = worksheet.Cells[row, col].Value;
                    if (row == 1)
                    {
                        // Create a new column in the DataTable
                        dtExcel.Columns.Add(cellValue != null ? cellValue.ToString() : "");
                    }
                    else
                    {
                        dataRow[col - 1] = cellValue;
                    }
                }

                // Add the DataRow to the DataTable
                dtExcel.Rows.Add(dataRow);
            }

            var dtHeader1 = dtExcel.AsEnumerable().Select(r => r.Field<string>("Assay Name")).ToList();
            var dtHeader2 = dtExcel.AsEnumerable().Select(r => r.Field<string>(_fileUploadOrExistingFileData.FileName)).ToList();
            DataTable _table = new();
            DataRow dr = _table.NewRow();
            _table.Rows.Add(dr);
            try
            {
                for (int i = 0; i < dtHeader1.Count; i++)
                {
                    if (dtHeader1[i] != null)
                    {
                        _table.Columns.Add(dtHeader1[i]);
                        _table.Rows[0][columnName: dtHeader1[i]] = dtHeader2[i];
                    }
                    if (dtHeader1[i] == "Units")
                        break;
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in editwidget page exports GetExcelData functionality. The error is {e.Message}");
            }
            dtExcel.Dispose();
            GC.Collect();

            return _table;
        }
    }
}
