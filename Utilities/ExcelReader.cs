using System;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Controls;
using AventStack.ExtentReports;
using DataTable = System.Data.DataTable;
using LicenseContext = OfficeOpenXml.LicenseContext;
using AventStack.ExtentReports.Utils;
using SHAProject.Page_Object;

namespace SHAProject.Utilities
{
    public class ExcelReader
    {    
        public ExtentTest? _extentTest;
        public ExtentTest? _extentTestNode;
        public ExtentReports? _extentReport;
        private readonly string rateEcar;
        private readonly LoginData _loginData;
        public readonly FilesTabData _filesTab;
        private readonly NormalizationData _normalization;
        private readonly WorkFlow5Data _workFlow5;
        private readonly WorkFlow8Data _workFlow8;
        private readonly string _currentBuildPath;
        private readonly CurrentBrowser _currentBrowser;
        public string PerRate { get; private set; }
        public string PerGraphUnits { get; private set; }
        public string PerNormGraphUnits { get; private set; }
        public static List<string> testidList = new List<string>();
        private readonly FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public ExcelReader(LoginData loginData, FileUploadOrExistingFileData fileUploadOrExistingFileData, NormalizationData normalization, WorkFlow5Data workFlow1, WorkFlow8Data workFlow8,
            string currentBuildPath, CurrentBrowser currentBrowser, ExtentTest extentTest,FilesTabData filesTab)
        {
            _loginData = loginData;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _normalization = normalization;
            _workFlow5 = workFlow1;
            _workFlow8 = workFlow8;
            _currentBuildPath = currentBuildPath;
            _currentBrowser = currentBrowser;
            _extentTest = extentTest;
            _filesTab = filesTab;   
        }

        public bool ReadDataFromExcel(string? sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new Exception("Excel sheetname is not given");

            try
            {
                string _excelTemplatePath = Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix ? "ExcelTemplate/AutomatedData.xlsx" : "ExcelTemplate\\AutomatedData.xlsx";
                string ExcelPath = _currentBuildPath + _excelTemplatePath;

                /* Create a FileInfo object using the path to the Excel file*/
                FileInfo fileInfo = new FileInfo(ExcelPath);

                /* Set the ExcelPackage license context to NonCommercial*/
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                /* Create a DataTable to store the test data*/
                DataTable sheetData = new DataTable();
                string selectedBrowser = String.Empty;

                /* Use a using statement to create an instance of ExcelPackage to read the Excel file*/
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    /* Read the worksheet*/
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                    /*Loop through each row and column*/
                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        /* Create a new row in the DataTable*/
                        DataRow dataRow = sheetData.NewRow();

                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            /*Get the cell value*/
                            var cellValue = worksheet.Cells[row, col].Value;
                            if (row == 1)
                            {
                                /*Create a new column in the DataTable*/
                                sheetData.Columns.Add(cellValue != null ? cellValue.ToString() : "");
                            }
                            else
                            {
                                dataRow[col - 1] = cellValue;
                            }
                        }
                        /*Add the DataRow to the DataTable*/
                        sheetData.Rows.Add(dataRow);
                    }

                    /*Check the selected box in the browser field in the excel sheet*/
                    if (sheetName == "Login")
                    {
                        List<string> browserList = new List<string>();
                        foreach (var drawings in worksheet.Drawings)
                        {
                            var checkbox = drawings as ExcelControlCheckBox;
                            var status = checkbox.Checked;

                            if (status.ToString() == "Checked")
                            {
                                selectedBrowser += checkbox.Text + ",";
                                browserList.Add(checkbox.Text);
                            }
                        }

                        if (selectedBrowser != string.Empty)
                        {
                            selectedBrowser = selectedBrowser.Remove(selectedBrowser.Length - 1, 1);
                        }
                        else
                        {
                            selectedBrowser = "Chrome";
                        }
                        sheetData.Rows[1][3] = selectedBrowser;
                    }
                }
                /*Pass the sheet data to the get excel tabel structure function to get the data table structure*/
                var table = GetExcelTableStructure(sheetData, sheetName);
                GC.Collect();
                /*Pass the data table structure to set the values in the variables*/
                bool excelDataStatus = FillExcelData(table, sheetName);
                table.Dispose();
                return excelDataStatus;
            }
            catch (Exception e)
            {
                _extentTest.Log(Status.Fail, "Error occured in reading excel. The Error is " + e.Message);
                return false;
            }
        }

        public static DataTable GetExcelTableStructure(DataTable dtExcel, string sheetName)
        {
            if (sheetName == "Normalization")
                return dtExcel;
            else
            {
                var dtRunName = dtExcel.AsEnumerable().Select(r => r.Field<string>("Run Name")).ToList();
                var dtHeaders = dtExcel.AsEnumerable().Select(r => r.Field<string>("Variables")).ToList();
                var dtValues = dtExcel.AsEnumerable().Select(r => r.Field<string>("Inputs")).ToList();
   
                DataTable _table = new();
                DataRow dr = _table.NewRow();
                _table.Rows.Add(dr);
                DataRow dr2 = _table.NewRow();
                _table.Rows.Add(dr2);
                var runname = string.Empty;
                for (int i = 0; i < dtHeaders.Count; i++)
                {
                    if (dtHeaders[i] != null & sheetName == "Login")
                    {
                        _table.Columns.Add(dtHeaders[i]);
                        _table.Rows[0][dtHeaders[i]] = dtValues[i];
                    }
                    else if (dtHeaders[i] != null)
                    {
                        runname = dtRunName[i] != null ? dtRunName[i] + "$" : runname;
                        _table.Columns.Add(runname + dtHeaders[i]);
                        _table.Rows[0][runname + dtHeaders[i]] = dtValues[i];
                    }
                }
                dtExcel.Dispose();
                return _table;
            }
        }

        private bool FillExcelData(DataTable table, string sheetName)
        {
            var message = string.Empty;
            DataColumnCollection tblcolumns = table.Columns;

            if (sheetName == "Login")
            {
                _loginData.BrowserName = table.Rows[0]["Browser"].ToString();
                if (string.IsNullOrEmpty(_currentBrowser.BrowserName))
                {
                    _extentTest.Log(Status.Fail, "Browser name is empty. So it will run in Chrome by Default");
                }
                else
                {
                    _extentTest.Log(Status.Pass, "The Selected browser is " + _currentBrowser.BrowserName);
                }

                _loginData.Website = table.Rows[0]["URL"].ToString();
                if (string.IsNullOrEmpty(_loginData.Website))
                {
                    _extentTest.Log(Status.Fail, "Web site is empty : " + _loginData.Website);
                    message += "Website&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Website: " + _loginData.Website);
                }

                _loginData.UserName = table.Rows[0]["UserName"].ToString();
                if (string.IsNullOrEmpty(_loginData.UserName))
                {
                    _extentTest.Log(Status.Fail, "Username is empty :" + _loginData.UserName);
                    message += "UserName&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Username: " + _loginData.UserName);
                }

                _loginData.Password = table.Rows[0]["Password"].ToString();
                if (string.IsNullOrEmpty(_loginData.Password))
                {
                    _extentTest.Log(Status.Fail, "Password is empty: " + _loginData.Password);
                    message += "Password&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Password: " + _loginData.Password);
                }
            }
            else if (sheetName == "Normalization")
            {
                List<string> NormalizationWellValue = new();
                try
                {
                    NormalizationWellValue.Clear();
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        for (int j = 1; j < table.Columns.Count; j++)
                        {
                            if (table.Rows[i].ItemArray[j] != DBNull.Value)
                            {
                                if (i == 9)
                                {
                                    _normalization.ScaleFactor = table.Rows[i].ItemArray[j].ToString();
                                    break;
                                }
                                else if (i == 10)
                                {
                                    _normalization.Units = table.Rows[i].ItemArray[j].ToString();
                                    break;
                                }
                                else
                                    NormalizationWellValue.Add(table.Rows[i].ItemArray[j].ToString());
                            }
                            else
                                NormalizationWellValue.Add("0");
                        }
                    }

                    _normalization.Values = NormalizationWellValue;
                }
                catch (Exception)
                {
                    message += "NormalizationData";
                }
            }
            else if (sheetName == "Workflow-5")
            {
                #region TestId -1

                _fileUploadOrExistingFileData.IsFileUploadRequired = table.Rows[0]["Upload_File$IsFileUploadRequired"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.IsFileUploadRequired)
                {
                    _extentTest.Log(Status.Pass, "FileUpload required status is true");

                    _fileUploadOrExistingFileData.FileUploadPath = table.Rows[0]["Upload_File$FileUploadPath"].ToString();
                     _extentTest.Log(Status.Pass, "Fileupload path is present");
                    
                    _fileUploadOrExistingFileData.FileName = table.Rows[0]["Upload_File$FileName"].ToString();
                    _extentTest.Log(Status.Pass, "FileName is present - " + _fileUploadOrExistingFileData.FileName);
                    
                    _fileUploadOrExistingFileData.FileExtension = table.Rows[0]["Upload_File$FileExtension"].ToString();
                    _extentTest.Log(Status.Pass, "FileExtension is present - " + _fileUploadOrExistingFileData.FileExtension);
                }
                else
                {
                    _extentTest.Log(Status.Pass, "FileUpload required status is false");
                }

                _fileUploadOrExistingFileData.OpenExistingFile = table.Rows[0]["Upload_File$OpenExistingFile"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.OpenExistingFile)
                {
                    _extentTest.Log(Status.Pass, "Existing file name status is true");

                    _fileUploadOrExistingFileData.FileName = table.Rows[0]["Upload_File$FileName"].ToString();
                    if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.FileName))
                    {
                        _extentTest.Log(Status.Fail, "FileName status is false");
                        message += "FileName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "FileName status is true");
                    }

                    _fileUploadOrExistingFileData.FileExtension = table.Rows[0]["Upload_File$FileExtension"].ToString();
                    _extentTest.Log(Status.Pass, "FileExtension is present - " + _fileUploadOrExistingFileData.FileExtension);
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Existing file name status is false");
                }

                if (_fileUploadOrExistingFileData.IsFileUploadRequired && _fileUploadOrExistingFileData.OpenExistingFile)
                    _extentTest.Log(Status.Fail, "Both FileUpload required status and Open Existing File status is true");

                if (!_fileUploadOrExistingFileData.IsFileUploadRequired && !_fileUploadOrExistingFileData.OpenExistingFile)
                    _extentTest.Log(Status.Fail, "Both FileUpload required status and Open Existing File status is false");

                _fileUploadOrExistingFileData.IsTitrationFile = table.Rows[0]["Upload_File$IsTitrationFile"].ToString() == "Yes";
                _extentTest.Log(Status.Pass, "File Titration type is " + _fileUploadOrExistingFileData.IsTitrationFile);
                

                _fileUploadOrExistingFileData.IsNormalized = table.Rows[0]["Upload_File$IsNormalized"].ToString() == "Yes";
                _extentTest.Log(Status.Pass, "File Normalization status is " + _fileUploadOrExistingFileData.IsNormalized);

                var filetype = table.Rows[0]["Upload_File$FileType"].ToString();
                _fileUploadOrExistingFileData.FileType = filetype == "Xfe24" ? FileType.Xfe24 : filetype == "Xfe96" ? FileType.Xfe96 : filetype == "Xfp" ? FileType.Xfp : filetype == "XfHsMini" ? FileType.XfHsMini : filetype == "XFPro" ? FileType.XFPro : FileType.XFPro;
                _extentTest.Log(Status.Pass, "File Type is " + filetype);
                
                //ToD0:  Need to Log all the files.
                _fileUploadOrExistingFileData.SelectedWidgets = new List<WidgetTypes>();
                if (table.Rows[0]["Upload_File$OCR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.KineticGraph);
                }
                if (table.Rows[0]["Upload_File$ECAR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.KineticGraphEcar);
                }
                if (table.Rows[0]["Upload_File$PER"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.KineticGraphPer);
                }
                if (table.Rows[0]["Upload_File$Bar Graph"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.BarChart);
                }
                if (table.Rows[0]["Upload_File$Energetic Map"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.EnergyMap);
                }
                if (table.Rows[0]["Upload_File$Heat Map"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.HeatMap);
                }
                #endregion

                #region TestId -2 & 3

                if (tblcolumns.Contains("Layout_Verification$Layout Verification"))
                {
                    _workFlow5.AnalysisLayoutVerification = table.Rows[0]["Layout_Verification$Layout Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.AnalysisLayoutVerification ? "Analysis page layout verification is true" : "Analysis page layout verification is false");
                }
                #endregion

                #region TestId -4

                if (tblcolumns.Contains("Navg_Bar_Icons$DeleteWidgetRequired"))
                {
                    _workFlow5.DeleteWidgetRequired = table.Rows[0]["Navg_Bar_Icons$DeleteWidgetRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DeleteWidgetRequired ? "Delete Widget Required is true" : "Delete Widget Required is false");
                    if (_workFlow5.DeleteWidgetRequired)
                    {
                        if (tblcolumns.Contains("Navg_Bar_Icons$DeleteWidgetName"))
                        {
                            var widgetName = table.Rows[0]["Navg_Bar_Icons$DeleteWidgetName"].ToString();

                            _workFlow5.DeleteWidgetName = widgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                            widgetName == "Bar Chart" ? WidgetTypes.BarChart : widgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                            widgetName == "Heat Map" ? WidgetTypes.KineticGraph : WidgetTypes.KineticGraph;

                            _extentTest.Log(Status.Pass, "DeleteWidgetName is " + _workFlow5.DeleteWidgetName);
                            
                        }
                    }
                }

                if (tblcolumns.Contains("Navg_Bar_Icons$AddWidgetRequired"))
                {
                    _workFlow5.AddWidgetRequired = table.Rows[0]["Navg_Bar_Icons$AddWidgetRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.AddWidgetRequired ? "Add Widget Required is true" : "Add Widget Required is false");
                    if (_workFlow5.AddWidgetRequired)
                    {
                        //Kinetic Graph - Ocr, Kinetic Graph - Ecar, Kinetic Graph - Per, Bar Chart, Energetic Map, Heat Map
                        if (tblcolumns.Contains("Navg_Bar_Icons$AddWidgetName"))
                        {
                            var widgetName = table.Rows[0]["Navg_Bar_Icons$AddWidgetName"].ToString();

                            _workFlow5.AddWidgetName = widgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                            widgetName == "Bar Chart" ? WidgetTypes.BarChart : widgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                            widgetName == "Heat Map" ? WidgetTypes.KineticGraph : WidgetTypes.KineticGraph;

                            _extentTest.Log(Status.Pass, "AddWidgetName is " + _workFlow5.AddWidgetName);
                            
                        }
                    }
                }
                #endregion

                #region TestId -5

                if (tblcolumns.Contains("Normalization_Icon$Normalization Verification"))
                {
                    _workFlow5.NormalizationVerification = table.Rows[0]["Normalization_Icon$Normalization Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.NormalizationVerification ? "Normalization Verification is true" : "Normalization Verification is false");
                    ReadDataFromExcel("Normalization");
                }

                if (tblcolumns.Contains("Normalization_Icon$Apply to all widgets"))
                {
                    _workFlow5.ApplyToAllWidgets = table.Rows[0]["Normalization_Icon$Apply to all widgets"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, "ScaleFactor for Barchart is " + _workFlow5.ApplyToAllWidgets);
                }

                #endregion

                #region TestId -6

                if (tblcolumns.Contains("Modify_Assay$ModifyAssay Verification"))
                {
                    _workFlow5.ModifyAssay = table.Rows[0]["Modify_Assay$ModifyAssay Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.ModifyAssay ? "ModifyAssay Verification is true" : "ModifyAssay Verification is false");
                }

                if (tblcolumns.Contains("Modify_Assay$Add Group Name"))
                {
                    _workFlow5.AddGroupName = table.Rows[0]["Modify_Assay$Add Group Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.AddGroupName))
                    {
                        _extentTest.Log(Status.Fail, " Group Name is empty :" + _workFlow5.AddGroupName);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Add group Name is : " + _workFlow5.AddGroupName);
                    }
                }

                if (tblcolumns.Contains("Modify_Assay$Select Controls"))
                {
                    _workFlow5.SelecttheControls = table.Rows[0]["Modify_Assay$Select Controls"].ToString();
                    _extentTest.Log(Status.Pass, "Select the control is : " + _workFlow5.SelecttheControls);
                }

                if (tblcolumns.Contains("Modify_Assay$Injection Name"))
                {
                    _workFlow5.InjectionName = table.Rows[0]["Modify_Assay$Injection Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.InjectionName))
                    {
                        _extentTest.Log(Status.Fail, "The Given injection name is :" + _workFlow5.InjectionName);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given injection name is : " + _workFlow5.InjectionName);
                    }
                }
                #endregion

                #region TestId -7

                if (tblcolumns.Contains("Edit_Page$SelectWidgetName"))
                {
                    var selectWidgetName = table.Rows[0]["Edit_Page$SelectWidgetName"].ToString();
                    _workFlow5.SelectWidgetName = selectWidgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                    selectWidgetName == "Bar Chart" ? WidgetTypes.BarChart : selectWidgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                    selectWidgetName == "Heat Map" ? WidgetTypes.HeatMap : WidgetTypes.KineticGraph;

                     _extentTest.Log(Status.Pass, "AddWidgetName is " + _workFlow5.SelectWidgetName);
                }

                #endregion

                #region TestId -8

                _workFlow5.KineticGraphOcr = new WidgetItems();
                _workFlow5.KineticGraphOcr.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Kinetic_Graph$Measurement"))
                {
                    _workFlow5.KineticGraphOcr.Measurement = table.Rows[0]["Kinetic_Graph$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.KineticGraphOcr.Measurement))
                    {
                        _extentTest.Log(Status.Fail, "Measurement for Kinetic graph -OCR is missing");
                        message += "Measurement for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Measurement value for Kineticgraph -OCR  is " + _workFlow5.KineticGraphOcr.Measurement);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Rate-OCR"))
                {
                    _workFlow5.KineticGraphOcr.Rate = table.Rows[0]["Kinetic_Graph$Rate-OCR"].ToString();
                    _extentTest.Log(Status.Pass, "Ratetype for Kineticgraph -OCR  is " + _workFlow5.KineticGraphOcr.Rate);
                }

                if (tblcolumns.Contains("Kinetic_Graph$Display"))
                {
                    _workFlow5.KineticGraphOcr.Display = table.Rows[0]["Kinetic_Graph$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Kineticgraph -OCR  is " + _workFlow5.KineticGraphOcr.Display);
                }

                if (tblcolumns.Contains("Kinetic_Graph$Y"))
                {
                    _workFlow5.KineticGraphOcr.Y = table.Rows[0]["Kinetic_Graph$Y"].ToString();
                    _extentTest.Log(Status.Pass, "Y-toggle for Kineticgraph -OCR  is " + _workFlow5.KineticGraphOcr.Y);
                }

                if (tblcolumns.Contains("Kinetic_Graph$Normalization"))
                {
                    _workFlow5.KineticGraphOcr.Normalization = table.Rows[0]["Kinetic_Graph$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.Normalization ? "Normalization for Kinetic graph -OCR is true" : "Normalization for Kinetic graph -OCR is false");
                }

                if (tblcolumns.Contains("Kinetic_Graph$Error Format"))
                {
                    _workFlow5.KineticGraphOcr.ErrorFormat = table.Rows[0]["Kinetic_Graph$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Kineticgraph -OCR  is " + _workFlow5.KineticGraphOcr.ErrorFormat);
                }

                if (tblcolumns.Contains("Kinetic_Graph$Background Correction"))
                {
                    _workFlow5.KineticGraphOcr.BackgroundCorrection = table.Rows[0]["Kinetic_Graph$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.BackgroundCorrection ? "Background correction for Kinetic graph -OCR is true" : "Background for Kinetic graph -OCR is false");
                }

                if (tblcolumns.Contains("Kinetic_Graph$Baseline"))
                {
                    _workFlow5.KineticGraphOcr.Baseline = table.Rows[0]["Kinetic_Graph$Baseline"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.KineticGraphOcr.Baseline))
                    {
                        _extentTest.Log(Status.Fail, "Baseline for Kinetic graph -OCR is missing");
                        message += "Baseline for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Baseline for Kineticgraph -OCR  is " + _workFlow5.KineticGraphOcr.Baseline);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Expected GraphUnits-OCR"))
                {
                    _workFlow5.KineticGraphOcr.ExpectedGraphUnits = table.Rows[0]["Kinetic_Graph$Expected GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.KineticGraphOcr.ExpectedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Expected GraphUnits for Kinetic graph OCR value is missing");
                        message += "Expected GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Expected GraphUnits for Kinetic graph OCR value is " + _workFlow5.KineticGraphOcr.ExpectedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$GraphSettingsRequired"))
                {
                    _workFlow5.KineticGraphOcr.GraphSettingsVerify = table.Rows[0]["Kinetic_Graph$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(_workFlow5.KineticGraphOcr.GraphSettingsVerify ? Status.Pass : Status.Fail, _workFlow5.KineticGraphOcr.GraphSettingsVerify ? "GraphSettingsVerify for Kinetic graph is true" : "GraphSettingsVerify for Kinetic graph is false");
                    if (_workFlow5.KineticGraphOcr.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Kinetic_Graph$Remove Y AutoScale"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveYAutoScale = table.Rows[0]["Kinetic_Graph$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Kinetic graph is true" : "Remove Y AutoScale in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$Remove ZeroLine"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveZeroLine = table.Rows[0]["Kinetic_Graph$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for Kinetic graph is true" : "Remove ZeroLine in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$Remove LineMarkers"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveLineMarkers = table.Rows[0]["Kinetic_Graph$Remove LineMarkers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveLineMarkers ? "Remove LineMarkers in GraphSettings for Kinetic graph is true" : "Remove LineMarkers in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$Remove RateHighlight"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveRateHighlight = table.Rows[0]["Kinetic_Graph$Remove RateHighlight"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveRateHighlight ? "Remove RateHighlight in GraphSettings for Kinetic graph is true" : "Remove RateHighlight in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$Remove InjectionMarkers"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveInjectionMarkers = table.Rows[0]["Kinetic_Graph$Remove InjectionMarkers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveInjectionMarkers ? "Remove InjectionMarkers in GraphSettings for Kinetic graph is true" : "Remove InjectionMarkers in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$Remove Zoom"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveZoom = table.Rows[0]["Kinetic_Graph$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Kinetic graph is true" : "Remove Zoom in GraphSettings for Kinetic graph is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$CheckNormalizationWithPlateMap"))
                {
                    _workFlow5.KineticGraphOcr.CheckNormalizationWithPlateMap = table.Rows[0]["Kinetic_Graph$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow5.KineticGraphOcr.IsExportRequired = table.Rows[0]["Kinetic_Graph$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow5.KineticGraphOcr.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow5.KineticGraphOcr.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow5.KineticGraphOcr.IsExportRequired);
                    }
                }

                var EcarGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$GraphUnits-ECAR"))
                {
                    EcarGraphUnits = table.Rows[0]["Kinetic_Graph$GraphUnits-ECAR"].ToString();
                    if (string.IsNullOrEmpty(EcarGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Kinetic graph ECAR value is missing");
                        message += "GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Kinetic graph ECAR value is " + EcarGraphUnits);
                    }
                }

                var EcarExpectedGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$Expected GraphUnits-ECAR"))
                {
                    EcarExpectedGraphUnits = table.Rows[0]["Kinetic_Graph$Expected GraphUnits-ECAR"].ToString();
                    if (string.IsNullOrEmpty(EcarExpectedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Expected GraphUnits for Kinetic graph ECAR value is missing");
                        message += "Expected GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Expected GraphUnits for Kinetic graph ECAR value is " + EcarExpectedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Rate-ECAR"))
                {
                    var rateEcar = table.Rows[0]["Kinetic_Graph$Rate-ECAR"].ToString();
                    if (string.IsNullOrEmpty(rateEcar))
                    {
                        _extentTest.Log(Status.Fail, "Rate for Kinetic graph ECAR value is missing");
                        message += "Rate for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Rate for Kinetic graph ECAR value is " + rateEcar);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$PlateMap Sync to View"))
                {
                    _workFlow5.KineticGraphOcr.PlateMapSynctoView = table.Rows[0]["Kinetic_Graph$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Kinetic_Graph$GraphSettings Sync to View"))
                {
                    _workFlow5.KineticGraphOcr.GraphSettings.SynctoView = table.Rows[0]["Kinetic_Graph$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                _workFlow5.KineticGraphEcar = new WidgetItems()
                {
                    Measurement = _workFlow5.KineticGraphOcr.Measurement,
                    Rate = rateEcar,
                    Display = _workFlow5.KineticGraphOcr.Display,
                    Y = _workFlow5.KineticGraphOcr.Y,
                    Normalization = _workFlow5.KineticGraphOcr.Normalization,
                    ErrorFormat = _workFlow5.KineticGraphOcr.ErrorFormat,
                    BackgroundCorrection = _workFlow5.KineticGraphOcr.BackgroundCorrection,
                    Baseline = _workFlow5.KineticGraphOcr.Baseline,
                    GraphSettings = _workFlow5.KineticGraphOcr.GraphSettings,
                    GraphSettingsVerify = _workFlow5.KineticGraphOcr.GraphSettingsVerify,
                    ExpectedGraphUnits = EcarExpectedGraphUnits,
                    PlateMapSynctoView = _workFlow5.KineticGraphOcr.PlateMapSynctoView
                };

                var PerGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$GraphUnits-PER"))
                {
                    PerGraphUnits = table.Rows[0]["Kinetic_Graph$GraphUnits-PER"].ToString();
                    if (string.IsNullOrEmpty(PerGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Kinetic graph PER value is missing");
                        message += "GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Kinetic graph PER value is " + PerGraphUnits);
                    }
                }

                var PerExpectedGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$Expected GraphUnits-PER"))
                {
                    PerExpectedGraphUnits = table.Rows[0]["Kinetic_Graph$Expected GraphUnits-PER"].ToString();
                    if (string.IsNullOrEmpty(PerExpectedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Expected GraphUnits for Kinetic graph PER value is missing");
                        message += "Expected GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Expected GraphUnits for Kinetic graph PER value is " + PerExpectedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Rate-PER"))
                {
                    var PerRate = table.Rows[0]["Kinetic_Graph$Rate-PER"].ToString();
                    if (string.IsNullOrEmpty(PerRate))
                    {
                        _extentTest.Log(Status.Fail, "Rate for Kinetic graph PER value is missing");
                        message += "Rate for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Rate for Kinetic graph PER value is " + PerRate);
                    }
                }

                _workFlow5.KineticGraphPer = new WidgetItems()
                {
                    Measurement = _workFlow5.KineticGraphOcr.Measurement,
                    Rate = PerRate,
                    Display = _workFlow5.KineticGraphOcr.Display,
                    Y = _workFlow5.KineticGraphOcr.Y,
                    Normalization = _workFlow5.KineticGraphOcr.Normalization,
                    ErrorFormat = _workFlow5.KineticGraphOcr.ErrorFormat,
                    BackgroundCorrection = _workFlow5.KineticGraphOcr.BackgroundCorrection,
                    Baseline = _workFlow5.KineticGraphOcr.Baseline,
                    GraphSettings = _workFlow5.KineticGraphOcr.GraphSettings,
                    GraphSettingsVerify = _workFlow5.KineticGraphOcr.GraphSettingsVerify,
                    ExpectedGraphUnits = PerExpectedGraphUnits,
                    PlateMapSynctoView = _workFlow5.KineticGraphOcr.PlateMapSynctoView
                };

                #endregion

                #region TestId -9

                _workFlow5.Barchart = new WidgetItems();
                _workFlow5.Barchart.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Bar_Chart$Measurement"))
                {
                    _workFlow5.Barchart.Measurement = table.Rows[0]["Bar_Chart$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.Barchart.Measurement))
                    {
                        _extentTest.Log(Status.Fail, "Measurement for Barchart is missing");
                        message += "Measurement for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Measurement value for Barchart is " + _workFlow5.Barchart.Measurement);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Rate"))
                {
                    _workFlow5.Barchart.Rate = table.Rows[0]["Bar_Chart$Rate"].ToString();
                    _extentTest.Log(Status.Pass, " Rate for Barchart Rate is " + _workFlow5.Barchart.Rate);
                }

                if (tblcolumns.Contains("Bar_Chart$Display"))
                {
                    _workFlow5.Barchart.Display = table.Rows[0]["Bar_Chart$Display"].ToString();
                    _extentTest.Log(Status.Pass, " Display value for Barchart is " + _workFlow5.Barchart.Display);
                }

                if (tblcolumns.Contains("Bar_Chart$Normalization"))
                {
                    _workFlow5.Barchart.Normalization = table.Rows[0]["Bar_Chart$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.Normalization ? "Normalization for Barchart is true" : "Normalization for Barchart is false");
                }

                if (tblcolumns.Contains("Bar_Chart$Error Format"))
                {
                    _workFlow5.Barchart.ErrorFormat = table.Rows[0]["Bar_Chart$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, " Error Format for Barchart is " + _workFlow5.Barchart.ErrorFormat);
                }

                if (tblcolumns.Contains("Bar_Chart$Background Correction"))
                {
                    _workFlow5.Barchart.BackgroundCorrection = table.Rows[0]["Bar_Chart$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.BackgroundCorrection ? "Background Correction for Barchart is true" : "Background Correction for Barchart is false");
                }

                if (tblcolumns.Contains("Bar_Chart$Baseline"))
                {
                    _workFlow5.Barchart.Baseline = table.Rows[0]["Bar_Chart$Baseline"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.Barchart.Baseline))
                    {
                        _extentTest.Log(Status.Fail, "Baseline for Barchart is missing");
                        message += "Baseline for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Baseline for Barchart is " + _workFlow5.Barchart.Baseline);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Sort By"))
                {
                    _workFlow5.Barchart.SortBy = table.Rows[0]["Bar_Chart$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Bar Chart  is " + _workFlow5.Barchart.SortBy);
                }

                if (tblcolumns.Contains("Bar_Chart$Expected GraphUnits"))
                {
                    _workFlow5.Barchart.ExpectedGraphUnits = table.Rows[0]["Bar_Chart$Expected GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.Barchart.ExpectedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Expected GraphUnits for Barchart is missing");
                        message += "Expected GraphUnits for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Expected GraphUnits for Barchart is " + _workFlow5.Barchart.ExpectedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$GraphSettingsRequired"))
                {
                    _workFlow5.Barchart.GraphSettingsVerify = table.Rows[0]["Bar_Chart$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.GraphSettingsVerify ? "GraphSettingsVerify for Barchart is true" : "GraphSettingsVerify for Barchart is false");
                    if (_workFlow5.Barchart.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Bar_Chart$Remove Y AutoScale"))
                        {
                            _workFlow5.Barchart.GraphSettings.RemoveYAutoScale = table.Rows[0]["Bar_Chart$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.Barchart.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Barchart is true" : "Y AutoScale in GraphSettings for Barchart is false");
                        }

                        if (tblcolumns.Contains("Bar_Chart$Remove ZeroLine"))
                        {
                            _workFlow5.Barchart.GraphSettings.RemoveZeroLine = table.Rows[0]["Bar_Chart$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.Barchart.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for Barchart is true" : "ZeroLine in GraphSettings for Barchart is false");
                        }

                        if (tblcolumns.Contains("Bar_Chart$Remove Zoom"))
                        {
                            _workFlow5.Barchart.GraphSettings.RemoveZoom = table.Rows[0]["Bar_Chart$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.Barchart.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Barchart is true" : "Zoom in GraphSettings for Barchart is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$CheckNormalizationWithPlateMap"))
                {
                    _workFlow5.Barchart.CheckNormalizationWithPlateMap = table.Rows[0]["Bar_Chart$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    if (tblcolumns.Contains("Bar_Chart$IsExportRequired"))
                    {
                        _workFlow5.Barchart.IsExportRequired = table.Rows[0]["Bar_Chart$IsExportRequired"].ToString() == "Yes";
                        if (_workFlow5.Barchart.IsExportRequired)
                        {
                            _extentTest.Log(Status.Pass, "File Export status is " + _workFlow5.Barchart.IsExportRequired);
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "File Export status is " + _workFlow5.Barchart.IsExportRequired);
                        }
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$PlateMap Sync to View"))
                {
                    _workFlow5.Barchart.PlateMapSynctoView = table.Rows[0]["Bar_Chart$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Bar_Chart$GraphSettings Sync to View"))
                {
                    _workFlow5.Barchart.GraphSettings.SynctoView = table.Rows[0]["Bar_Chart$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                #endregion

                #region TestId -10

                _workFlow5.EnergyMap = new WidgetItems();
                _workFlow5.EnergyMap.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Energy_Map$Measurement"))
                {
                    _workFlow5.EnergyMap.Measurement = table.Rows[0]["Energy_Map$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.EnergyMap.Measurement))
                    {
                        _extentTest.Log(Status.Fail, " Measurement for EnergyMap is missing");
                        message += " Measurement for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Measurement value for EnergyMap is " + _workFlow5.EnergyMap.Measurement);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$Rate"))
                {
                    _workFlow5.EnergyMap.Rate = table.Rows[0]["Energy_Map$Rate"].ToString();
                    _extentTest.Log(Status.Pass, " Rate for EnergyMap is " + _workFlow5.EnergyMap.Rate);
                }

                if (tblcolumns.Contains("Energy_Map$Display"))
                {
                    _workFlow5.EnergyMap.Display = table.Rows[0]["Energy_Map$Display"].ToString();
                    _extentTest.Log(Status.Pass, " Display for EnergyMap is " + _workFlow5.EnergyMap.Display);
                }

                if (tblcolumns.Contains("Energy_Map$Normalization"))
                {
                    _workFlow5.EnergyMap.Normalization = table.Rows[0]["Energy_Map$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.Normalization ? " normalization for Energy graph is true" : " normalization for Energy graph is false");
                }

                if (tblcolumns.Contains("Energy_Map$Error Format"))
                {
                    _workFlow5.EnergyMap.ErrorFormat = table.Rows[0]["Energy_Map$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, " Error Format for EnergyMap is " + _workFlow5.EnergyMap.ErrorFormat);
                }

                if (tblcolumns.Contains("Energy_Map$Background Correction"))
                {
                    _workFlow5.EnergyMap.BackgroundCorrection = table.Rows[0]["Energy_Map$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.BackgroundCorrection ? " Background Correction for Energy graph is true" : " Background Correction for Energy graph is false");
                }

                if (tblcolumns.Contains("Energy_Map$BaseLine"))
                {
                    _workFlow5.EnergyMap.Baseline = table.Rows[0]["Energy_Map$BaseLine"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.EnergyMap.Baseline))
                    {
                        _extentTest.Log(Status.Fail, " Baseline for EnergyMap is missing");
                        message += " Baseline for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Baseline for EnergyMap is " + _workFlow5.EnergyMap.Baseline);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$Expected GraphUnits"))
                {
                    _workFlow5.EnergyMap.ExpectedGraphUnits = table.Rows[0]["Energy_Map$Expected GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.EnergyMap.ExpectedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for EnergyMap is missing");
                        message += "Expected GraphUnits for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Expected Graph Units for Energy Map is " + _workFlow5.EnergyMap.ExpectedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$GraphSettingsRequired"))
                {
                    _workFlow5.EnergyMap.GraphSettingsVerify = table.Rows[0]["Energy_Map$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.GraphSettingsVerify ? "GraphSettingsVerify for EnergyMap is true" : "GraphSettingsVerify for EnergyMap is false");

                    if (_workFlow5.EnergyMap.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Energy_Map$Remove X AutoScale"))
                        {
                            _workFlow5.EnergyMap.GraphSettings.RemoveXAutoScale = table.Rows[0]["Energy_Map$Remove X AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.GraphSettings.RemoveXAutoScale ? "Remove X AutoScale in GraphSettings for EnergyMap is true" : "X AutoScale in GraphSettings for EnergyMap is false");
                        }

                        if (tblcolumns.Contains("Energy_Map$Remove Y AutoScale"))
                        {
                            _workFlow5.EnergyMap.GraphSettings.RemoveYAutoScale = table.Rows[0]["Energy_Map$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for EnergyMap is true" : "Y AutoScale in GraphSettings for EnergyMap is false");
                        }

                        if (tblcolumns.Contains("Energy_Map$Remove Zoom"))
                        {
                            _workFlow5.EnergyMap.GraphSettings.RemoveZoom = table.Rows[0]["Energy_Map$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for EnergyMap is true" : "Zoom in GraphSettings for EnergyMap is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Energy_Map$CheckNormalizationWithPlateMap"))
                {
                    _workFlow5.EnergyMap.CheckNormalizationWithPlateMap = table.Rows[0]["Energy_Map$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Energy_Map$IsExportRequired"))
                {
                    _workFlow5.EnergyMap.IsExportRequired = table.Rows[0]["Energy_Map$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow5.EnergyMap.IsExportRequired);
                }

                if (tblcolumns.Contains("Energy_Map$PlateMap Sync to View"))
                {
                    _workFlow5.EnergyMap.PlateMapSynctoView = table.Rows[0]["Energy_Map$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Energy_Map$GraphSettings Sync to View"))
                {
                    _workFlow5.EnergyMap.GraphSettings.SynctoView = table.Rows[0]["Energy_Map$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.EnergyMap.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                #endregion

                #region TestId -11

                _workFlow5.HeatMap = new WidgetItems();
                _workFlow5.HeatMap.GraphSettings = new GraphSettings();
                _workFlow5.HeatMap.KitValidation = new KitValidation();
                _workFlow5.HeatMap.HeatTolerance = new HeatTolerance();

                if (tblcolumns.Contains("Heat_Map$Measurement"))
                {
                    _workFlow5.HeatMap.Measurement = table.Rows[0]["Heat_Map$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.HeatMap.Measurement))
                    {
                        _extentTest.Log(Status.Fail, " Measurement for HeatMap is missing");
                        message += " Measurement for HeatMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Measurement value for HeatMap is " + _workFlow5.HeatMap.Measurement);
                    }
                }

                if (tblcolumns.Contains("Heat_Map$Rate"))
                {
                    _workFlow5.HeatMap.Rate = table.Rows[0]["Heat_Map$Rate"].ToString();
                    _extentTest.Log(Status.Pass, " Rate for HeatMap is " + _workFlow5.HeatMap.Rate);
                }

                if (tblcolumns.Contains("Heat_Map$Normalization"))
                {
                    _workFlow5.HeatMap.Normalization = table.Rows[0]["Heat_Map$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.HeatMap.Normalization ? " normalization for HeatMap is true" : " normalization for HeatMap is false");
                }

                if (tblcolumns.Contains("Heat_Map$Background Correction"))
                {
                    _workFlow5.HeatMap.BackgroundCorrection = table.Rows[0]["Heat_Map$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.HeatMap.BackgroundCorrection ? " Background Correction for HeatMap is true" : " Background Correction for HeatMap is false");
                }

                if (tblcolumns.Contains("Heat_Map$BaseLine"))
                {
                    _workFlow5.HeatMap.Baseline = table.Rows[0]["Heat_Map$BaseLine"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.HeatMap.Baseline))
                    {
                        _extentTest.Log(Status.Fail, " Baseline for HeatMap is missing");
                        message += " Baseline for HeatMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Baseline for HeatMap is " + _workFlow5.HeatMap.Baseline);
                    }
                }

                if (tblcolumns.Contains("Heat_Map$GraphSettingsRequired"))
                {
                    _workFlow5.HeatMap.GraphSettingsVerify = table.Rows[0]["Heat_Map$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettingsVerify ? "GraphSettingsVerify for HeatMap is true" : "GraphSettingsVerify for HeatMap is false");
                    if (_workFlow5.HeatMap.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Heat_Map$IsExportRequired"))
                        {
                            if (tblcolumns.Contains("Heat_Map$Remove Y AutoScale"))
                            {
                                _workFlow5.HeatMap.GraphSettings.RemoveYAutoScale = table.Rows[0]["Heat_Map$Remove Y AutoScale"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for HeatMap is true" : "Remove Y AutoScale in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$Remove ZeroLine"))
                            {
                                _workFlow5.HeatMap.GraphSettings.RemoveZeroLine = table.Rows[0]["Heat_Map$Remove ZeroLine"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveZeroLine ? "Remove Zeroline in GraphSettings for HeatMap is true" : "Remove Zeroline in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$Remove LineMarkers"))
                            {
                                _workFlow5.HeatMap.GraphSettings.RemoveLineMarkers = table.Rows[0]["Heat_Map$Remove LineMarkers"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveLineMarkers ? "Remove LineMarkers in GraphSettings for HeatMap is true" : "Remove LineMarkers in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$Remove RateHighlight"))
                            {
                                _workFlow5.HeatMap.GraphSettings.RemoveRateHighlight = table.Rows[0]["Heat_Map$Remove RateHighlight"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveRateHighlight ? "Remove RateHighlight in GraphSettings for HeatMap is true" : "Remove RateHighlight in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$Remove InjectionMarkers"))
                            {
                                _workFlow5.HeatMap.GraphSettings.RemoveInjectionMarkers = table.Rows[0]["Heat_Map$Remove InjectionMarkers"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveInjectionMarkers ? "Remove InjectionMarkers in GraphSettings for HeatMap is true" : "Remove InjectionMarkers in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$Remove Zoom"))
                            {
                                _workFlow5.HeatMap.GraphSettings.RemoveZoom = table.Rows[0]["Heat_Map$Remove Zoom"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for HeatMap is true" : "Remove Zoom in GraphSettings for HeatMap is false");
                            }
                        }
                    }
                }

                if (tblcolumns.Contains("Heat_Map$AssayKit Validation"))
                {
                    _workFlow5.HeatMap.KitValidation.AssayKitValidation = table.Rows[0]["Heat_Map$AssayKit Validation"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.HeatMap.KitValidation.AssayKitValidation ? "Assaykit Validation for HeatMap is true" : "Assaykit Validation for HeatMap is false");
                    if (_workFlow5.HeatMap.KitValidation.AssayKitValidation)
                    {
                        if (tblcolumns.Contains("Heat_Map$Cat Number"))
                        {
                            _workFlow5.HeatMap.KitValidation.CatNumber = table.Rows[0]["Heat_Map$Cat Number"].ToString();
                            if (string.IsNullOrEmpty(_workFlow5.HeatMap.KitValidation.CatNumber))
                            {
                                _extentTest.Log(Status.Fail, "Cat Number for HeatMap is missing");
                                message += "Cat Number for HeatMap&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "Cat Number for HeatMap is " + _workFlow5.HeatMap.KitValidation.CatNumber);
                            }
                        }

                        if (tblcolumns.Contains("Heat_Map$Lot Number"))
                        {
                            _workFlow5.HeatMap.KitValidation.LotNumber = table.Rows[0]["Heat_Map$Lot Number"].ToString();
                            if (string.IsNullOrEmpty(_workFlow5.HeatMap.KitValidation.LotNumber))
                            {
                                _extentTest.Log(Status.Fail, "Lot Number for HeatMap is missing");
                                message += "Lot Number for HeatMap&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "Lot Number for HeatMap is " + _workFlow5.HeatMap.KitValidation.LotNumber);
                            }
                        }

                        if (tblcolumns.Contains("Heat_Map$SW ID"))
                        {
                            _workFlow5.HeatMap.KitValidation.SWID = table.Rows[0]["Heat_Map$SW ID"].ToString();
                            if (string.IsNullOrEmpty(_workFlow5.HeatMap.KitValidation.SWID))
                            {
                                _extentTest.Log(Status.Fail, "SW ID Number for HeatMap is missing");
                                message += "SW ID Number for HeatMap&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "SW ID Number for HeatMap is " + _workFlow5.HeatMap.KitValidation.SWID);
                            }
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$Colour Options"))
                    {
                        _workFlow5.HeatMap.HeatTolerance.ColourOptions = table.Rows[0]["Heat_Map$Colour Options"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _workFlow5.HeatMap.HeatTolerance.ColourOptions ? "Colour Options for HeatMap is true" : "Colour Options for HeatMap is false");
                        if (_workFlow5.HeatMap.HeatTolerance.ColourOptions)
                        {
                            if (tblcolumns.Contains("Heat_Map$Colour Tolerance %"))
                            {
                                _workFlow5.HeatMap.HeatTolerance.ColourTolerance = table.Rows[0]["Heat_Map$Colour Tolerance %"].ToString();
                                if (string.IsNullOrEmpty(_workFlow5.HeatMap.HeatTolerance.ColourTolerance))
                                {
                                    _extentTest.Log(Status.Fail, "Colour Tolerance % for HeatMap is missing");
                                    message += "Colour Tolerance % for HeatMap&";
                                }
                                else
                                {
                                    _extentTest.Log(Status.Pass, "Colour Tolerance for HeatMap is " + _workFlow5.HeatMap.HeatTolerance.ColourTolerance + " % ");
                                }
                            }
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$Expected GraphUnits"))
                    {
                        _workFlow5.HeatMap.ExpectedGraphUnits = table.Rows[0]["Heat_Map$Expected GraphUnits"].ToString();
                        if (string.IsNullOrEmpty(_workFlow5.HeatMap.ExpectedGraphUnits))
                        {
                            _extentTest.Log(Status.Fail, "Expected GraphUnits for HeatMap is missing");
                            message += "Expected GraphUnits for HeatMap&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "Expected Graph Units for Heat Map is " + _workFlow5.HeatMap.ExpectedGraphUnits);
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$CheckNormalizationWithPlateMap"))
                    {
                        _workFlow5.HeatMap.CheckNormalizationWithPlateMap = table.Rows[0]["Heat_Map$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _workFlow5.HeatMap.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                    }

                    if (tblcolumns.Contains("Heat_Map$IsExportRequired"))
                    {
                        _workFlow5.HeatMap.IsExportRequired = table.Rows[0]["Heat_Map$IsExportRequired"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, "File Export status is " + _workFlow5.HeatMap.IsExportRequired);
                    }
                }

                #endregion

                #region TestId -12

                _workFlow5.DoseResponseWidget = new WidgetItems();

                if (tblcolumns.Contains("Dose_Response_Add_Widget$Prerequisite"))
                {
                    _workFlow5.DoseResponseWidget.DoseResponseAddWidget = table.Rows[0]["Dose_Response_Add_Widget$DoseResponseAddWidget"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponseWidget.DoseResponseAddWidget ? "Add widget for Dose response is true" : "Add widget for Dose response is false");
                }

                #endregion

                #region TestId -13

                _workFlow5.DoseResponseView = new WidgetItems();

                if (tblcolumns.Contains("Dose_Response_Add_View$DoseResponseAddView"))
                {
                    _workFlow5.DoseResponseView.DoseResponseAddView = table.Rows[0]["Dose_Response_Add_View$DoseResponseAddView"].ToString() == "Yes";

                    if (_workFlow5.DoseResponseView.DoseResponseAddView)
                    {
                        _workFlow5.AddDoseWidget = new List<WidgetTypes>();
                        _workFlow5.AddDoseWidget.Add(WidgetTypes.DoseResponse);
                    }

                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponseView.DoseResponseAddView ? "Add view for Dose response is true" : "Add view for Dose response is false");
                }

                #endregion

                #region TestId -14

                _workFlow5.DoseResponse = new WidgetItems();
                _workFlow5.DoseResponse.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Dose_Response$Measurement"))
                {
                    _workFlow5.DoseResponse.Measurement = table.Rows[0]["Dose_Response$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.DoseResponse.Measurement))
                    {
                        _extentTest.Log(Status.Fail, " Measurement for Dose Response is missing");
                        message += " Measurement for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Measurement value for Dose Response is " + _workFlow5.DoseResponse.Measurement);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$Rate"))
                {
                    _workFlow5.DoseResponse.Rate = table.Rows[0]["Dose_Response$Rate"].ToString();
                    _extentTest.Log(Status.Pass, " Rate for Dose response is " + _workFlow5.DoseResponse.Rate);
                }

                if (tblcolumns.Contains("Dose_Response$Normalization"))
                {
                    _workFlow5.DoseResponse.Normalization = table.Rows[0]["Dose_Response$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.Normalization ? " normalization for Dose response is true" : " normalization for Dose response is false");
                }

                if (tblcolumns.Contains("Dose_Response$Error Format"))
                {
                    _workFlow5.DoseResponse.ErrorFormat = table.Rows[0]["Dose_Response$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, " Error Format for Dose Response is " + _workFlow5.DoseResponse.ErrorFormat);
                }

                if (tblcolumns.Contains("Dose_Response$Background Correction"))
                {
                    _workFlow5.DoseResponse.BackgroundCorrection = table.Rows[0]["Dose_Response$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.BackgroundCorrection ? " Background Correction for Dose Response is true" : " Background Correction for Dose Response is false");
                }

                if (tblcolumns.Contains("Dose_Response$GraphSettingsRequired"))
                {
                    _workFlow5.DoseResponse.GraphSettingsVerify = table.Rows[0]["Dose_Response$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettingsVerify ? "GraphSettingsVerify for Dose Response is true" : "GraphSettingsVerify for Dose Response is false");
                    if (_workFlow5.DoseResponse.GraphSettingsVerify)
                    {
                        // Dose kinetic graph properties
                        if (tblcolumns.Contains("Dose_Response$Remove Y AutoScale"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveYAutoScale = table.Rows[0]["Dose_Response$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Dose Kinetic graph is true" : "Remove Y AutoScale in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove ZeroLine"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveZeroLine = table.Rows[0]["Dose_Response$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for Dose Kinetic graph is true" : "Remove ZeroLine in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove LineMarkers"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveLineMarkers = table.Rows[0]["Dose_Response$Remove LineMarkers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveLineMarkers ? "Remove LineMarkers in GraphSettings for Dose Kinetic graph is true" : "Remove LineMarkers in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove RateHighlight"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveRateHighlight = table.Rows[0]["Dose_Response$Remove RateHighlight"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveRateHighlight ? "Remove RateHighlight in GraphSettings for Dose Kinetic graph is true" : "Remove RateHighlight in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove InjectionMarkers"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveInjectionMarkers = table.Rows[0]["Dose_Response$Remove InjectionMarkers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveInjectionMarkers ? "Remove InjectionMarkers in GraphSettings for Dose Kinetic graph is true" : "Remove InjectionMarkers in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove Zoom"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveZoom = table.Rows[0]["Dose_Response$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Dose Kinetic graph is true" : "Remove Zoom in GraphSettings for Dose Kinetic graph is false");
                        }

                        // Dose graph properties
                        if (tblcolumns.Contains("Dose_Response$Remove Dose X AutoScale"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDoseXAutoScale = table.Rows[0]["Dose_Response$Remove Dose X AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDoseXAutoScale ? "Remove Dose X AutoScale in GraphSettings for Dose Kinetic graph is true" : "Remove Dose X AutoScale in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove Dose Y AutoScale"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDoseYAutoScale = table.Rows[0]["Dose_Response$Remove Dose Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDoseYAutoScale ? "Remove Dose Y AutoScale in GraphSettings for Dose Kinetic graph is true" : "Remove Dose Y AutoScale in GraphSettings for Dose Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove Dose ZeroLine"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDoseZeroLine = table.Rows[0]["Dose_Response$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDoseZeroLine ? "Remove Dose Zeroline in GraphSettings for Dose response is true" : "Remove Dose Zeroline in GraphSettings for Dose Response is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove Dose LineMarkers"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDoseLineMarkers = table.Rows[0]["Dose_Response$Remove Dose LineMarkers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDoseLineMarkers ? "Remove Dose LineMarkers in GraphSettings for Dose response is true" : "Remove Dose LineMarkers in GraphSettings for Dose Response is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Remove Dose Zoom"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDoseZoom = table.Rows[0]["Dose_Response$Remove Dose Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDoseZoom ? "Remove Dose Zoom in GraphSettings for Dose Kinetic graph is true" : "Remove Dose Zoom in GraphSettings for Dose Kinetic graph is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Dose_Response$Expected GraphUnits"))
                {
                    _workFlow5.DoseResponse.ExpectedGraphUnits = table.Rows[0]["Dose_Response$Expected GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.DoseResponse.ExpectedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Expected GraphUnits for Dose Response is missing");
                        message += "Expected GraphUnits for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Expected Graph Units for Dose Response is " + _workFlow5.DoseResponse.ExpectedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$CheckNormalizationWithPlateMap"))
                {
                    _workFlow5.DoseResponse.CheckNormalizationWithPlateMap = table.Rows[0]["Dose_Response$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Dose_Response$IsExportRequired"))
                {
                    _workFlow5.DoseResponse.IsExportRequired = table.Rows[0]["Dose_Response$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, "File Export status is " + _workFlow5.DoseResponse.IsExportRequired);
                }

                if (tblcolumns.Contains("Dose_Response$PlateMap Sync to View"))
                {
                    _workFlow5.DoseResponse.PlateMapSynctoView = table.Rows[0]["Dose_Response$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Dose_Response$GraphSettings Sync to View"))
                {
                    _workFlow5.DoseResponse.GraphSettings.SynctoView = table.Rows[0]["Dose_Response$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Dose_Response$GraphSettings Dose Sync to View"))
                {
                    _workFlow5.DoseResponse.GraphSettings.DoseSynctoView = table.Rows[0]["Dose_Response$GraphSettings Dose Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.DoseSynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                #endregion

                #region TestId -15

                if (tblcolumns.Contains("Blank_View$CreateBlankView"))
                {
                    _workFlow5.CreateBlankView = table.Rows[0]["Blank_View$CreateBlankView"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.CreateBlankView ? "Create Blank View is true" : "Create Blank View is false");
                    if (_workFlow5.CreateBlankView)
                    {
                        //Kinetic Graph - Ocr, Kinetic Graph - Ecar, Kinetic Graph - Per, Bar Chart, Energetic Map, Heat Map

                        var widgetName = table.Rows[0]["Blank_View$AddBlankWidget"].ToString();
                         _workFlow5.AddBlankWidget = widgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                        widgetName == "Bar Chart" ? WidgetTypes.BarChart : widgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                        widgetName == "Heat Map" ? WidgetTypes.HeatMap : widgetName == "Dose Response" ? WidgetTypes.DoseResponse : WidgetTypes.KineticGraph;

                        _extentTest.Log(Status.Pass, "AddBlankWidget is " + _workFlow5.AddBlankWidget);
                    }
                }
                #endregion

                #region TestId -16

                if (tblcolumns.Contains("Custom_View$CustomViewName"))
                {
                    _workFlow5.CustomViewName = table.Rows[0]["Custom_View$CustomViewName"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.CustomViewName))
                    {
                        _extentTest.Log(Status.Fail, "Custom View Name is empty :" + _workFlow5.CustomViewName);
                        message += "CustomViewName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "CustomView Name is: " + _workFlow5.CustomViewName);
                    }
                }

                if (tblcolumns.Contains("Custom_View$CustomViewDescription"))
                {
                    _workFlow5.CustomViewDescription = table.Rows[0]["Custom_View$CustomViewDescription"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.CustomViewDescription))
                    {
                        _extentTest.Log(Status.Fail, "Custom view description is empty :" + _workFlow5.CustomViewDescription);
                        message += "CustomViewDescription&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Custom View description is: " + _workFlow5.CustomViewDescription);
                    }
                }
                #endregion
            }
            else if (sheetName == "FilesTab")
            {
                #region TestId - 1
                if (tblcolumns.Contains("Layout_Verification$Layout Verification"))
                {
                    _filesTab.LayoutVerification = table.Rows[0]["Layout_Verification$Layout Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.LayoutVerification ? "Layout Verification for the files tab is true" : "Layout Verifiaction for the files tab is false");
                }
                #endregion

                #region TestId - 2
                if (tblcolumns.Contains("Pagination$Pagination Verification"))
                {
                    _filesTab.PaginationVerification = table.Rows[0]["Pagination$Pagination Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.PaginationVerification ? "Pagination Verification for the files tab is true" : "Pagination Verifiaction for the files tab is false");


                    if (_filesTab.PaginationVerification)
                    {
                        if (tblcolumns.Contains("Pagination$Page Number"))
                        {
                            _filesTab.PageNumber = table.Rows[0]["Pagination$Page Number"].ToString();
                            if (string.IsNullOrEmpty(_filesTab.PageNumber))
                            {
                                _extentTest.Log(Status.Fail, "Page Number is missing");
                                message += "Page Number&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The given page number is" + _filesTab.PageNumber);
                            }
                        }

                        if (tblcolumns.Contains("Pagination$Files List"))
                        {
                            _filesTab.FilesList = table.Rows[0]["Pagination$Files List"].ToString();
                            if (string.IsNullOrEmpty(_filesTab.FilesList))
                            {
                                _extentTest.Log(Status.Fail, "Files List is missing");
                                message += "Files List&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The given file list number is" + _filesTab.FilesList);
                            }
                        }
                    }
                }
                #endregion

                #region TestId - 3
                _filesTab.searchBoxDataList = new();
                if (tblcolumns.Contains("Search_Box$File First Name"))
                {
                    _filesTab.FileFirstName = table.Rows[0]["Search_Box$File First Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.FileFirstName))
                    {
                        _extentTest.Log(Status.Fail, "File first name is missing");
                        message += "File first name&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given file first name is " + _filesTab.FileFirstName);
                        _filesTab.searchBoxDataList.Add(_filesTab.FileFirstName);
                    }
                }

                if (tblcolumns.Contains("Search_Box$File Middle Name"))
                {
                    _filesTab.FileMiddleName = table.Rows[0]["Search_Box$File Middle Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.FileFirstName))
                    {
                        _extentTest.Log(Status.Fail, "File middle name is missing");
                        message += "File middle name&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given file middle name is " + _filesTab.FileMiddleName);
                        _filesTab.searchBoxDataList.Add(_filesTab.FileMiddleName);
                    }
                }

                if (tblcolumns.Contains("Search_Box$File Last Name"))
                {
                    _filesTab.FileLastName = table.Rows[0]["Search_Box$File Last Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.FileLastName))
                    {
                        _extentTest.Log(Status.Fail, "File last name is missing");
                        message += "File last name&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given file last name is " + _filesTab.FileLastName);
                        _filesTab.searchBoxDataList.Add(_filesTab.FileLastName);
                    }
                }

                if (tblcolumns.Contains("Search_Box$File Full Name"))
                {
                    _filesTab.FileFullName = table.Rows[0]["Search_box$File Full Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.FileFullName))
                    {
                        _extentTest.Log(Status.Fail, "File full name is missing");
                        message += "File full name&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given file full name is " + _filesTab.FileFullName);
                        _filesTab.searchBoxDataList.Add(_filesTab.FileFullName);
                    }
                }

                if (tblcolumns.Contains("Search_Box$Categories"))
                {
                    _filesTab.Categories = table.Rows[0]["Search_Box$Categories"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.Categories))
                    {
                        _extentTest.Log(Status.Fail, "Categories  is missing");
                        message += "Categories&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given categories is " + _filesTab.Categories);
                        _filesTab.searchBoxDataList.Add(_filesTab.Categories);
                    }
                }

                if (tblcolumns.Contains("Search_Box$Date"))
                {
                    _filesTab.Date = (table.Rows[0]["Search_Box$Date"].ToString());
                    if (string.IsNullOrEmpty(_filesTab.Date))
                    {
                        _extentTest.Log(Status.Fail, "Date is missing");
                        message += "Date&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given date is " + _filesTab.Date);
                        _filesTab.searchBoxDataList.Add(_filesTab.Date);
                    }
                }

                if (tblcolumns.Contains("Search_Box$Instrument"))
                {
                    _filesTab.Instrument = table.Rows[0]["Search_Box$Instrument"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.Instrument))
                    {
                        _extentTest.Log(Status.Fail, "Instrument type is missing");
                        message += "Instrument type&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given Instrument type is " + _filesTab.Instrument);
                        _filesTab.searchBoxDataList.Add(_filesTab.Instrument);
                    }
                }

                if (tblcolumns.Contains("Search_Box$License"))
                {
                    _filesTab.License = table.Rows[0]["Search_Box$License"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.License))
                    {
                        _extentTest.Log(Status.Fail, "License is missing");
                        message += "License&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given License is " + _filesTab.License);
                        _filesTab.searchBoxDataList.Add(_filesTab.License);
                    }
                }
                #endregion

                #region TestId- 4 & 5

                #endregion

                #region TestId - 6
                if (tblcolumns.Contains("New_Folder$Create New Folder"))
                {
                    _filesTab.CreateNewFolder = table.Rows[0]["New_Folder$Create New Folder"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.CreateNewFolder ? "Create New Folder is true" : "Create New Folder is False");
                }

                if (tblcolumns.Contains("New_Folder$Folder Name"))
                {
                    _filesTab.FolderName = table.Rows[0]["New_Folder$Folder Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.FolderName))
                    {
                        _extentTest.Log(Status.Fail, "Folder Name is missing");
                        message += "FolderName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given folder name is " + _filesTab.FolderName);
                    }
                }

                if (tblcolumns.Contains("New_Folder$Sub Folder Name"))
                {
                    _filesTab.SubFolderName = table.Rows[0]["New_Folder$Sub Folder Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.SubFolderName))
                    {
                        _extentTest.Log(Status.Fail, "Sub Folder Name is missing");
                        message += "Sub Folder Name&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given sub folder name is " + _filesTab.SubFolderName);
                    }
                }

                if (tblcolumns.Contains("New_Folder$Last Folder Name"))
                {
                    _filesTab.LastFolderName = table.Rows[0]["New_Folder$Last Folder Name"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.LastFolderName))
                    {
                        _extentTest.Log(Status.Fail, "Last Folder Name is missing");
                        message += "Last Folder Name&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given sub folder name is " + _filesTab.LastFolderName);
                    }
                }
                #endregion

                #region TestId - 7
                _filesTab.FileUploadPath = table.Rows[0]["Upload_File$File Upload Path"].ToString();
                if (string.IsNullOrEmpty(_filesTab.FileUploadPath))
                {
                    _extentTest.Log(Status.Fail, "FileUploadPath  is empty");
                    message += "FileUploadPath&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Fileupload path is present");
                }

                _filesTab.FileName = table.Rows[0]["Upload_File$File Name"].ToString();
                if (string.IsNullOrEmpty(_filesTab.FileName))
                {
                    _extentTest.Log(Status.Fail, "FileName field is empty");
                    message += "FileName&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "FileName is present - " + _filesTab.FileName);
                }

                _filesTab.FileLocatedFolderPath = table.Rows[0]["Upload_File$File Located Folder Path"].ToString();
                if (string.IsNullOrEmpty(_filesTab.FileLocatedFolderPath))
                {
                    _extentTest.Log(Status.Fail, "File loacted folder path field is empty");
                    message += "File located folder path&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "File loacted folder path is " + _filesTab.FileLocatedFolderPath);
                }

                _filesTab.AddCategories = table.Rows[0]["Upload_File$Add Categories"].ToString();
                if (string.IsNullOrEmpty(_filesTab.AddCategories))
                {
                    _extentTest.Log(Status.Fail, "Add Categories field is empty");
                    message += "AddCategories&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Add categories is" + _filesTab.AddCategories);
                }
                #endregion

                #region TestId - 8

                #endregion

                #region TestId - 9
                if (tblcolumns.Contains("Rename_Delete$Rename"))
                {
                    _filesTab.Rename = table.Rows[0]["Rename_Delete$Rename"].ToString();
                    if (string.IsNullOrEmpty(_filesTab.Rename))
                    {
                        _extentTest.Log(Status.Fail, "Rename field is empty");
                        message += "Rename&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Rename for the folder is" + _filesTab.Rename);
                    }
                }

                if (tblcolumns.Contains("Rename_Delete$Delete The Folder"))
                {
                    _filesTab.DeleteFolder = table.Rows[0]["Rename_Delete$Delete The Folder"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.DeleteFolder ? "Delete the folder Verification for the files tab is true" : "Delete the folder Verification for the files tab is false");
                }
                #endregion

                #region TestId - 10
                if (tblcolumns.Contains("Header_Icons$Download File Verification"))
                {
                    _filesTab.DownloadFileVerification = table.Rows[0]["Header_Icons$Download File Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.DownloadFileVerification ? "Download File Verification for the files tab is true" : "Download File Verification for the files tab is false");
                }

                if (tblcolumns.Contains("Header_Icons$Make a Copy"))
                {
                    _filesTab.MakeACopy = table.Rows[0]["Header_Icons$Make a Copy"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.MakeACopy ? "Make a Copy Verification for the files tab is true" : "Make a Copy Verification for the files tab is false");

                    if (_filesTab.MakeACopy)
                    {
                        if (tblcolumns.Contains("Header_Icons$Copy File Path"))
                        {
                            _filesTab.CopyFilePath = table.Rows[0]["Header_Icons$Copy File Path"].ToString();
                            if (string.IsNullOrEmpty(_filesTab.CopyFilePath))
                            {
                                _extentTest.Log(Status.Fail, "Copy File Path is missing");
                                message += "Copy File Path&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The given copy file path is" + _filesTab.CopyFilePath);
                            }
                        }
                    }
                }

                if (tblcolumns.Contains("Header_Icons$Move To Folder"))
                {
                    _filesTab.MoveToFolder = table.Rows[0]["Header_Icons$Move To Folder"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.MoveToFolder ? "Move to folder Verification for the files tab is true" : "Move to folder Verification for the files tab is false");

                    if (_filesTab.MoveToFolder)
                    {

                        _filesTab.FolderPath = table.Rows[0]["Header_Icons$Folder Path"].ToString();
                        if (string.IsNullOrEmpty(_filesTab.FolderPath))
                        {
                            _extentTest.Log(Status.Fail, "Folder Path is missing");
                            message += "Folder Path&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "The given folder path is" + _filesTab.FolderPath);
                        }

                        _filesTab.ReplaceOrRename = table.Rows[0]["Header_Icons$Replace Or Rename"].ToString();
                        if (string.IsNullOrEmpty(_filesTab.ReplaceOrRename))
                        {
                            _extentTest.Log(Status.Fail, "Replace or Rename  is missing");
                            message += "Replace Or Rename&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "The selected file is" + _filesTab.ReplaceOrRename);
                        }
                    }
                }

                if (tblcolumns.Contains("Header_Icons$Delete The Files"))
                {
                    _filesTab.DeleteFile = table.Rows[0]["Header_Icons$Delete The Files"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.DeleteFile ? "Delete the files for the files tab is true" : "Delete the files for the files tab is false");
                }

                if (tblcolumns.Contains("Header_Icons$Assay Kit Verification"))
                {
                    _filesTab.AssayKitVerification = table.Rows[0]["Header_Icons$Assay Kit Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _filesTab.AssayKitVerification ? "Assay Kit Verification for the files tab is true" : "Assay kit Verification for the files tab is false");

                    if (_filesTab.AssayKitVerification)
                    {
                        _filesTab.CatNumber = table.Rows[0]["Header_Icons$Cat Number"].ToString();
                        if (string.IsNullOrEmpty(_filesTab.CatNumber))
                        {
                            _extentTest.Log(Status.Fail, "Cat Number is missing");
                            message += "Cat Number&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "The given cat number is" + _filesTab.CatNumber);
                        }

                        _filesTab.LotNumber = table.Rows[0]["Header_Icons$Lot Number"].ToString();
                        if (string.IsNullOrEmpty(_filesTab.LotNumber))
                        {
                            _extentTest.Log(Status.Fail, "Lot Number is missing");
                            message += "Lot Number&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "The given lot number is" + _filesTab.LotNumber);
                        }

                        _filesTab.SWID = table.Rows[0]["Header_Icons$SWID"].ToString();
                        if (string.IsNullOrEmpty(_filesTab.SWID))
                        {
                            _extentTest.Log(Status.Fail, "SWID is missing");
                            message += "SWID&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "The given SWID number is" + _filesTab.SWID);
                        }
                    }

                    if (tblcolumns.Contains("Header_Icons$Export Files Verification"))
                    {
                        _filesTab.ExportFilesVerification = table.Rows[0]["Header_Icons$Export Files Verification"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _filesTab.ExportFilesVerification ? "Export files verification for the files tab is true" : "Export files verification for the files tab is false");
                    }

                    if (tblcolumns.Contains("Header_Icons$Send To Verification"))
                    {
                        _filesTab.SendToVerfication = table.Rows[0]["Header_Icons$Send To Verification"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _filesTab.SendToVerfication ? "Send to verification for the files tab is true" : "Send to verification for the files tab is false");

                        if (_filesTab.SendToVerfication)
                        {
                            _filesTab.FirstMailRecepient = table.Rows[0]["Header_Icons$First Mail Recepient"].ToString();
                            if (string.IsNullOrEmpty(_filesTab.FirstMailRecepient))
                            {
                                _extentTest.Log(Status.Fail, "Email Id is missing");
                                message += "FIrst Mail Recepient &";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The given mail id is" + _filesTab.FirstMailRecepient);
                            }
                        }
                    }

                    if (tblcolumns.Contains("Header_Icons$Rename Verification"))
                    {
                        _filesTab.RenameVerification = table.Rows[0]["Header_Icons$Rename Verification"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _filesTab.RenameVerification ? "Rename verification for the files tab is true" : "Rename verification for the files tab is false");
                    }

                    if (tblcolumns.Contains("Header_Icons$Add Favorite"))
                    {
                        _filesTab.AddFavorite = table.Rows[0]["Header_Icons$Add Favorite"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _filesTab.AddFavorite ? "Add Favorite verification for the files tab is true" : "Add Favorite verification for the files tab is false");
                    }

                    if (tblcolumns.Contains("Header_Icons$Add Category"))
                    {
                        _filesTab.AddCategory = table.Rows[0]["Header_Icons$Add Category"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _filesTab.AddCategory ? "Add category for the files is true" : "Add category for the files tab is false");

                        if (_filesTab.AddCategory)
                        {
                            _filesTab.AddCategoryName = table.Rows[0]["Header_Icons$Add Category Name"].ToString();
                            if (string.IsNullOrEmpty(_filesTab.AddCategoryName))
                            {
                                _extentTest.Log(Status.Fail, "Category Name is missing");
                                message += "Category Name&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The given Category Name is" + _filesTab.AddCategoryName);
                            }

                            _filesTab.EditCategoryName = table.Rows[0]["Header_Icons$Edited Category Name"].ToString();
                            if (string.IsNullOrEmpty(_filesTab.EditCategoryName))
                            {
                                _extentTest.Log(Status.Fail, "Edited Category Name is missing");
                                message += "Edited Category Name&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The given edited Category Name is" + _filesTab.EditCategoryName);
                            }
                        }
                    }
                }
                #endregion
            }
            else if (sheetName == "Workflow-8")
            {
                #region TestId -1

                _fileUploadOrExistingFileData.IsFileUploadRequired = table.Rows[0]["Upload_File$IsFileUploadRequired"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.IsFileUploadRequired)
                {
                    _extentTest.Log(Status.Pass, "FileUpload required status is true");

                    _fileUploadOrExistingFileData.FileUploadPath = table.Rows[0]["Upload_File$FileUploadPath"].ToString();
                    if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.FileUploadPath))
                    {
                        _extentTest.Log(Status.Fail, "FileUploadPath  is empty");
                        message += "FileUploadPath&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Fileupload path is present");
                    }
                    _fileUploadOrExistingFileData.FileName = table.Rows[0]["Upload_File$FileName"].ToString();
                    if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.FileName))
                    {
                        _extentTest.Log(Status.Fail, "FileName field is empty");
                        message += "FileName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "FileName is present - " + _fileUploadOrExistingFileData.FileName);
                    }
                    _fileUploadOrExistingFileData.FileExtension = table.Rows[0]["Upload_File$FileExtension"].ToString();
                    if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.FileExtension))
                    {
                        _extentTest.Log(Status.Fail, "FileExtension field is empty");
                        message += "FileExtension&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "FileExtension is present - " + _fileUploadOrExistingFileData.FileExtension);
                    }
                }
                else
                {
                    _extentTest.Log(Status.Pass, "FileUpload required status is false");
                }

                _fileUploadOrExistingFileData.OpenExistingFile = table.Rows[0]["Upload_File$OpenExistingFile"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.OpenExistingFile)
                {
                    _extentTest.Log(Status.Pass, "Existing file name status is true");

                    _fileUploadOrExistingFileData.FileName = table.Rows[0]["Upload_File$FileName"].ToString();
                    if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.FileName))
                    {
                        _extentTest.Log(Status.Fail, "FileName status is false");
                        message += "FileName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "FileName status is true");
                    }
                    _fileUploadOrExistingFileData.FileExtension = table.Rows[0]["Upload_File$FileExtension"].ToString();
                    if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.FileExtension))
                    {
                        _extentTest.Log(Status.Fail, "File Extension field is empty");
                        message += "File Extension&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "FileExtension is present - " + _fileUploadOrExistingFileData.FileExtension);
                    }
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Existing file name status is false");
                }

                if (_fileUploadOrExistingFileData.IsFileUploadRequired && _fileUploadOrExistingFileData.OpenExistingFile)
                    _extentTest.Log(Status.Fail, "Both FileUpload required status and Open Existing File status is true");

                if (!_fileUploadOrExistingFileData.IsFileUploadRequired && !_fileUploadOrExistingFileData.OpenExistingFile)
                    _extentTest.Log(Status.Fail, "Both FileUpload required status and Open Existing File status is false");

                _fileUploadOrExistingFileData.IsTitrationFile = table.Rows[0]["Upload_File$IsTitrationFile"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.IsTitrationFile)
                {
                    _extentTest.Log(Status.Pass, "File Titration type is " + _fileUploadOrExistingFileData.IsTitrationFile);
                }
                else
                {
                    _extentTest.Log(Status.Pass, "File Titration type is " + _fileUploadOrExistingFileData.IsTitrationFile);
                }

                _fileUploadOrExistingFileData.IsNormalized = table.Rows[0]["Upload_File$IsNormalized"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.IsNormalized)
                {
                    _extentTest.Log(Status.Pass, "File Normalization status is " + _fileUploadOrExistingFileData.IsNormalized);
                }
                else
                {
                    _extentTest.Log(Status.Pass, "File Normalization status is " + _fileUploadOrExistingFileData.IsNormalized);
                }

                var filetype = table.Rows[0]["Upload_File$FileType"].ToString();
                if (String.IsNullOrEmpty(filetype))
                {
                    _extentTest.Log(Status.Fail, "File Type is empty");
                    message += "File Type&";
                }
                else
                {
                    _fileUploadOrExistingFileData.FileType = filetype == "Xfe24" ? FileType.Xfe24 : filetype == "Xfe96" ? FileType.Xfe96 : filetype == "Xfp" ? FileType.Xfp : filetype == "XfHsMini" ? FileType.XfHsMini : filetype == "XFPro" ? FileType.XFPro : FileType.XFPro;

                    _extentTest.Log(Status.Pass, "File Type is " + filetype);
                }

                //ToD0:  Need to Log all the files.
                _fileUploadOrExistingFileData.SelectedWidgets = new List<WidgetTypes>();
                if (table.Rows[0]["Upload_File$XF Cell Energy Phenotype"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.XfCellEnergyPhenotype);
                }
                if (table.Rows[0]["Upload_File$Metabolic Potential OCR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.MetabolicPotentialOcr);
                }
                if (table.Rows[0]["Upload_File$Metabolic Potential ECAR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.MetabolicPotentialEcar);
                }
                if (table.Rows[0]["Upload_File$Baseline OCR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.BaselineOcr);
                }
                if (table.Rows[0]["Upload_File$Baseline ECAR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.BaselineEcar);
                }
                if (table.Rows[0]["Upload_File$Stressed OCR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.StressedOcr);
                }
                if (table.Rows[0]["Upload_File$Stressed ECAR"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.StressedEcar);
                }
                if (table.Rows[0]["Upload_File$Data Table"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.DataTable);
                }

                #endregion

                #region TestId -2

                if (tblcolumns.Contains("Layout_Verification$Layout Verification"))
                {
                    _workFlow8.LayoutVerification = table.Rows[0]["Layout_Verification$Layout Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.LayoutVerification ? "Analysis page layout verification is true" : "Analysis page layout verification is false");
                }

                #endregion

                #region TestId -3

                _workFlow8.CellEnergyPhenotype = new WidgetItems();
                _workFlow8.CellEnergyPhenotype.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("CellEnergy_Phenotype$Normalization"))
                {
                    _workFlow8.CellEnergyPhenotype.Normalization = table.Rows[0]["CellEnergy_Phenotype$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow8.CellEnergyPhenotype.Normalization ? "Normalization for CellEnergyPhenotype is true" : "Normalization for CellEnergyPhenotype is false");
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$Error Format"))
                {
                    _workFlow8.CellEnergyPhenotype.ErrorFormat = table.Rows[0]["CellEnergy_Phenotype$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for CellEnergyPhenotype  is " + _workFlow8.CellEnergyPhenotype.ErrorFormat);
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$IsExportRequired"))
                {
                    _workFlow8.CellEnergyPhenotype.IsExportRequired = table.Rows[0]["CellEnergy_Phenotype$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -4

                _workFlow8.MetabolicPotentialOCR = new WidgetItems();
                _workFlow8.MetabolicPotentialOCR.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("MetabolicPotential_OCR$Display"))
                {
                    _workFlow8.MetabolicPotentialOCR.Display = table.Rows[0]["MetabolicPotential_OCR$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Display mode for MetabolicPotential_OCR  is " + _workFlow8.MetabolicPotentialOCR.Display);
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$Error Format"))
                {
                    _workFlow8.MetabolicPotentialOCR.ErrorFormat = table.Rows[0]["MetabolicPotential_OCR$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for MetabolicPotentialOCR  is " + _workFlow8.MetabolicPotentialOCR.ErrorFormat);
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$Sort By"))
                {
                    _workFlow8.MetabolicPotentialOCR.SortBy = table.Rows[0]["MetabolicPotential_OCR$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for MetabolicPotentialOCR  is " + _workFlow8.MetabolicPotentialOCR.SortBy);
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["MetabolicPotential_OCR$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for MetabolicPotentialOCR value is " + GraphUnits);
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$GraphSettingsRequired"))
                {
                    _workFlow8.MetabolicPotentialOCR.GraphSettingsVerify = table.Rows[0]["MetabolicPotential_OCR$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.MetabolicPotentialOCR.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("MetabolicPotential_OCR$Remove Y AutoScale"))
                        {
                            _workFlow8.MetabolicPotentialOCR.GraphSettings.RemoveYAutoScale = table.Rows[0]["MetabolicPotential_OCR$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for MetabolicPotentialOCR graph is true" : "Remove Y AutoScale in GraphSettings for MetabolicPotentialOCR graph is false");
                        }

                        if (tblcolumns.Contains("MetabolicPotential_OCR$Remove ZeroLine"))
                        {
                            _workFlow8.MetabolicPotentialOCR.GraphSettings.RemoveZeroLine = table.Rows[0]["MetabolicPotential_OCR$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for MetabolicPotentialOCR graph is true" : "Remove ZeroLine in GraphSettings for MetabolicPotentialOCR graph is false");
                        }

                        if (tblcolumns.Contains("MetabolicPotential_OCR$Remove Zoom"))
                        {
                            _workFlow8.MetabolicPotentialOCR.GraphSettings.RemoveZoom = table.Rows[0]["MetabolicPotential_OCR$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for MetabolicPotentialOCR graph is true" : "Remove Zoom in GraphSettings for MetabolicPotentialOCR graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.MetabolicPotentialOCR.CheckNormalizationWithPlateMap = table.Rows[0]["MetabolicPotential_OCR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.MetabolicPotentialOCR.CheckNormalizationWithPlateMap = table.Rows[0]["MetabolicPotential_OCR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$PlateMap Sync to View"))
                {
                    _workFlow8.MetabolicPotentialOCR.PlateMapSynctoView = table.Rows[0]["MetabolicPotential_OCR$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$GraphSettings Sync to View"))
                {
                    _workFlow8.MetabolicPotentialOCR.GraphSettings.SynctoView = table.Rows[0]["MetabolicPotential_OCR$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_OCR$IsExportRequired"))
                {
                    _workFlow8.MetabolicPotentialOCR.IsExportRequired = table.Rows[0]["MetabolicPotential_OCR$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialOCR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -5

                _workFlow8.MetabolicPotentialECAR = new WidgetItems();
                _workFlow8.MetabolicPotentialECAR.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("MetabolicPotential_ECAR$Display"))
                {
                    _workFlow8.MetabolicPotentialECAR.Display = table.Rows[0]["MetabolicPotential_ECAR$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Display mode for MetabolicPotentialECAR  is " + _workFlow8.MetabolicPotentialECAR.Display);
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$Error Format"))
                {
                    _workFlow8.MetabolicPotentialECAR.ErrorFormat = table.Rows[0]["MetabolicPotential_ECAR$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for MetabolicPotentialECAR  is " + _workFlow8.MetabolicPotentialECAR.ErrorFormat);
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$Sort By"))
                {
                    _workFlow8.MetabolicPotentialECAR.SortBy = table.Rows[0]["MetabolicPotential_ECAR$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for MetabolicPotentialECAR  is " + _workFlow8.MetabolicPotentialECAR.SortBy);
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["MetabolicPotential_ECAR$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for MetabolicPotentialECAR value is " + GraphUnits);
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$GraphSettingsRequired"))
                {
                    _workFlow8.MetabolicPotentialECAR.GraphSettingsVerify = table.Rows[0]["MetabolicPotential_ECAR$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.MetabolicPotentialECAR.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("MetabolicPotential_ECAR$Remove Y AutoScale"))
                        {
                            _workFlow8.MetabolicPotentialECAR.GraphSettings.RemoveYAutoScale = table.Rows[0]["MetabolicPotential_ECAR$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for MetabolicPotentialECAR graph is true" : "Remove Y AutoScale in GraphSettings for MetabolicPotentialECAR graph is false");
                        }

                        if (tblcolumns.Contains("MetabolicPotential_ECAR$Remove ZeroLine"))
                        {
                            _workFlow8.MetabolicPotentialECAR.GraphSettings.RemoveZeroLine = table.Rows[0]["MetabolicPotential_ECAR$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for MetabolicPotentialECAR graph is true" : "Remove ZeroLine in GraphSettings for MetabolicPotentialECAR graph is false");
                        }

                        if (tblcolumns.Contains("MetabolicPotential_ECAR$Remove Zoom"))
                        {
                            _workFlow8.MetabolicPotentialECAR.GraphSettings.RemoveZoom = table.Rows[0]["MetabolicPotential_ECAR$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for MetabolicPotentialECAR graph is true" : "Remove Zoom in GraphSettings for MetabolicPotentialECAR graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.MetabolicPotentialECAR.CheckNormalizationWithPlateMap = table.Rows[0]["MetabolicPotential_ECAR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.MetabolicPotentialECAR.CheckNormalizationWithPlateMap = table.Rows[0]["MetabolicPotential_ECAR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$PlateMap Sync to View"))
                {
                    _workFlow8.MetabolicPotentialECAR.PlateMapSynctoView = table.Rows[0]["MetabolicPotential_ECAR$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$GraphSettings Sync to View"))
                {
                    _workFlow8.MetabolicPotentialECAR.GraphSettings.SynctoView = table.Rows[0]["MetabolicPotential_ECAR$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("MetabolicPotential_ECAR$IsExportRequired"))
                {
                    _workFlow8.MetabolicPotentialECAR.IsExportRequired = table.Rows[0]["MetabolicPotential_ECARR$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.MetabolicPotentialECAR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -6

                _workFlow8.BaselineOCR = new WidgetItems();
                _workFlow8.BaselineOCR.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Baseline_OCR$Display"))
                {
                    _workFlow8.BaselineOCR.Display = table.Rows[0]["Baseline_OCR$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Display mode for Baseline_OCR  is " + _workFlow8.BaselineOCR.Display);
                }

                if (tblcolumns.Contains("Baseline_OCR$Normalization"))
                {
                    _workFlow8.BaselineOCR.Normalization = table.Rows[0]["Baseline_OCR$Normalization"].ToString().ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.Normalization ? "Normalization for Baseline_OCR is true" : "Normalization for Baseline_OCR is false");
                }

                if (tblcolumns.Contains("Baseline_OCR$Error Format"))
                {
                    _workFlow8.BaselineOCR.ErrorFormat = table.Rows[0]["Baseline_OCR$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for BaselineOCR  is " + _workFlow8.BaselineOCR.ErrorFormat);
                }

                if (tblcolumns.Contains("Baseline_OCR$Sort By"))
                {
                    _workFlow8.BaselineOCR.SortBy = table.Rows[0]["Baseline_OCR$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for BaselineOCR  is " + _workFlow8.BaselineOCR.SortBy);
                }

                if (tblcolumns.Contains("Baseline_OCR$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["Baseline_OCR$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for BaselineOCR value is " + GraphUnits);
                }

                if (tblcolumns.Contains("Baseline_OCR$GraphSettingsRequired"))
                {
                    _workFlow8.BaselineOCR.GraphSettingsVerify = table.Rows[0]["Baseline_OCR$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.BaselineOCR.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("Baseline_OCR$Remove Y AutoScale"))
                        {
                            _workFlow8.BaselineOCR.GraphSettings.RemoveYAutoScale = table.Rows[0]["Baseline_OCR$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for BaselineOCR graph is true" : "Remove Y AutoScale in GraphSettings for BaselineOCR graph is false");
                        }

                        if (tblcolumns.Contains("Baseline_OCR$Remove ZeroLine"))
                        {
                            _workFlow8.BaselineOCR.GraphSettings.RemoveZeroLine = table.Rows[0]["Baseline_OCR$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for BaselineOCR graph is true" : "Remove ZeroLine in GraphSettings for BaselineOCR graph is false");
                        }

                        if (tblcolumns.Contains("Baseline_OCR$Remove Zoom"))
                        {
                            _workFlow8.BaselineOCR.GraphSettings.RemoveZoom = table.Rows[0]["Baseline_OCR$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for BaselineOCR graph is true" : "Remove Zoom in GraphSettings for BaselineOCR graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("Baseline_OCR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.BaselineOCR.CheckNormalizationWithPlateMap = table.Rows[0]["Baseline_OCR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_OCR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.BaselineOCR.CheckNormalizationWithPlateMap = table.Rows[0]["Baseline_OCR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_OCR$PlateMap Sync to View"))
                {
                    _workFlow8.BaselineOCR.PlateMapSynctoView = table.Rows[0]["Baseline_OCR$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_OCR$GraphSettings Sync to View"))
                {
                    _workFlow8.BaselineOCR.GraphSettings.SynctoView = table.Rows[0]["Baseline_OCR$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_OCR$IsExportRequired"))
                {
                    _workFlow8.BaselineOCR.IsExportRequired = table.Rows[0]["Baseline_OCR$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineOCR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -7

                _workFlow8.BaselineECAR = new WidgetItems();
                _workFlow8.BaselineECAR.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Baseline_ECAR$Display"))
                {
                    _workFlow8.BaselineECAR.Display = table.Rows[0]["Baseline_ECAR$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Display mode for Baseline_ECAR  is " + _workFlow8.BaselineECAR.Display);
                }

                if (tblcolumns.Contains("Baseline_ECAR$Normalization"))
                {
                    _workFlow8.BaselineECAR.Normalization = table.Rows[0]["Baseline_ECAR$Normalization"].ToString().ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.Normalization ? "Normalization for BaselineECAR is true" : "Normalization for BaselineECAR is false");
                }

                if (tblcolumns.Contains("Baseline_ECAR$Error Format"))
                {
                    _workFlow8.BaselineECAR.ErrorFormat = table.Rows[0]["Baseline_ECAR$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for BaselineECAR  is " + _workFlow8.BaselineECAR.ErrorFormat);
                }

                if (tblcolumns.Contains("Baseline_ECAR$Sort By"))
                {
                    _workFlow8.BaselineECAR.SortBy = table.Rows[0]["Baseline_ECAR$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for BaselineECAR  is " + _workFlow8.BaselineECAR.SortBy);
                }

                if (tblcolumns.Contains("Baseline_ECAR$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["Baseline_ECAR$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for BaselineECAR value is " + GraphUnits);
                }

                if (tblcolumns.Contains("Baseline_ECAR$GraphSettingsRequired"))
                {
                    _workFlow8.BaselineECAR.GraphSettingsVerify = table.Rows[0]["Baseline_ECAR$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.BaselineECAR.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("Baseline_ECAR$Remove Y AutoScale"))
                        {
                            _workFlow8.BaselineECAR.GraphSettings.RemoveYAutoScale = table.Rows[0]["Baseline_ECAR$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for BaselineECAR graph is true" : "Remove Y AutoScale in GraphSettings for BaselineECAR graph is false");
                        }

                        if (tblcolumns.Contains("Baseline_ECAR$Remove ZeroLine"))
                        {
                            _workFlow8.BaselineECAR.GraphSettings.RemoveZeroLine = table.Rows[0]["Baseline_ECAR$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for BaselineECAR graph is true" : "Remove ZeroLine in GraphSettings for BaselineECAR graph is false");
                        }

                        if (tblcolumns.Contains("Baseline_ECAR$Remove Zoom"))
                        {
                            _workFlow8.BaselineECAR.GraphSettings.RemoveZoom = table.Rows[0]["Baseline_ECAR$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for BaselineECAR graph is true" : "Remove Zoom in GraphSettings for BaselineECAR graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("Baseline_ECAR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.BaselineECAR.CheckNormalizationWithPlateMap = table.Rows[0]["Baseline_ECAR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_ECAR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.BaselineECAR.CheckNormalizationWithPlateMap = table.Rows[0]["Baseline_ECAR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_ECAR$PlateMap Sync to View"))
                {
                    _workFlow8.BaselineECAR.PlateMapSynctoView = table.Rows[0]["Baseline_ECAR$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_ECAR$GraphSettings Sync to View"))
                {
                    _workFlow8.BaselineECAR.GraphSettings.SynctoView = table.Rows[0]["Baseline_ECAR$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Baseline_ECAR$IsExportRequired"))
                {
                    _workFlow8.BaselineECAR.IsExportRequired = table.Rows[0]["Baseline_ECAR$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.BaselineECAR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -8

                _workFlow8.StressedOCR = new WidgetItems();
                _workFlow8.StressedOCR.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Stressed_OCR$Display"))
                {
                    _workFlow8.StressedOCR.Display = table.Rows[0]["Stressed_OCR$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Display mode for Stressed_OCR  is " + _workFlow8.StressedOCR.Display);
                }

                if (tblcolumns.Contains("Stressed_OCR$Normalization"))
                {
                    _workFlow8.StressedOCR.Normalization = table.Rows[0]["Stressed_OCR$Normalization"].ToString().ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.Normalization ? "Normalization for Stressed_OCR is true" : "Normalization for Stressed_OCR is false");
                }

                if (tblcolumns.Contains("Stressed_OCR$Error Format"))
                {
                    _workFlow8.StressedOCR.ErrorFormat = table.Rows[0]["Stressed_OCR$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for StressedOCR  is " + _workFlow8.StressedOCR.ErrorFormat);
                }

                if (tblcolumns.Contains("Stressed_OCR$Sort By"))
                {
                    _workFlow8.StressedOCR.SortBy = table.Rows[0]["Stressed_OCR$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for StressedOCR  is " + _workFlow8.StressedOCR.SortBy);
                }

                if (tblcolumns.Contains("Stressed_OCR$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["Stressed_OCR$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for StressedOCR value is " + GraphUnits);
                }

                if (tblcolumns.Contains("Stressed_OCR$GraphSettingsRequired"))
                {
                    _workFlow8.StressedOCR.GraphSettingsVerify = table.Rows[0]["Stressed_OCR$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.StressedOCR.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("Stressed_OCR$Remove Y AutoScale"))
                        {
                            _workFlow8.StressedOCR.GraphSettings.RemoveYAutoScale = table.Rows[0]["Stressed_OCR$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for StressedOCR graph is true" : "Remove Y AutoScale in GraphSettings for StressedOCR graph is false");
                        }

                        if (tblcolumns.Contains("Stressed_OCR$Remove ZeroLine"))
                        {
                            _workFlow8.StressedOCR.GraphSettings.RemoveZeroLine = table.Rows[0]["Stressed_OCR$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for StressedOCR graph is true" : "Remove ZeroLine in GraphSettings for StressedOCR graph is false");
                        }

                        if (tblcolumns.Contains("Stressed_OCR$Remove Zoom"))
                        {
                            _workFlow8.StressedOCR.GraphSettings.RemoveZoom = table.Rows[0]["Stressed_OCR$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for StressedOCR graph is true" : "Remove Zoom in GraphSettings for StressedOCR graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("Stressed_OCR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.StressedOCR.CheckNormalizationWithPlateMap = table.Rows[0]["Stressed_OCR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_OCR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.StressedOCR.CheckNormalizationWithPlateMap = table.Rows[0]["Stressed_OCR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_OCR$PlateMap Sync to View"))
                {
                    _workFlow8.StressedOCR.PlateMapSynctoView = table.Rows[0]["Stressed_OCR$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_OCR$GraphSettings Sync to View"))
                {
                    _workFlow8.StressedOCR.GraphSettings.SynctoView = table.Rows[0]["Stressed_OCR$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_OCR$IsExportRequired"))
                {
                    _workFlow8.StressedOCR.IsExportRequired = table.Rows[0]["Stressed_OCR$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedOCR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -9

                _workFlow8.StressedECAR = new WidgetItems();
                _workFlow8.StressedECAR.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Stressed_ECAR$Display"))
                {
                    _workFlow8.StressedECAR.Display = table.Rows[0]["Stressed_ECAR$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Display mode for Stressed_ECAR  is " + _workFlow8.StressedECAR.Display);
                }

                if (tblcolumns.Contains("Stressed_ECAR$Normalization"))
                {
                    _workFlow8.StressedECAR.Normalization = table.Rows[0]["Stressed_ECAR$Normalization"].ToString().ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.Normalization ? "Normalization for Stressed_ECAR is true" : "Normalization for Stressed_ECAR is false");
                }

                if (tblcolumns.Contains("Stressed_ECAR$Error Format"))
                {
                    _workFlow8.StressedECAR.ErrorFormat = table.Rows[0]["Stressed_ECAR$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for StressedECAR  is " + _workFlow8.StressedECAR.ErrorFormat);
                }

                if (tblcolumns.Contains("Stressed_ECAR$Sort By"))
                {
                    _workFlow8.StressedECAR.SortBy = table.Rows[0]["Stressed_ECAR$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for StressedECAR  is " + _workFlow8.StressedECAR.SortBy);
                }

                if (tblcolumns.Contains("Stressed_ECAR$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["Stressed_ECAR$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for StressedECAR value is " + GraphUnits);
                }

                if (tblcolumns.Contains("Stressed_ECAR$GraphSettingsRequired"))
                {
                    _workFlow8.StressedECAR.GraphSettingsVerify = table.Rows[0]["Stressed_ECAR$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.StressedECAR.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("Stressed_ECAR$Remove Y AutoScale"))
                        {
                            _workFlow8.StressedECAR.GraphSettings.RemoveYAutoScale = table.Rows[0]["Stressed_ECAR$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for StressedECAR graph is true" : "Remove Y AutoScale in GraphSettings for StressedECAR graph is false");
                        }

                        if (tblcolumns.Contains("Stressed_ECAR$Remove ZeroLine"))
                        {
                            _workFlow8.StressedECAR.GraphSettings.RemoveZeroLine = table.Rows[0]["Stressed_ECAR$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for StressedECAR graph is true" : "Remove ZeroLine in GraphSettings for StressedECAR graph is false");
                        }

                        if (tblcolumns.Contains("Stressed_ECAR$Remove Zoom"))
                        {
                            _workFlow8.StressedECAR.GraphSettings.RemoveZoom = table.Rows[0]["Stressed_ECAR$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for StressedECAR graph is true" : "Remove Zoom in GraphSettings for StressedECAR graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("Stressed_ECAR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.StressedECAR.CheckNormalizationWithPlateMap = table.Rows[0]["Stressed_ECAR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_ECAR$CheckNormalizationWithPlateMap"))
                {
                    _workFlow8.StressedECAR.CheckNormalizationWithPlateMap = table.Rows[0]["Stressed_ECAR$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_ECAR$PlateMap Sync to View"))
                {
                    _workFlow8.StressedECAR.PlateMapSynctoView = table.Rows[0]["Stressed_ECAR$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_ECAR$GraphSettings Sync to View"))
                {
                    _workFlow8.StressedECAR.GraphSettings.SynctoView = table.Rows[0]["Stressed_ECAR$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Stressed_ECAR$IsExportRequired"))
                {
                    _workFlow8.StressedECAR.IsExportRequired = table.Rows[0]["Stressed_ECAR$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.StressedECAR.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion 
            }

            //Final check if any message contains the list
            if (string.IsNullOrEmpty(message))
            {
                _extentTest.Log(Status.Pass, "All required Data's are available! and the verified sheet name is " + sheetName);
                return true;
            }
            else
            {
                message = message.Replace("&", ", ");
                _extentTest.Log(Status.Fail, message + " are required details and it is missing from excel sheet and sheet name is " + sheetName);
                return false;
            }
        }
    }
}

