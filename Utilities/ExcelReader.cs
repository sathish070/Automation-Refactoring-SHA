using System;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Controls;
using AventStack.ExtentReports;
using DataTable = System.Data.DataTable;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace SHAProject.Utilities
{
    public class ExcelReader
    {    
        public ExtentTest? _extentTest;
        public ExtentTest? _extentTestNode;
        public ExtentReports? _extentReport;
        private readonly LoginData _loginData;
        public readonly FilesTabData _filesTab;
        private readonly NormalizationData _normalization;
        private readonly WorkFlow1Data _workFlow1;
        private readonly WorkFlow6Data _workFlow6;
        private readonly WorkFlow8Data _workFlow8;
        private readonly string _currentBuildPath;
        private readonly CurrentBrowser _currentBrowser;
        public static List<string> testidList = new List<string>();
        private readonly FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public ExcelReader(LoginData loginData, FileUploadOrExistingFileData fileUploadOrExistingFileData, NormalizationData normalization, WorkFlow1Data workFlow1,WorkFlow6Data workFlow6 , WorkFlow8Data workFlow8,
            string currentBuildPath, CurrentBrowser currentBrowser, ExtentTest extentTest,FilesTabData filesTab)
        {
            _loginData = loginData;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _normalization = normalization;
            _workFlow1 = workFlow1;
            _workFlow6 = workFlow6;
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
            else if (sheetName == "Workflow-1")
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
                    _workFlow1.AnalysisLayoutVerification = table.Rows[0]["Layout_Verification$Layout Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.AnalysisLayoutVerification ? "Analysis page layout verification is true" : "Analysis page layout verification is false");
                }
                #endregion

                #region TestId -4

                if (tblcolumns.Contains("Navg_Bar_Icons$ExportViewOption"))
                {
                    _workFlow1.ExportViewOption = table.Rows[0]["Navg_Bar_Icons$ExportViewOption"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.ExportViewOption))
                    {
                        _extentTest.Log(Status.Fail, "ExportViewOption is missing");
                        message += "ExportViewOption&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "ExportViewOption is " + _workFlow1.ExportViewOption);
                    }
                }

                if (tblcolumns.Contains("Navg_Bar_Icons$DeleteWidgetRequired"))
                {
                    _workFlow1.DeleteWidgetRequired = table.Rows[0]["Navg_Bar_Icons$DeleteWidgetRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.DeleteWidgetRequired ? "Delete Widget Required is true" : "Delete Widget Required is false");
                    if (_workFlow1.DeleteWidgetRequired)
                    {
                        if (tblcolumns.Contains("Navg_Bar_Icons$DeleteWidgetName"))
                        {
                            var widgetName = table.Rows[0]["Navg_Bar_Icons$DeleteWidgetName"].ToString();
                            if (string.IsNullOrEmpty(widgetName))
                            {
                                _extentTest.Log(Status.Fail, "DeleteWidgetName is missing");
                                message += "DeleteWidgetName&";
                            }
                            else
                            {
                                _workFlow1.DeleteWidgetName = widgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                               widgetName == "Bar Chart" ? WidgetTypes.BarChart : widgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                               widgetName == "Heat Map" ? WidgetTypes.KineticGraph : WidgetTypes.KineticGraph;

                                _extentTest.Log(Status.Pass, "DeleteWidgetName is " + _workFlow1.DeleteWidgetName);
                            }
                        }
                    }
                }
                if (tblcolumns.Contains("Navg_Bar_Icons$AddWidgetRequired"))
                {
                    _workFlow1.AddWidgetRequired = table.Rows[0]["Navg_Bar_Icons$AddWidgetRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.AddWidgetRequired ? "Add Widget Required is true" : "Add Widget Required is false");
                    if (_workFlow1.AddWidgetRequired)
                    {
                        //Kinetic Graph - Ocr, Kinetic Graph - Ecar, Kinetic Graph - Per, Bar Chart, Energetic Map, Heat Map
                        if (tblcolumns.Contains("Navg_Bar_Icons$AddWidgetName"))
                        {
                            var widgetName = table.Rows[0]["Navg_Bar_Icons$AddWidgetName"].ToString();
                            if (string.IsNullOrEmpty(widgetName))
                            {
                                _extentTest.Log(Status.Fail, "AddWidgetName is missing");
                                message += "AddWidgetName&";
                            }
                            else
                            {
                                _workFlow1.AddWidgetName = widgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                               widgetName == "Bar Chart" ? WidgetTypes.BarChart : widgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                               widgetName == "Heat Map" ? WidgetTypes.KineticGraph : WidgetTypes.KineticGraph;

                                _extentTest.Log(Status.Pass, "AddWidgetName is " + _workFlow1.AddWidgetName);
                            }
                        }
                    }
                }
                #endregion

                #region TestId -5

                if (tblcolumns.Contains("Normalization_Icon$Normalization Verification"))
                {
                    _workFlow1.NormalizationVerification = table.Rows[0]["Normalization_Icon$Normalization Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.NormalizationVerification ? "Normalization Verification is true" : "Normalization Verification is false");

                    _workFlow1.NormalizedFileName = table.Rows[0]["Normalization_Icon$Normalization File"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.NormalizedFileName))
                    {
                        _extentTest.Log(Status.Fail, "Normalization inbuilt file name is Empty");
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalization inbuilt file name is "+ _workFlow1.NormalizedFileName);
                    }
                    ReadDataFromExcel("Normalization");
                    //if (_workFlow1.Normalization)
                    //{
                    //    if (tblcolumns.Contains("Normalization_Icon$Normalization Label"))
                    //    {
                    //        _workFlow1.NormalizationLabel = table.Rows[0]["Normalization_Icon$Normalization Label"].ToString();
                    //        if (string.IsNullOrEmpty(_workFlow1.NormalizationLabel))
                    //            _extentTest.Log(Status.Fail, "Normalization label cannot be empty");
                    //        else
                    //            _extentTest.Log(Status.Pass, "NormalizationUnits for Barchart is " + _workFlow1.NormalizationLabel);
                    //    }

                    //    if (tblcolumns.Contains("Normalization_Icon$Scale Factor"))
                    //    {
                    //        var scaleFactorValue = table.Rows[0]["Normalization_Icon$Scale Factor"] == DBNull.Value ? null : table.Rows[0]["Normalization_Icon$Scale Factor"];
                    //        if (string.IsNullOrEmpty(scaleFactorValue.ToString()))
                    //            _extentTest.Log(Status.Pass, "The default scale factor value is 1.");
                    //        else
                    //        {
                    //            _workFlow1.NormalizationScaleFactor = scaleFactorValue.ToString();
                    //            _extentTest.Log(Status.Pass, "ScaleFactor for Barchart is " + _workFlow1.NormalizationScaleFactor);
                    //        }
                    //    }

                    //    if (tblcolumns.Contains("Normalization_Icon$Apply to all widgets"))
                    //    {
                    //        _workFlow1.ApplyToAllWidgets = table.Rows[0]["Normalization_Icon$Apply to all widgets"].ToString() == "Yes";
                    //        _extentTest.Log(Status.Pass, "ScaleFactor for Barchart is " + _workFlow1.ApplyToAllWidgets);
                    //    }
                    //    ReadDataFromExcel("Normalization");
                    //}
                }

                if (tblcolumns.Contains("Normalization_Icon$Apply to all widgets"))
                {
                    _workFlow1.ApplyToAllWidgets = table.Rows[0]["Normalization_Icon$Apply to all widgets"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, "ScaleFactor for Barchart is " + _workFlow1.ApplyToAllWidgets);
                }

                #endregion

                #region TestId -6

                if (tblcolumns.Contains("Modify_Assay$ModifyAssay Verification"))
                {
                    _workFlow1.ModifyAssay = table.Rows[0]["Modify_Assay$ModifyAssay Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.ModifyAssay ? "ModifyAssay Verification is true" : "ModifyAssay Verification is false");
                }

                if (tblcolumns.Contains("Modify_Assay$Add Group Name"))
                {
                    _workFlow1.AddGroupName = table.Rows[0]["Modify_Assay$Add Group Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.AddGroupName))
                    {
                        _extentTest.Log(Status.Fail, " Group Name is empty :" + _workFlow1.AddGroupName);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Add group Name is : " + _workFlow1.AddGroupName);
                    }
                }

                if (tblcolumns.Contains("Modify_Assay$Select Controls"))
                {
                    _workFlow1.SelecttheControls = table.Rows[0]["Modify_Assay$Select Controls"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.SelecttheControls))
                    {
                        _extentTest.Log(Status.Fail, "Select the control is :" + _workFlow1.SelecttheControls);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Select the control is : " + _workFlow1.SelecttheControls);
                    }
                }

                if (tblcolumns.Contains("Modify_Assay$Injection Name"))
                {
                    _workFlow1.InjectionName = table.Rows[0]["Modify_Assay$Injection Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.InjectionName))
                    {
                        _extentTest.Log(Status.Fail, "The Given injection name is :" + _workFlow1.InjectionName);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given injection name is : " + _workFlow1.InjectionName);
                    }
                }
                #endregion

                #region TestId -7

                if (tblcolumns.Contains("Edit_Page$GraphProperties Verification"))
                {
                    _workFlow1.GraphProperties = table.Rows[0]["Edit_Page$GraphProperties Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.GraphProperties ? "GraphProperties Verification is true" : "GraphProperties Verification is false");
                    if (_workFlow1.GraphProperties)
                    {
                        if (tblcolumns.Contains("Edit_Page$SelectWidgetName"))
                        {
                            var selectWidgetName = table.Rows[0]["Edit_Page$SelectWidgetName"].ToString();
                            if (string.IsNullOrEmpty(selectWidgetName))
                            {
                                _extentTest.Log(Status.Fail, "SelectWidgetName is missing");
                                message += "SelectWidgetName&";
                            }
                            else
                            {
                                _workFlow1.SelectWidgetName = selectWidgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                               selectWidgetName == "Bar Chart" ? WidgetTypes.BarChart : selectWidgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                               selectWidgetName == "Heat Map" ? WidgetTypes.HeatMap : WidgetTypes.KineticGraph;

                                _extentTest.Log(Status.Pass, "AddWidgetName is " + _workFlow1.SelectWidgetName);
                            }
                        }
                    }
                }

                #endregion

                #region TestId -8

                _workFlow1.KineticGraphOcr = new WidgetItems();
                _workFlow1.KineticGraphOcr.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Kinetic_Graph$Measurement"))
                {
                    _workFlow1.KineticGraphOcr.Measurement = table.Rows[0]["Kinetic_Graph$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.Measurement))
                    {
                        _extentTest.Log(Status.Fail, "Measurement for Kinetic graph -OCR is missing");
                        message += "Measurement for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Measurement value for Kineticgraph -OCR  is " + _workFlow1.KineticGraphOcr.Measurement);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Rate-OCR"))
                {
                    _workFlow1.KineticGraphOcr.Rate = table.Rows[0]["Kinetic_Graph$Rate-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.Rate))
                    {
                        _extentTest.Log(Status.Fail, "Ratetype for Kinetic graph -OCR is missing");
                        message += "Ratetype for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Ratetype for Kineticgraph -OCR  is " + _workFlow1.KineticGraphOcr.Rate);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Display"))
                {
                    _workFlow1.KineticGraphOcr.Display = table.Rows[0]["Kinetic_Graph$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Kinetic graph -OCR is missing");
                        message += "Display for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Kineticgraph -OCR  is " + _workFlow1.KineticGraphOcr.Display);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Y"))
                {
                    _workFlow1.KineticGraphOcr.Y = table.Rows[0]["Kinetic_Graph$Y"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.Y))
                    {
                        _extentTest.Log(Status.Fail, "Y-toggle for Kinetic graph -OCR is missing");
                        message += "Y-toggle for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Y-toggle for Kineticgraph -OCR  is " + _workFlow1.KineticGraphOcr.Y);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Normalization"))
                {
                    _workFlow1.KineticGraphOcr.Normalization = table.Rows[0]["Kinetic_Graph$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.Normalization ? "Normalization for Kinetic graph -OCR is true" : "Normalization for Kinetic graph -OCR is false");
                }

                if (tblcolumns.Contains("Kinetic_Graph$Error Format"))
                {
                    _workFlow1.KineticGraphOcr.ErrorFormat = table.Rows[0]["Kinetic_Graph$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Kinetic graph -OCR is missing");
                        message += "Error format for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Kineticgraph -OCR  is " + _workFlow1.KineticGraphOcr.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Background Correction"))
                {
                    _workFlow1.KineticGraphOcr.BackgroundCorrection = table.Rows[0]["Kinetic_Graph$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.BackgroundCorrection ? "Background correction for Kinetic graph -OCR is true" : "Background for Kinetic graph -OCR is false");
                }

                if (tblcolumns.Contains("Kinetic_Graph$Baseline"))
                {
                    _workFlow1.KineticGraphOcr.Baseline = table.Rows[0]["Kinetic_Graph$Baseline"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.Baseline))
                    {
                        _extentTest.Log(Status.Fail, "Baseline for Kinetic graph -OCR is missing");
                        message += "Baseline for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Baseline for Kineticgraph -OCR  is " + _workFlow1.KineticGraphOcr.Baseline);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$GraphUnits-OCR"))
                {
                    _workFlow1.KineticGraphOcr.GraphUnits = table.Rows[0]["Kinetic_Graph$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Kinetic graph OCR value is missing");
                        message += "GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Kinetic graph OCR value is " + _workFlow1.KineticGraphOcr.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$Normalized GraphUnits-OCR"))
                {
                    _workFlow1.KineticGraphOcr.NormalizedGraphUnits = table.Rows[0]["Kinetic_Graph$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.KineticGraphOcr.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Kinetic graph OCR value is missing");
                        message += "Normalized GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Kinetic graph OCR value is " + _workFlow1.KineticGraphOcr.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$GraphSettingsRequired"))
                {
                    _workFlow1.KineticGraphOcr.GraphSettingsVerify = table.Rows[0]["Kinetic_Graph$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(_workFlow1.KineticGraphOcr.GraphSettingsVerify ? Status.Pass : Status.Fail, _workFlow1.KineticGraphOcr.GraphSettingsVerify ? "GraphSettingsVerify for Kinetic graph is true" : "GraphSettingsVerify for Kinetic graph is false");
                    if (_workFlow1.KineticGraphOcr.GraphSettingsVerify)
                    {

                        if (tblcolumns.Contains("Kinetic_Graph$ZeroLine"))
                        {
                            _workFlow1.KineticGraphOcr.GraphSettings.Zeroline = table.Rows[0]["Kinetic_Graph$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Kinetic graph is true" : "Zeroline in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$Linemaker"))
                        {
                            _workFlow1.KineticGraphOcr.GraphSettings.Linemarker = table.Rows[0]["Kinetic_Graph$Linemaker"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.GraphSettings.Linemarker ? "Linemaker in GraphSettings for Kinetic graph is true" : "Linemaker in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$RateHighlight"))
                        {
                            _workFlow1.KineticGraphOcr.GraphSettings.RateHighlight = table.Rows[0]["Kinetic_Graph$RateHighlight"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.GraphSettings.RateHighlight ? "RateHighlight in GraphSettings for Kinetic graph is true" : "RateHighlight in GraphSettings for Kinetic graph is false");
                        }

                        if (tblcolumns.Contains("Kinetic_Graph$InjectionMakers"))
                        {
                            _workFlow1.KineticGraphOcr.GraphSettings.InjectionMakers = table.Rows[0]["Kinetic_Graph$InjectionMakers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.GraphSettings.InjectionMakers ? "InjectionMakers in GraphSettings for Kinetic graph is true" : "InjectionMakers in GraphSettings for Kinetic graph is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Kinetic_Graph$CheckNormalizationWithPlateMap"))
                {
                    _workFlow1.KineticGraphOcr.CheckNormalizationWithPlateMap = table.Rows[0]["Kinetic_Graph$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.KineticGraphOcr.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow1.KineticGraphOcr.IsExportRequired = table.Rows[0]["Kinetic_Graph$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow1.KineticGraphOcr.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.KineticGraphOcr.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.KineticGraphOcr.IsExportRequired);
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

                var EcarNormGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$Normalized GraphUnits-ECAR"))
                {
                    EcarNormGraphUnits = table.Rows[0]["Kinetic_Graph$Normalized GraphUnits-ECAR"].ToString();
                    if (string.IsNullOrEmpty(EcarNormGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Kinetic graph ECAR value is missing");
                        message += "Normalized GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Kinetic graph ECAR value is " + EcarNormGraphUnits);
                    }
                }

                var EcarRate = "";
                if (tblcolumns.Contains("Kinetic_Graph$Rate-ECAR"))
                {
                    EcarRate = table.Rows[0]["Kinetic_Graph$Rate-ECAR"].ToString();
                    if (string.IsNullOrEmpty(EcarRate))
                    {
                        _extentTest.Log(Status.Fail, "Rate for Kinetic graph ECAR value is missing");
                        message += "Rate for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Rate for Kinetic graph ECAR value is " + EcarRate);
                    }
                }

                _workFlow1.KineticGraphEcar = new WidgetItems()
                {
                    Measurement = _workFlow1.KineticGraphOcr.Measurement,
                    Rate = EcarRate,
                    Display = _workFlow1.KineticGraphOcr.Display,
                    Y = _workFlow1.KineticGraphOcr.Y,
                    Normalization = _workFlow1.KineticGraphOcr.Normalization,
                    ErrorFormat = _workFlow1.KineticGraphOcr.ErrorFormat,
                    BackgroundCorrection = _workFlow1.KineticGraphOcr.BackgroundCorrection,
                    Baseline = _workFlow1.KineticGraphOcr.Baseline,
                    GraphSettings = _workFlow1.KineticGraphOcr.GraphSettings,
                    GraphSettingsVerify = _workFlow1.KineticGraphOcr.GraphSettingsVerify,
                    GraphUnits = EcarGraphUnits,
                    NormalizedGraphUnits = EcarNormGraphUnits
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

                var PerNormGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$Normalized GraphUnits-PER"))
                {
                    PerNormGraphUnits = table.Rows[0]["Kinetic_Graph$Normalized GraphUnits-PER"].ToString();
                    if (string.IsNullOrEmpty(PerNormGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Kinetic graph PER value is missing");
                        message += "Normalized GraphUnits for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Kinetic graph PER value is " + PerNormGraphUnits);
                    }
                }
                var PerRate = "";
                if (tblcolumns.Contains("Kinetic_Graph$Rate-PER"))
                {
                    PerRate = table.Rows[0]["Kinetic_Graph$Rate-PER"].ToString();
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


                _workFlow1.KineticGraphPer = new WidgetItems()
                {
                    Measurement = _workFlow1.KineticGraphOcr.Measurement,
                    Rate = PerRate,
                    Display = _workFlow1.KineticGraphOcr.Display,
                    Y = _workFlow1.KineticGraphOcr.Y,
                    Normalization = _workFlow1.KineticGraphOcr.Normalization,
                    ErrorFormat = _workFlow1.KineticGraphOcr.ErrorFormat,
                    BackgroundCorrection = _workFlow1.KineticGraphOcr.BackgroundCorrection,
                    Baseline = _workFlow1.KineticGraphOcr.Baseline,
                    GraphSettings = _workFlow1.KineticGraphOcr.GraphSettings,
                    GraphSettingsVerify = _workFlow1.KineticGraphOcr.GraphSettingsVerify,
                    GraphUnits = PerGraphUnits,
                    NormalizedGraphUnits = PerNormGraphUnits
                };

                #endregion

                #region TestId -9

                _workFlow1.Barchart = new WidgetItems();
                _workFlow1.Barchart.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Bar_Chart$Measurement"))
                {
                    _workFlow1.Barchart.Measurement = table.Rows[0]["Bar_Chart$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.Measurement))
                    {
                        _extentTest.Log(Status.Fail, "Measurement for Barchart is missing");
                        message += "Measurement for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Measurement value for Barchart is " + _workFlow1.Barchart.Measurement);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Rate"))
                {
                    _workFlow1.Barchart.Rate = table.Rows[0]["Bar_Chart$Rate"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.Rate))
                    {
                        _extentTest.Log(Status.Fail, "Rate for Barchart is missing");
                        message += "Rate for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Rate for Barchart Rate is " + _workFlow1.Barchart.Rate);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Display"))
                {
                    _workFlow1.Barchart.Display = table.Rows[0]["Bar_Chart$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display value for Barchart is missing");
                        message += "Display for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Display value for Barchart is " + _workFlow1.Barchart.Display);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Normalization"))
                {
                    _workFlow1.Barchart.Normalization = table.Rows[0]["Bar_Chart$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.Barchart.Normalization ? "Normalization for Barchart is true" : "Normalization for Barchart is false");
                }

                if (tblcolumns.Contains("Bar_Chart$Error Format"))
                {
                    _workFlow1.Barchart.ErrorFormat = table.Rows[0]["Bar_Chart$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error Format for Barchart is missing");
                        message += "Error Format for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Error Format for Barchart is " + _workFlow1.Barchart.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Background Correction"))
                {
                    _workFlow1.Barchart.BackgroundCorrection = table.Rows[0]["Bar_Chart$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.Barchart.BackgroundCorrection ? "Background Correction for Barchart is true" : "Background Correction for Barchart is false");
                }

                if (tblcolumns.Contains("Bar_Chart$Baseline"))
                {
                    _workFlow1.Barchart.Baseline = table.Rows[0]["Bar_Chart$Baseline"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.Baseline))
                    {
                        _extentTest.Log(Status.Fail, "Baseline for Barchart is missing");
                        message += "Baseline for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Baseline for Barchart is " + _workFlow1.Barchart.Baseline);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$GraphUnits"))
                {
                    _workFlow1.Barchart.GraphUnits = table.Rows[0]["Bar_Chart$GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Barchart is missing");
                        message += "GraphUnits for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Barchart is " + _workFlow1.Barchart.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$Normalized GraphUnits"))
                {
                    _workFlow1.Barchart.NormalizedGraphUnits = table.Rows[0]["Bar_Chart$Normalized GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.Barchart.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Barchart is missing");
                        message += "Normalized GraphUnits for Barchart&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Barchart is " + _workFlow1.Barchart.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Bar_Chart$GraphSettingsRequired"))
                {
                    _workFlow1.Barchart.GraphSettingsVerify = table.Rows[0]["Bar_Chart$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.Barchart.GraphSettingsVerify ? "GraphSettingsVerify for Barchart is true" : "GraphSettingsVerify for Barchart is false");
                    if (_workFlow1.Barchart.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Bar_Chart$ZeroLine"))
                        {
                            _workFlow1.Barchart.GraphSettings.Zeroline = table.Rows[0]["Bar_Chart$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.Barchart.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Barchart is true" : "Zeroline in GraphSettings for Barchart is false");
                        }

                    }
                }

                if (tblcolumns.Contains("Bar_Chart$CheckNormalizationWithPlateMap"))
                {
                    _workFlow1.Barchart.CheckNormalizationWithPlateMap = table.Rows[0]["Bar_Chart$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.Barchart.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    if (tblcolumns.Contains("Bar_Chart$IsExportRequired"))
                    {
                        _workFlow1.Barchart.IsExportRequired = table.Rows[0]["Bar_Chart$IsExportRequired"].ToString() == "Yes";
                        if (_workFlow1.Barchart.IsExportRequired)
                        {
                            _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.Barchart.IsExportRequired);
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.Barchart.IsExportRequired);
                        }
                    }
                }
                #endregion

                #region TestId -10

                _workFlow1.EnergyMap = new WidgetItems();
                _workFlow1.EnergyMap.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Energy_Map$Measurement"))
                {
                    _workFlow1.EnergyMap.Measurement = table.Rows[0]["Energy_Map$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.Measurement))
                    {
                        _extentTest.Log(Status.Fail, " Measurement for EnergyMap is missing");
                        message += " Measurement for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Measurement value for EnergyMap is " + _workFlow1.EnergyMap.Measurement);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$Rate"))
                {
                    _workFlow1.EnergyMap.Rate = table.Rows[0]["Energy_Map$Rate"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.Rate))
                    {
                        _extentTest.Log(Status.Fail, " Rate for EnergyMap is missing");
                        message += " Rate for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Rate for EnergyMap is " + _workFlow1.EnergyMap.Rate);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$Display"))
                {
                    _workFlow1.EnergyMap.Display = table.Rows[0]["Energy_Map$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.Display))
                    {
                        _extentTest.Log(Status.Fail, " Display for EnergyMap is missing");
                        message += " Display for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Display for EnergyMap is " + _workFlow1.EnergyMap.Display);
                    }
                }


                if (tblcolumns.Contains("Energy_Map$Normalization"))
                {
                    _workFlow1.EnergyMap.Normalization = table.Rows[0]["Energy_Map$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.EnergyMap.Normalization ? " normalization for Energy graph is true" : " normalization for Energy graph is false");
                }


                if (tblcolumns.Contains("Energy_Map$Error Format"))
                {
                    _workFlow1.EnergyMap.ErrorFormat = table.Rows[0]["Energy_Map$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, " Error Format for EnergyMap is missing");
                        message += " Error Format for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Error Format for EnergyMap is " + _workFlow1.EnergyMap.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$Background Correction"))
                {
                    _workFlow1.EnergyMap.BackgroundCorrection = table.Rows[0]["Energy_Map$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.EnergyMap.BackgroundCorrection ? " Background Correction for Energy graph is true" : " Background Correction for Energy graph is false");
                }

                if (tblcolumns.Contains("Energy_Map$BaseLine"))
                {
                    _workFlow1.EnergyMap.Baseline = table.Rows[0]["Energy_Map$BaseLine"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.Baseline))
                    {
                        _extentTest.Log(Status.Fail, " Baseline for EnergyMap is missing");
                        message += " Baseline for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Baseline for EnergyMap is " + _workFlow1.EnergyMap.Baseline);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$GraphUnits"))
                {
                    _workFlow1.EnergyMap.GraphUnits = table.Rows[0]["Energy_Map$GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for EnergyMap is missing");
                        message += "GraphUnits for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for EnergyMap is " + _workFlow1.EnergyMap.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$Normalized GraphUnits"))
                {
                    _workFlow1.EnergyMap.NormalizedGraphUnits = table.Rows[0]["Energy_Map$Normalized GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.EnergyMap.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for EnergyMap is missing");
                        message += "Normalized GraphUnits for EnergyMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized Graph Units for Energy Map is " + _workFlow1.EnergyMap.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Energy_Map$GraphSettingsRequired"))
                {
                    _workFlow1.EnergyMap.GraphSettingsVerify = table.Rows[0]["Energy_Map$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.EnergyMap.GraphSettingsVerify ? "GraphSettingsVerify for EnergyMap is true" : "GraphSettingsVerify for EnergyMap is false");
                }

                if (tblcolumns.Contains("Energy_Map$CheckNormalizationWithPlateMap"))
                {
                    _workFlow1.EnergyMap.CheckNormalizationWithPlateMap = table.Rows[0]["Energy_Map$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.EnergyMap.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Energy_Map$IsExportRequired"))
                {
                    _workFlow1.EnergyMap.IsExportRequired = table.Rows[0]["Energy_Map$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow1.EnergyMap.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.EnergyMap.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.EnergyMap.IsExportRequired);
                    }
                }

                #endregion

                #region TestId -11
                _workFlow1.HeatMap = new WidgetItems();
                _workFlow1.HeatMap.GraphSettings = new GraphSettings();
                _workFlow1.HeatMap.KitValidation = new KitValidation();
                _workFlow1.HeatMap.HeatTolerance = new HeatTolerance();


                if (tblcolumns.Contains("Heat_Map$Measurement"))
                {
                    _workFlow1.HeatMap.Measurement = table.Rows[0]["Heat_Map$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.HeatMap.Measurement))
                    {
                        _extentTest.Log(Status.Fail, " Measurement for HeatMap is missing");
                        message += " Measurement for HeatMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Measurement value for HeatMap is " + _workFlow1.HeatMap.Measurement);
                    }
                }

                if (tblcolumns.Contains("Heat_Map$Rate"))
                {
                    _workFlow1.HeatMap.Rate = table.Rows[0]["Heat_Map$Rate"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.HeatMap.Rate))
                    {
                        _extentTest.Log(Status.Fail, " Rate for HeatMap is missing");
                        message += " Rate for HeatMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Rate for HeatMap is " + _workFlow1.HeatMap.Rate);
                    }
                }

                if (tblcolumns.Contains("Heat_Map$Normalization"))
                {
                    _workFlow1.HeatMap.Normalization = table.Rows[0]["Heat_Map$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.HeatMap.Normalization ? " normalization for HeatMap is true" : " normalization for HeatMap is false");
                }

                if (tblcolumns.Contains("Heat_Map$Background Correction"))
                {
                    _workFlow1.HeatMap.BackgroundCorrection = table.Rows[0]["Heat_Map$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.HeatMap.BackgroundCorrection ? " Background Correction for HeatMap is true" : " Background Correction for HeatMap is false");
                }

                if (tblcolumns.Contains("Heat_Map$BaseLine"))
                {
                    _workFlow1.HeatMap.Baseline = table.Rows[0]["Heat_Map$BaseLine"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.HeatMap.Baseline))
                    {
                        _extentTest.Log(Status.Fail, " Baseline for HeatMap is missing");
                        message += " Baseline for HeatMap&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Baseline for HeatMap is " + _workFlow1.HeatMap.Baseline);
                    }
                }

                if (tblcolumns.Contains("Heat_Map$GraphSettingsRequired"))
                {
                    _workFlow1.HeatMap.GraphSettingsVerify = table.Rows[0]["Heat_Map$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.HeatMap.GraphSettingsVerify ? "GraphSettingsVerify for HeatMap is true" : "GraphSettingsVerify for HeatMap is false");
                    if (_workFlow1.HeatMap.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Energy_Map$IsExportRequired"))
                        {
                            if (tblcolumns.Contains("Heat_Map$ZeroLine"))
                            {
                                _workFlow1.HeatMap.GraphSettings.Zeroline = table.Rows[0]["Heat_Map$ZeroLine"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow1.HeatMap.GraphSettings.Zeroline ? "Zeroline in GraphSettings for HeatMap is true" : "Zeroline in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$Linemaker"))
                            {
                                _workFlow1.HeatMap.GraphSettings.Linemarker = table.Rows[0]["Heat_Map$Linemaker"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow1.HeatMap.GraphSettings.Linemarker ? "Linemaker in GraphSettings for HeatMap is true" : "Linemaker in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$RateHighlight"))
                            {
                                _workFlow1.HeatMap.GraphSettings.RateHighlight = table.Rows[0]["Heat_Map$RateHighlight"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow1.HeatMap.GraphSettings.RateHighlight ? "RateHighlight in GraphSettings for HeatMap is true" : "RateHighlight in GraphSettings for HeatMap is false");
                            }

                            if (tblcolumns.Contains("Heat_Map$InjectionMakers"))
                            {
                                _workFlow1.HeatMap.GraphSettings.InjectionMakers = table.Rows[0]["Heat_Map$InjectionMakers"].ToString() == "Yes";
                                _extentTest.Log(Status.Pass, _workFlow1.HeatMap.GraphSettings.InjectionMakers ? "InjectionMakers in GraphSettings for HeatMap is true" : "InjectionMakers in GraphSettings for HeatMap is false");
                            }
                        }
                    }
                }

                if (tblcolumns.Contains("Heat_Map$AssayKit Validation"))
                {
                    _workFlow1.HeatMap.KitValidation.AssayKitValidation = table.Rows[0]["Heat_Map$AssayKit Validation"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.HeatMap.KitValidation.AssayKitValidation ? "Assaykit Validation for HeatMap is true" : "Assaykit Validation for HeatMap is false");
                    if (_workFlow1.HeatMap.KitValidation.AssayKitValidation)
                    {
                        if (tblcolumns.Contains("Heat_Map$Cat Number"))
                        {
                            _workFlow1.HeatMap.KitValidation.CatNumber = table.Rows[0]["Heat_Map$Cat Number"].ToString();
                            if (string.IsNullOrEmpty(_workFlow1.HeatMap.KitValidation.CatNumber))
                            {
                                _extentTest.Log(Status.Fail, "Cat Number for HeatMap is missing");
                                message += "Cat Number for HeatMap&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "Cat Number for HeatMap is " + _workFlow1.HeatMap.KitValidation.CatNumber);
                            }
                        }

                        if (tblcolumns.Contains("Heat_Map$Lot Number"))
                        {
                            _workFlow1.HeatMap.KitValidation.LotNumber = table.Rows[0]["Heat_Map$Lot Number"].ToString();
                            if (string.IsNullOrEmpty(_workFlow1.HeatMap.KitValidation.LotNumber))
                            {
                                _extentTest.Log(Status.Fail, "Lot Number for HeatMap is missing");
                                message += "Lot Number for HeatMap&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "Lot Number for HeatMap is " + _workFlow1.HeatMap.KitValidation.LotNumber);
                            }
                        }

                        if (tblcolumns.Contains("Heat_Map$SW ID"))
                        {
                            _workFlow1.HeatMap.KitValidation.SWID = table.Rows[0]["Heat_Map$SW ID"].ToString();
                            if (string.IsNullOrEmpty(_workFlow1.HeatMap.KitValidation.SWID))
                            {
                                _extentTest.Log(Status.Fail, "SW ID Number for HeatMap is missing");
                                message += "SW ID Number for HeatMap&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "SW ID Number for HeatMap is " + _workFlow1.HeatMap.KitValidation.SWID);
                            }
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$Colour Options"))
                    {
                        _workFlow1.HeatMap.HeatTolerance.ColourOptions = table.Rows[0]["Heat_Map$Colour Options"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _workFlow1.HeatMap.HeatTolerance.ColourOptions ? "Colour Options for HeatMap is true" : "Colour Options for HeatMap is false");
                        if (_workFlow1.HeatMap.HeatTolerance.ColourOptions)
                        {
                            if (tblcolumns.Contains("Heat_Map$Colour Tolerance %"))
                            {
                                _workFlow1.HeatMap.HeatTolerance.ColourTolerance = table.Rows[0]["Heat_Map$Colour Tolerance %"].ToString();
                                if (string.IsNullOrEmpty(_workFlow1.HeatMap.HeatTolerance.ColourTolerance))
                                {
                                    _extentTest.Log(Status.Fail, "Colour Tolerance % for HeatMap is missing");
                                    message += "Colour Tolerance % for HeatMap&";
                                }
                                else
                                {
                                    _extentTest.Log(Status.Pass, "Colour Tolerance for HeatMap is " + _workFlow1.HeatMap.HeatTolerance.ColourTolerance + " % ");
                                }
                            }
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$GraphUnits"))
                    {
                        _workFlow1.HeatMap.GraphUnits = table.Rows[0]["Heat_Map$GraphUnits"].ToString();
                        if (string.IsNullOrEmpty(_workFlow1.HeatMap.GraphUnits))
                        {
                            _extentTest.Log(Status.Fail, "GraphUnits for Heat Map is missing");
                            message += "GraphUnits for Heat Map&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "GraphUnits for Heat Map is " + _workFlow1.HeatMap.GraphUnits);
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$Normalized GraphUnits"))
                    {
                        _workFlow1.HeatMap.NormalizedGraphUnits = table.Rows[0]["Heat_Map$Normalized GraphUnits"].ToString();
                        if (string.IsNullOrEmpty(_workFlow1.HeatMap.NormalizedGraphUnits))
                        {
                            _extentTest.Log(Status.Fail, "Normalized GraphUnits for HeatMap is missing");
                            message += "Normalized GraphUnits for HeatMap&";
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "Normalized Graph Units for Heat Map is " + _workFlow1.HeatMap.NormalizedGraphUnits);
                        }
                    }

                    if (tblcolumns.Contains("Heat_Map$CheckNormalizationWithPlateMap"))
                    {
                        _workFlow1.HeatMap.CheckNormalizationWithPlateMap = table.Rows[0]["Heat_Map$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                        _extentTest.Log(Status.Pass, _workFlow1.HeatMap.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                    }

                    if (tblcolumns.Contains("Heat_Map$IsExportRequired"))
                    {
                        _workFlow1.HeatMap.IsExportRequired = table.Rows[0]["Heat_Map$IsExportRequired"].ToString() == "Yes";
                        if (_workFlow1.HeatMap.IsExportRequired)
                        {
                            _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.HeatMap.IsExportRequired);
                        }
                        else
                        {
                            _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.HeatMap.IsExportRequired);
                        }
                    }
                }
                #endregion

                #region TestId -12
                _workFlow1.DoseResponseWidget = new WidgetItems();

                if (tblcolumns.Contains("Dose_Response_Add_Widget$Prerequisite"))
                {
                    _workFlow1.DoseResponseWidget.DoseResponseAddWidget = table.Rows[0]["Dose_Response_Add_Widget$DoseResponseAddWidget"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.DoseResponseWidget.DoseResponseAddWidget ? "Add widget for Dose response is true" : "Add widget for Dose response is false");
                }
                #endregion

                #region TestId -13
                _workFlow1.DoseResponseView = new WidgetItems();

                if (tblcolumns.Contains("Dose_Response_Add_View$DoseResponseAddView"))
                {
                    _workFlow1.DoseResponseView.DoseResponseAddView = table.Rows[0]["Dose_Response_Add_View$DoseResponseAddView"].ToString() == "Yes";
                    if (_workFlow1.DoseResponseView.DoseResponseAddView)
                    {
                        _workFlow1.AddDoseWidget = new List<WidgetTypes>();
                        _workFlow1.AddDoseWidget.Add(WidgetTypes.DoseResponse);
                    }
                    _extentTest.Log(Status.Pass, _workFlow1.DoseResponseView.DoseResponseAddView ? "Add view for Dose response is true" : "Add view for Dose response is false");
                }
                    
                #endregion

                #region TestId -14

                _workFlow1.DoseResponse = new WidgetItems();
                _workFlow1.DoseResponse.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Dose_Response$Measurement"))
                {
                    _workFlow1.DoseResponse.Measurement = table.Rows[0]["Dose_Response$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.DoseResponse.Measurement))
                    {
                        _extentTest.Log(Status.Fail, " Measurement for Dose Response is missing");
                        message += " Measurement for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Measurement value for Dose Response is " + _workFlow1.DoseResponse.Measurement);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$Rate"))
                {
                    _workFlow1.DoseResponse.Rate = table.Rows[0]["Dose_Response$Rate"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.DoseResponse.Rate))
                    {
                        _extentTest.Log(Status.Fail, " Rate for Dose Response is missing");
                        message += " Rate for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Rate for Dose response is " + _workFlow1.DoseResponse.Rate);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$Normalization"))
                {
                    _workFlow1.DoseResponse.Normalization = table.Rows[0]["Dose_Response$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.Normalization ? " normalization for Dose response is true" : " normalization for Dose response is false");
                }

                if (tblcolumns.Contains("Dose_Response$Error Format"))
                {
                    _workFlow1.DoseResponse.ErrorFormat = table.Rows[0]["Dose_Response$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.DoseResponse.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, " Error Format for Dose response is missing");
                        message += " Error Format for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, " Error Format for Dose Response is " + _workFlow1.DoseResponse.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$Background Correction"))
                {
                    _workFlow1.DoseResponse.BackgroundCorrection = table.Rows[0]["Dose_Response$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.BackgroundCorrection ? " Background Correction for Dose Response is true" : " Background Correction for Dose Response is false");
                }

                if (tblcolumns.Contains("Dose_Response$GraphSettingsRequired"))
                {
                    _workFlow1.DoseResponse.GraphSettingsVerify = table.Rows[0]["Dose_Response$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.GraphSettingsVerify ? "GraphSettingsVerify for Dose Response is true" : "GraphSettingsVerify for Dose Response is false");
                    if (_workFlow1.DoseResponse.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Dose_Response$ZeroLine"))
                        {
                            _workFlow1.DoseResponse.GraphSettings.Zeroline = table.Rows[0]["Dose_Response$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Dose response is true" : "Zeroline in GraphSettings for Dose Response is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$Linemaker"))
                        {
                            _workFlow1.DoseResponse.GraphSettings.Linemarker = table.Rows[0]["Dose_Response$Linemaker"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.GraphSettings.Linemarker ? "Linemaker in GraphSettings for Dose response is true" : "Linemaker in GraphSettings for Dose Response is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$RateHighlight"))
                        {
                            _workFlow1.DoseResponse.GraphSettings.RateHighlight = table.Rows[0]["Dose_Response$RateHighlight"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.GraphSettings.RateHighlight ? "RateHighlight in GraphSettings for Dose Response is true" : "RateHighlight in GraphSettings for Dose Response is false");
                        }

                        if (tblcolumns.Contains("Dose_Response$InjectionMakers"))
                        {
                            _workFlow1.DoseResponse.GraphSettings.InjectionMakers = table.Rows[0]["Dose_Response$InjectionMakers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.GraphSettings.InjectionMakers ? "InjectionMakers in GraphSettings for Dose Response is true" : "InjectionMakers in GraphSettings for Dose Response is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Dose_Response$GraphUnits"))
                {
                    _workFlow1.DoseResponse.GraphUnits = table.Rows[0]["Dose_Response$GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.DoseResponse.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Dose Response is missing");
                        message += "GraphUnits for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Dose Response is " + _workFlow1.DoseResponse.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$Normalized GraphUnits"))
                {
                    _workFlow1.DoseResponse.NormalizedGraphUnits = table.Rows[0]["Dose_Response$Normalized GraphUnits"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.DoseResponse.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Dose Response is missing");
                        message += "Normalized GraphUnits for Dose Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized Graph Units for Dose Response is " + _workFlow1.DoseResponse.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Dose_Response$CheckNormalizationWithPlateMap"))
                {
                    _workFlow1.DoseResponse.CheckNormalizationWithPlateMap = table.Rows[0]["Dose_Response$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.DoseResponse.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Dose_Response$IsExportRequired"))
                {
                    _workFlow1.DoseResponse.IsExportRequired = table.Rows[0]["Dose_Response$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow1.DoseResponse.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.DoseResponse.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow1.DoseResponse.IsExportRequired);
                    }
                }

                #endregion

                #region TestId -15

                if (tblcolumns.Contains("Blank_View$CreateBlankView"))
                {
                    _workFlow1.CreateBlankView = table.Rows[0]["Blank_View$CreateBlankView"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow1.CreateBlankView ? "Create Blank View is true" : "Create Blank View is false");
                    if (_workFlow1.CreateBlankView)
                    {
                        //Kinetic Graph - Ocr, Kinetic Graph - Ecar, Kinetic Graph - Per, Bar Chart, Energetic Map, Heat Map

                        var widgetName = table.Rows[0]["Blank_View$AddBlankWidget"].ToString();
                        if (string.IsNullOrEmpty(widgetName))
                        {
                            _extentTest.Log(Status.Fail, "AddBlankWidget is missing");
                            message += "AddBlankWidget&";
                        }
                        else
                        {
                            _workFlow1.AddBlankWidget = widgetName == "Kinetic Graph - Ocr" ? WidgetTypes.KineticGraph :
                            widgetName == "Bar Chart" ? WidgetTypes.BarChart : widgetName == "Energy Map" ? WidgetTypes.EnergyMap :
                            widgetName == "Heat Map" ? WidgetTypes.HeatMap : widgetName == "Dose Response" ? WidgetTypes.DoseResponse : WidgetTypes.KineticGraph;

                            _extentTest.Log(Status.Pass, "AddBlankWidget is " + _workFlow1.AddBlankWidget);
                        }
                    }
                }
                #endregion

                #region TestId -16

                if (tblcolumns.Contains("Custom_View$CustomViewName"))
                {
                    _workFlow1.CustomViewName = table.Rows[0]["Custom_View$CustomViewName"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.CustomViewName))
                    {
                        _extentTest.Log(Status.Fail, "Custom View Name is empty :" + _workFlow1.CustomViewName);
                        message += "CustomViewName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "CustomView Name is: " + _workFlow1.CustomViewName);
                    }
                }

                if (tblcolumns.Contains("Custom_View$CustomViewDescription"))
                {
                    _workFlow1.CustomViewDescription = table.Rows[0]["Custom_View$CustomViewDescription"].ToString();
                    if (string.IsNullOrEmpty(_workFlow1.CustomViewDescription))
                    {
                        _extentTest.Log(Status.Fail, "Custom view description is empty :" + _workFlow1.CustomViewDescription);
                        message += "CustomViewDescription&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Custom View description is: " + _workFlow1.CustomViewDescription);
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
            else if (sheetName == "Workflow-6")
            {
                #region TestId-1
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
                }
                else
                {
                    _extentTest.Log(Status.Pass, "FileUpload required status is false");
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

                _fileUploadOrExistingFileData.OligoInjection = table.Rows[0]["Upload_File$Oligo Injection"].ToString();
                if (string.IsNullOrEmpty(_fileUploadOrExistingFileData.OligoInjection))
                {
                    _extentTest.Log(Status.Fail, "Oligo Injection is empty");
                    message += "Oligo Injection&";
                }
                else
                {
                    _extentTest.Log(Status.Pass, "Oligo Injection is selected as - " + _fileUploadOrExistingFileData.OligoInjection + " Injection");
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

                _fileUploadOrExistingFileData.SelectedWidgets = new List<WidgetTypes>();
                if (table.Rows[0]["Upload_File$Mitochondrial Respiration"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.MitochondrialRespiration);
                }
                if (table.Rows[0]["Upload_File$Basal Respiration"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.Basal);
                }
                if (table.Rows[0]["Upload_File$Acute Response"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.AcuteResponse);
                }
                if (table.Rows[0]["Upload_File$Proton Leak"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.ProtonLeak);
                }
                if (table.Rows[0]["Upload_File$Maximal Respiration"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.MaximalRespiration);
                }
                if (table.Rows[0]["Upload_File$Spare Respiratory Capacity"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.SpareRespiratoryCapacity);
                }
                if (table.Rows[0]["Upload_File$Non mito O2 consumption"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.NonMitoO2Consumption);
                }
                if (table.Rows[0]["Upload_File$ATP Production Coupled Respiration"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.AtpProductionCoupledRespiration);
                }
                if (table.Rows[0]["Upload_File$Coupling Efficiency"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.CouplingEfficiencyPercent);
                }
                if (table.Rows[0]["Upload_File$Spare Respiratory Capacity Percentage"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.SpareRespiratoryCapacityPercent);
                }
                if (table.Rows[0]["Upload_File$Data Table"].ToString() == "Yes")
                {
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.DataTable);
                }
                #endregion

                #region TestId - 5
                _workFlow6.MitochondrialRespiration = new WidgetItems();
                _workFlow6.MitochondrialRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Mitochondrial_Respiration$Measurement"))
                {
                    _workFlow6.MitochondrialRespiration.Measurement = table.Rows[0]["Mitochondrial_Respiration$Measurement"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.Measurement))
                    {
                        _extentTest.Log(Status.Fail, "Measurement for Mitochondrial Respiration is missing");
                        message += "Measurement for Mitochiondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Measurement value for Mitochondrial Respiration is " + _workFlow6.MitochondrialRespiration.Measurement);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Rate"))
                {
                    _workFlow6.MitochondrialRespiration.Rate = table.Rows[0]["Mitochondrial_Respiration$Rate"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.Rate))
                    {
                        _extentTest.Log(Status.Fail, "Ratetype for Mitochondrial Respiration is missing");
                        message += "Ratetype for Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Ratetype for Mitochondrial Respiration is " + _workFlow6.MitochondrialRespiration.Rate);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Display"))
                {
                    _workFlow6.MitochondrialRespiration.Display = table.Rows[0]["Mitochondrial_Respiration$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Mitochondrial Respiration is missing");
                        message += "Display for Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Mitochondrial Respiration  is " + _workFlow6.MitochondrialRespiration.Display);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Y"))
                {
                    _workFlow6.MitochondrialRespiration.Y = table.Rows[0]["Mitochondrial_Respiration$Y"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.Y))
                    {
                        _extentTest.Log(Status.Fail, "Y-toggle for Mitochondrial Respiration is missing");
                        message += "Default y-toggle for Kinetic graph&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Y-toggle for Mitochondrial Respiration  is " + _workFlow6.MitochondrialRespiration.Y);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Normalization"))
                {
                    _workFlow6.MitochondrialRespiration.Normalization = table.Rows[0]["Mitochondrial_Respiration$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.Normalization ? "Normalization for Mitochiondrial Respiration is true" : "Normalization for Mitochiondrial Respiration is false");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Error Format"))
                {
                    _workFlow6.MitochondrialRespiration.ErrorFormat = table.Rows[0]["Mitochondrial_Respiration$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Mitochondrial Respiration is missing");
                        message += "Error format for Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Mitochondrial Respiration  is " + _workFlow6.MitochondrialRespiration.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Background Correction"))
                {
                    _workFlow6.MitochondrialRespiration.BackgroundCorrection = table.Rows[0]["Mitochondrial_Respiration$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.BackgroundCorrection ? "Background correction for Mitochondrial Respiration is true" : "Background for Mitochondrial Respiration is false");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Baseline"))
                {
                    _workFlow6.MitochondrialRespiration.Baseline = table.Rows[0]["Mitochondrial_Respiration$Baseline"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.Baseline))
                    {
                        _extentTest.Log(Status.Fail, "Baseline for Mitochondrial Respiration is missing");
                        message += "Baseline for Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Baseline for Mitochiondrial Respiration  is " + _workFlow6.MitochondrialRespiration.Baseline);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$GraphUnits-OCR"))
                {
                    _workFlow6.MitochondrialRespiration.GraphUnits = table.Rows[0]["Mitochondrial_Respiration$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Mitochondrial Respiration value is missing");
                        message += "GraphUnits for Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Mitochondrial Respiration value is " + _workFlow6.MitochondrialRespiration.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.MitochondrialRespiration.NormalizedGraphUnits = table.Rows[0]["Mitochondrial_Respiration$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MitochondrialRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Mitochondrial Respiration value is missing");
                        message += "Normalized GraphUnits for Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Mitochondrial Respiration value is " + _workFlow6.MitochondrialRespiration.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$GraphSettingsRequired"))
                {
                    _workFlow6.MitochondrialRespiration.GraphSettingsVerify = table.Rows[0]["Mitochondrial_Respiration$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettingsVerify ? "GraphSettingsVerify for Kinetic graph is true" : "GraphSettingsVerify for Kinetic graph is false");
                    if (_workFlow6.MitochondrialRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Mitochondrial_Respiration$ZeroLine"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.Zeroline = table.Rows[0]["Mitochondrial_Respiration$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Mitochondrial Respiration is true" : "Zeroline in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$Linemaker"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.Linemarker = table.Rows[0]["Mitochondrial_Respiration$Linemaker"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.Linemarker ? "Linemaker in GraphSettings for Mitochondrial Respiration is true" : "Linemaker in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$RateHighlight"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RateHighlight = table.Rows[0]["Mitochondrial_Respiration$RateHighlight"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RateHighlight ? "RateHighlight in GraphSettings for Mitochondrial Respiration is true" : "RateHighlight in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$InjectionMakers"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.InjectionMakers = table.Rows[0]["Mitochondrial_Respiration$InjectionMakers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.InjectionMakers ? "InjectionMakers in GraphSettings for Mitochondrial Respiration is true" : "InjectionMakers in GraphSettings for Mitochondrial Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.MitochondrialRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["Mitochondrial_Respiration$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.MitochondrialRespiration.IsExportRequired = table.Rows[0]["Mitochondrial_Respiration$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.MitochondrialRespiration.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.MitochondrialRespiration.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Fail, "File Normalization status is " + _workFlow6.MitochondrialRespiration.IsExportRequired);
                    }
                }
                #endregion

                #region TestId -6

                _workFlow6.BasalRespiration = new WidgetItems();
                _workFlow6.BasalRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Basal_Respiration$Oligo"))
                {
                    _workFlow6.BasalRespiration.Oligo = table.Rows[0]["Basal_Respiration$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for Basal Respiration is missing");
                        message += "Default oligo for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for Basal Respiration  is " + _workFlow6.BasalRespiration.Oligo);
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$Display"))
                {
                    _workFlow6.BasalRespiration.Display = table.Rows[0]["Basal_Respiration$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Basal Respiration is missing");
                        message += "Display for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Basal Respiration  is " + _workFlow6.BasalRespiration.Display);
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$Normalization"))
                {
                    _workFlow6.BasalRespiration.Normalization = table.Rows[0]["Basal_Respiration$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.Normalization ? "Normalization for Basal Respiration is true" : "Normalization for Basal Respiration is false");
                }

                if (tblcolumns.Contains("Basal_Respiration$Error Format"))
                {
                    _workFlow6.BasalRespiration.ErrorFormat = table.Rows[0]["Basal_Respiration$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Basal Respiration is missing");
                        message += "Error format for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Basal Respiration  is " + _workFlow6.BasalRespiration.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$GraphUnits-OCR"))
                {
                    _workFlow6.BasalRespiration.GraphUnits = table.Rows[0]["Basal_Respiration$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Basal Respiration value is missing");
                        message += "GraphUnits for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Basal Respiration value is " + _workFlow6.BasalRespiration.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.BasalRespiration.NormalizedGraphUnits = table.Rows[0]["Basal_Respiration$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Basal Respiration value is missing");
                        message += "Normalized GraphUnits for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Basal Respiration value is " + _workFlow6.BasalRespiration.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$GraphSettingsRequired"))
                {
                    _workFlow6.BasalRespiration.GraphSettingsVerify = table.Rows[0]["Basal_Respiration$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettingsVerify ? "GraphSettingsVerify for Basal Respiration is true" : "GraphSettingsVerify for Basal Respiration is false");
                    if (_workFlow6.BasalRespiration.GraphSettingsVerify)
                    {

                        if (tblcolumns.Contains("Basal_Respiration$ZeroLine"))
                        {
                            _workFlow6.BasalRespiration.GraphSettings.Zeroline = table.Rows[0]["Basal_Respiration$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Basal Respiration is true" : "Zeroline in GraphSettings for Basal Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.BasalRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["Basal_Respiration$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.BasalRespiration.IsExportRequired = table.Rows[0]["Basal_Respiration$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.BasalRespiration.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.BasalRespiration.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.BasalRespiration.IsExportRequired);
                    }
                }
                #endregion

                #region Testid -7

                _workFlow6.AcuteResponse = new WidgetItems();
                _workFlow6.AcuteResponse.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Acute_Response$Display"))
                {
                    _workFlow6.AcuteResponse.Display = table.Rows[0]["Acute_Response$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.AcuteResponse.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Acute Response is missing");
                        message += "Display for Acute Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Acute Response  is " + _workFlow6.AcuteResponse.Display);
                    }
                }

                if (tblcolumns.Contains("Acute_Response$Normalization"))
                {
                    _workFlow6.AcuteResponse.Normalization = table.Rows[0]["Acute_Response$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.Normalization ? "Normalization for Acute Response is true" : "Normalization for Acute Response is false");
                }

                if (tblcolumns.Contains("Acute_Response$Error Format"))
                {
                    _workFlow6.AcuteResponse.ErrorFormat = table.Rows[0]["Acute_Response$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.AcuteResponse.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Acute Response is missing");
                        message += "Error format for Acute Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Acute Response  is " + _workFlow6.AcuteResponse.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Acute_Response$GraphUnits-OCR"))
                {
                    _workFlow6.AcuteResponse.GraphUnits = table.Rows[0]["Acute_Response$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.AcuteResponse.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Acute Response value is missing");
                        message += "GraphUnits for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Acute Response value is " + _workFlow6.AcuteResponse.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Acute_Response$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.AcuteResponse.NormalizedGraphUnits = table.Rows[0]["Acute_Response$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Acute Response value is missing");
                        message += "Normalized GraphUnits for Acute Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Acute Response value is " + _workFlow6.AcuteResponse.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Acute_Response$GraphSettingsRequired"))
                {
                    _workFlow6.AcuteResponse.GraphSettingsVerify = table.Rows[0]["Acute_Response$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettingsVerify ? "GraphSettingsVerify for Acute Response is true" : "GraphSettingsVerify for Acute Response is false");
                    if (_workFlow6.AcuteResponse.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Acute_Response$ZeroLine"))
                        {
                            _workFlow6.AcuteResponse.GraphSettings.Zeroline = table.Rows[0]["Acute_Response$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Acute Response is true" : "Zeroline in GraphSettings for Acute Response is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Acute_Response$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.AcuteResponse.CheckNormalizationWithPlateMap = table.Rows[0]["Acute_Response$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.AcuteResponse.IsExportRequired = table.Rows[0]["Acute_Response$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.AcuteResponse.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.AcuteResponse.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.AcuteResponse.IsExportRequired);
                    }
                }
                #endregion

                #region Testid-8

                _workFlow6.ProtonLeak = new WidgetItems();
                _workFlow6.ProtonLeak.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Proton_Leak$Oligo"))
                {
                    _workFlow6.ProtonLeak.Oligo = table.Rows[0]["Proton_Leak$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ProtonLeak.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for Proton Leak is missing");
                        message += "Oligo for Proton Leak&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for Proton Leak is " + _workFlow6.ProtonLeak.Oligo);
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$Display"))
                {
                    _workFlow6.ProtonLeak.Display = table.Rows[0]["Proton_Leak$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ProtonLeak.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Acute Response is missing");
                        message += "Display for Acute Response&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Proton Leak  is " + _workFlow6.ProtonLeak.Display);
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$Normalization"))
                {
                    _workFlow6.ProtonLeak.Normalization = table.Rows[0]["Proton_Leak$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.Normalization ? "Normalization for Proton Leak is true" : "Normalization for Proton Leak is false");
                }

                if (tblcolumns.Contains("Proton_Leak$Error Format"))
                {
                    _workFlow6.ProtonLeak.ErrorFormat = table.Rows[0]["Proton_Leak$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ProtonLeak.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Proton Leak is missing");
                        message += "Error format for Proton Leak&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Proton Leak  is " + _workFlow6.ProtonLeak.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$GraphUnits-OCR"))
                {
                    _workFlow6.ProtonLeak.GraphUnits = table.Rows[0]["Proton_Leak$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ProtonLeak.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Proton Leak value is missing");
                        message += "GraphUnits for Proton Leak&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Proton Leak value is " + _workFlow6.ProtonLeak.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.ProtonLeak.NormalizedGraphUnits = table.Rows[0]["Proton_Leak$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ProtonLeak.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Proton Leak value is missing");
                        message += "Normalized GraphUnits for Proton Leak&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Proton Leak value is " + _workFlow6.ProtonLeak.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$GraphSettingsRequired"))
                {
                    _workFlow6.ProtonLeak.GraphSettingsVerify = table.Rows[0]["Proton_Leak$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettingsVerify ? "GraphSettingsVerify for Proton Leak is true" : "GraphSettingsVerify for Proton Leak is false");
                    if (_workFlow6.ProtonLeak.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Proton_Leak$ZeroLine"))
                        {
                            _workFlow6.ProtonLeak.GraphSettings.Zeroline = table.Rows[0]["Proton_Leak$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Proton Leak is true" : "Zeroline in GraphSettings for Proton Leak is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.ProtonLeak.CheckNormalizationWithPlateMap = table.Rows[0]["Proton_Leak$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.ProtonLeak.IsExportRequired = table.Rows[0]["Proton_Leak$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.ProtonLeak.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.ProtonLeak.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.ProtonLeak.IsExportRequired);
                    }
                }

                #endregion

                #region Testid -9

                _workFlow6.MaximalRespiration = new WidgetItems();
                _workFlow6.MaximalRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Maximal_Respiration$Oligo"))
                {
                    _workFlow6.MaximalRespiration.Oligo = table.Rows[0]["Maximal_Respiration$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for SpareRespiratory Capacity is missing");
                        message += "Oligo for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for SpareRespiratory Capacity  is " + _workFlow6.MaximalRespiration.Oligo);
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$Display"))
                {
                    _workFlow6.MaximalRespiration.Display = table.Rows[0]["Maximal_Respiration$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.Display))
                    {
                        _extentTest.Log(Status.Pass, "Display for Maximal Respiration Capacity is missing");
                        message += "Display for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Maximal Respiration Capacity  is " + _workFlow6.MaximalRespiration.Display);
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$Normalization"))
                {
                    _workFlow6.MaximalRespiration.Normalization = table.Rows[0]["Maximal_Respiration$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.Normalization ? "Normalization for SpareRespiratory Capacity is true" : "Normalization for SpareRespiratory Capacity is false");
                }

                if (tblcolumns.Contains("Maximal_Respiration$Error Format"))
                {
                    _workFlow6.MaximalRespiration.ErrorFormat = table.Rows[0]["Maximal_Respiration$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for SpareRespiratory Capacity is missing");
                        message += "Error format for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for SpareRespiratory Capacity  is " + _workFlow6.MaximalRespiration.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$GraphUnits-OCR"))
                {
                    _workFlow6.MaximalRespiration.GraphUnits = table.Rows[0]["Maximal_Respiration$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Basal Respiration value is missing");
                        message += "GraphUnits for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Basal Respiration value is " + _workFlow6.MaximalRespiration.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.MaximalRespiration.NormalizedGraphUnits = table.Rows[0]["Maximal_Respiration$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.BasalRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Basal Respiration value is missing");
                        message += "Normalized GraphUnits for Basal Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Basal Respiration value is " + _workFlow6.MaximalRespiration.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$GraphSettingsRequired"))
                {
                    _workFlow6.MaximalRespiration.GraphSettingsVerify = table.Rows[0]["Maximal_Respiration$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettingsVerify ? "GraphSettingsVerify for Basal Respiration is true" : "GraphSettingsVerify for Basal Respiration is false");
                    if (_workFlow6.MaximalRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Maximal_Respiration$ZeroLine"))
                        {
                            _workFlow6.MaximalRespiration.GraphSettings.Zeroline = table.Rows[0]["Maximal_Respiration$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Basal Respiration is true" : "Zeroline in GraphSettings for Basal Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.MaximalRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["Maximal_Respiration$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.MaximalRespiration.IsExportRequired = table.Rows[0]["Maximal_Respiration$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.MaximalRespiration.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.MaximalRespiration.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.MaximalRespiration.IsExportRequired);
                    }
                }
                #endregion

                #region Testid -10

                _workFlow6.SpareRespiratoryCapacity = new WidgetItems();
                _workFlow6.SpareRespiratoryCapacity.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Spare_Respiratory$Oligo"))
                {
                    _workFlow6.SpareRespiratoryCapacity.Oligo = table.Rows[0]["Spare_Respiratory$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacity.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for SpareRespiratory Capacity is missing");
                        message += "Default oligo for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.Oligo);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory$Display"))
                {
                    _workFlow6.SpareRespiratoryCapacity.Display = table.Rows[0]["Spare_Respiratory$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacity.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for SpareRespiratory Capacity is missing");
                        message += "Default display for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.Display);
                    }
                }


                if (tblcolumns.Contains("Spare_Respiratory$Normalization"))
                {
                    _workFlow6.SpareRespiratoryCapacity.Normalization = table.Rows[0]["Spare_Respiratory$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.Normalization ? "Dormalization for SpareRespiratory Capacity is true" : "Default normalization for SpareRespiratory Capacity is false");
                }

                if (tblcolumns.Contains("Spare_Respiratory$Error Format"))
                {
                    _workFlow6.SpareRespiratoryCapacity.ErrorFormat = table.Rows[0]["Spare_Respiratory$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacity.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for SpareRespiratory Capacity is missing");
                        message += "Error format for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory$GraphUnits-OCR"))
                {
                    _workFlow6.SpareRespiratoryCapacity.GraphUnits = table.Rows[0]["Spare_Respiratory$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacity.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for SpareRespiratory Capacity value is missing");
                        message += "GraphUnits for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for SpareRespiratory Capacity value is " + _workFlow6.SpareRespiratoryCapacity.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.SpareRespiratoryCapacity.NormalizedGraphUnits = table.Rows[0]["Spare_Respiratory$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for SpareRespiratory Capacity value is missing");
                        message += "Normalized GraphUnits for SpareRespiratory Capacity&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for SpareRespiratory Capacity value is " + _workFlow6.SpareRespiratoryCapacity.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory$GraphSettingsRequired"))
                {
                    _workFlow6.SpareRespiratoryCapacity.GraphSettingsVerify = table.Rows[0]["Spare_Respiratory$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettingsVerify ? "GraphSettingsVerify for SpareRespiratory Capacity is true" : "GraphSettingsVerify for SpareRespiratory Capacity is false");
                    if (_workFlow6.SpareRespiratoryCapacity.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Spare_Respiratory$ZeroLine"))
                        {
                            _workFlow6.SpareRespiratoryCapacity.GraphSettings.Zeroline = table.Rows[0]["Spare_Respiratory$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettings.Zeroline ? "Zeroline in GraphSettings for SpareRespiratory Capacity is true" : "Zeroline in GraphSettings for SpareRespiratory Capacity is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.SpareRespiratoryCapacity.CheckNormalizationWithPlateMap = table.Rows[0]["Spare_Respiratory$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.SpareRespiratoryCapacity.IsExportRequired = table.Rows[0]["Spare_Respiratory$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.SpareRespiratoryCapacity.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.SpareRespiratoryCapacity.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.SpareRespiratoryCapacity.IsExportRequired);
                    }
                }
                #endregion

                #region Testid -11

                _workFlow6.NonmitoO2Consumption = new WidgetItems();
                _workFlow6.NonmitoO2Consumption.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Non_Mitochondrial$Oligo"))
                {
                    _workFlow6.NonmitoO2Consumption.Oligo = table.Rows[0]["Non_Mitochondrial$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.NonmitoO2Consumption.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for Non Mitochondrial Respiration is missing");
                        message += "Oligo for Non Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.Oligo);
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Display"))
                {
                    _workFlow6.NonmitoO2Consumption.Display = table.Rows[0]["Non_Mitochondrial$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.NonmitoO2Consumption.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Non Mitochondrial Respiration is missing");
                        message += "Display for Non Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.Display);
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Normalization"))
                {
                    _workFlow6.NonmitoO2Consumption.Normalization = table.Rows[0]["Non_Mitochondrial$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.Normalization ? "Normalization for Non Mitochondrial Respiration is true" : "Normalization for Non Mitochondrial Respiration is false");
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Error Format"))
                {
                    _workFlow6.NonmitoO2Consumption.ErrorFormat = table.Rows[0]["Non_Mitochondrial$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.NonmitoO2Consumption.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Non Mitochondrial Respiration is missing");
                        message += "Error format for Non Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$GraphUnits-OCR"))
                {
                    _workFlow6.NonmitoO2Consumption.GraphUnits = table.Rows[0]["Non_Mitochondrial$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.NonmitoO2Consumption.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Non Mitochondrial Respiration value is missing");
                        message += "GraphUnits for Non Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Non Mitochondrial Respiration value is " + _workFlow6.NonmitoO2Consumption.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.NonmitoO2Consumption.NormalizedGraphUnits = table.Rows[0]["Non_Mitochondrial$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Non Mitochondrial Respiration value is missing");
                        message += "Normalized GraphUnits for Non Mitochondrial Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Non Mitochondrial Respiration value is " + _workFlow6.NonmitoO2Consumption.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$GraphSettingsRequired"))
                {
                    _workFlow6.NonmitoO2Consumption.GraphSettingsVerify = table.Rows[0]["Non_Mitochondrial$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettingsVerify ? "GraphSettingsVerify for Non Mitochondrial Respiration is true" : "GraphSettingsVerify for Non Mitochondrial Respiration is false");
                    if (_workFlow6.NonmitoO2Consumption.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Non_Mitochondrial$ZeroLine"))
                        {
                            _workFlow6.NonmitoO2Consumption.GraphSettings.Zeroline = table.Rows[0]["Non_Mitochondrial$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Non Mitochondrial Respiration is true" : "Zeroline in GraphSettings for Non Mitochondrial Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.NonmitoO2Consumption.CheckNormalizationWithPlateMap = table.Rows[0]["Non_Mitochondrial$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.NonmitoO2Consumption.IsExportRequired = table.Rows[0]["Non_Mitochondrial$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.NonmitoO2Consumption.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.NonmitoO2Consumption.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.NonmitoO2Consumption.IsExportRequired);
                    }
                }
                #endregion

                #region Testid -12

                _workFlow6.ATPProductionCoupledRespiration = new WidgetItems();
                _workFlow6.ATPProductionCoupledRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("ATP_Production$Oligo"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.Oligo = table.Rows[0]["ATP_Production$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ATPProductionCoupledRespiration.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for ATP-Production Coupled Respiration is missing");
                        message += "Oligo for ATP-Production Coupled Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.Oligo);
                    }
                }

                if (tblcolumns.Contains("ATP_Production$Display"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.Display = table.Rows[0]["ATP_Production$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ATPProductionCoupledRespiration.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for ATP-Production Coupled Respiration is missing");
                        message += "Display for ATP-Production Coupled Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.Display);
                    }
                }

                if (tblcolumns.Contains("ATP_Production$Normalization"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.Normalization = table.Rows[0]["ATP_Production$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.Normalization ? "Normalization for ATP-Production Coupled Respiration is true" : "Normalization for ATP-Production Coupled Respiration is false");
                }

                if (tblcolumns.Contains("ATP_Production$Error Format"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.ErrorFormat = table.Rows[0]["ATP_Production$Default Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ATPProductionCoupledRespiration.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for ATP-Production Coupled Respiration is missing");
                        message += "Error format for ATP-Production Coupled Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("ATP_Production$GraphUnits-OCR"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.GraphUnits = table.Rows[0]["ATP_Production$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.ATPProductionCoupledRespiration.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for ATP-Production Coupled Respiration value is missing");
                        message += "GraphUnits for ATP-Production Coupled Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for ATP-Production Coupled Respiration value is " + _workFlow6.ATPProductionCoupledRespiration.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("ATP_Production$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.NormalizedGraphUnits = table.Rows[0]["ATP_Production$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.MaximalRespiration.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for ATP-Production Coupled Respiration value is missing");
                        message += "Normalized GraphUnits for ATP-Production Coupled Respiration&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for ATP-Production Coupled Respiration value is " + _workFlow6.ATPProductionCoupledRespiration.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("ATP_Production$GraphSettingsRequired"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.GraphSettingsVerify = table.Rows[0]["ATP_Production$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.GraphSettingsVerify ? "GraphSettingsVerify for ATP-Production Coupled Respiration is true" : "GraphSettingsVerify for ATP-Production Coupled Respiration is false");
                    if (_workFlow6.ATPProductionCoupledRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("ATP_Production$ZeroLine"))
                        {
                            _workFlow6.ATPProductionCoupledRespiration.GraphSettings.Zeroline = table.Rows[0]["ATP_Production$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.GraphSettings.Zeroline ? "Zeroline in GraphSettings for ATP-Production Coupled Respiration is true" : "Zeroline in GraphSettings for ATP-Production Coupled Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("ATP_Production$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["ATP_Production$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.ATPProductionCoupledRespiration.IsExportRequired = table.Rows[0]["ATP_Production$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.ATPProductionCoupledRespiration.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.ATPProductionCoupledRespiration.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.ATPProductionCoupledRespiration.IsExportRequired);
                    }
                }
                #endregion

                #region Testid-13

                _workFlow6.CouplingEfficiency = new WidgetItems();
                _workFlow6.CouplingEfficiency.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Coupling_Efficiency$Oligo"))
                {
                    _workFlow6.CouplingEfficiency.Oligo = table.Rows[0]["Coupling_Efficiency$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.CouplingEfficiency.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for Coupling Efficiency is missing");
                        message += "Oligo for Coupling Efficiency&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for Coupling Efficiency is " + _workFlow6.CouplingEfficiency.Oligo);
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Display"))
                {
                    _workFlow6.CouplingEfficiency.Display = table.Rows[0]["Coupling_Efficiency$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.CouplingEfficiency.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Coupling Efficiency is missing");
                        message += "Display for Coupling Efficiency&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Displaymode for Coupling Efficiency  is " + _workFlow6.CouplingEfficiency.Display);
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Normalization"))
                {
                    _workFlow6.CouplingEfficiency.Normalization = table.Rows[0]["Coupling_Efficiency$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.Normalization ? "Normalization for Coupling Efficiency is true" : "Normalization for Coupling Efficiency is false");
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Error Format"))
                {
                    _workFlow6.CouplingEfficiency.ErrorFormat = table.Rows[0]["Coupling_Efficiency$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.CouplingEfficiency.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Coupling Efficiency is missing");
                        message += "Error format for Coupling Efficiency&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Coupling Efficiency  is " + _workFlow6.CouplingEfficiency.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$GraphUnits-OCR"))
                {
                    _workFlow6.CouplingEfficiency.GraphUnits = table.Rows[0]["Coupling_Efficiency$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.CouplingEfficiency.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Coupling Efficiency value is missing");
                        message += "GraphUnits for Coupling Efficiency&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Coupling Efficiency value is " + _workFlow6.CouplingEfficiency.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.CouplingEfficiency.NormalizedGraphUnits = table.Rows[0]["Coupling_Efficiency$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.CouplingEfficiency.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Coupling Efficiency value is missing");
                        message += "Normalized GraphUnits for Coupling Efficiency&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Coupling Efficiency value is " + _workFlow6.CouplingEfficiency.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$GraphSettingsRequired"))
                {
                    _workFlow6.CouplingEfficiency.GraphSettingsVerify = table.Rows[0]["Coupling_Efficiency$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettingsVerify ? "GraphSettingsVerify for Coupling Efficiency is true" : "GraphSettingsVerify for Coupling Efficiency is false");
                    if (_workFlow6.CouplingEfficiency.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Coupling_Efficiency$ZeroLine"))
                        {
                            _workFlow6.CouplingEfficiency.GraphSettings.Zeroline = table.Rows[0]["Coupling_Efficiency$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Coupling Efficiency is true" : "Zeroline in GraphSettings for Coupling Efficiency is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.CouplingEfficiency.CheckNormalizationWithPlateMap = table.Rows[0]["Coupling_Efficiency$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.CouplingEfficiency.IsExportRequired = table.Rows[0]["Coupling_Efficiency$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.CouplingEfficiency.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.CouplingEfficiency.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.CouplingEfficiency.IsExportRequired);
                    }
                }

                #endregion

                #region Testid-14
                _workFlow6.SpareRespiratoryCapacityPercentage = new WidgetItems();
                _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Oligo"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.Oligo = table.Rows[0]["Spare_Respiratory_Capacity$Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacityPercentage.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Oligo for Spare Respiratory Capacity Percentage is missing");
                        message += "Oligo for Spare Respiratory Capacity Percentage&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Oligo for Spare Respiratory Capacity Percentage is " + _workFlow6.SpareRespiratoryCapacityPercentage.Oligo);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Display"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.Display = table.Rows[0]["Spare_Respiratory_Capacity$Display"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacityPercentage.Display))
                    {
                        _extentTest.Log(Status.Fail, "Display for Spare Respiratory Capacity Percentage is missing");
                        message += "Display for Spare Respiratory Capacity Percentage&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Default displaymode for Spare Respiratory Capacity Percentage  is " + _workFlow6.SpareRespiratoryCapacityPercentage.Display);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Normalization"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.Normalization = table.Rows[0]["Spare_Respiratory_Capacity$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.Normalization ? "Normalization for Coupling Efficiency is true" : "Normalization for Coupling Efficiency is false");
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Error Format"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.ErrorFormat = table.Rows[0]["Spare_Respiratory_Capacity$Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacityPercentage.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Error format for Spare Respiratory Capacity Percentage is missing");
                        message += "Error format for Spare Respiratory Capacity Percentage&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Error format for Spare Respiratory Capacity Percentage  is " + _workFlow6.SpareRespiratoryCapacityPercentage.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$GraphUnits-OCR"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.GraphUnits = table.Rows[0]["Spare_Respiratory_Capacity$GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacityPercentage.GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for Spare Respiratory Capacity Percentage value is missing");
                        message += "GraphUnits for Spare Respiratory Capacity Percentage&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for Spare Respiratory Capacity Percentage value is " + _workFlow6.SpareRespiratoryCapacityPercentage.GraphUnits);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Normalized GraphUnits-OCR"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.NormalizedGraphUnits = table.Rows[0]["Spare_Respiratory_Capacity$Normalized GraphUnits-OCR"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.SpareRespiratoryCapacityPercentage.NormalizedGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for Spare Respiratory Capacity Percentage value is missing");
                        message += "Normalized GraphUnits for Spare Respiratory Capacity Percentage&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for Spare Respiratory Capacity Percentage value is " + _workFlow6.SpareRespiratoryCapacityPercentage.NormalizedGraphUnits);
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$GraphSettingsRequired"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettingsVerify = table.Rows[0]["Spare_Respiratory_Capacity$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettingsVerify ? "GraphSettingsVerify for Spare Respiratory Capacity Percentage is true" : "GraphSettingsVerify for Spare Respiratory Capacity Percentage is false");
                    if (_workFlow6.SpareRespiratoryCapacityPercentage.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Spare_Respiratory_Capacity$ZeroLine"))
                        {
                            _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.Zeroline = table.Rows[0]["Spare_Respiratory_Capacity$ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.Zeroline ? "Zeroline in GraphSettings for Spare Respiratory Capacity Percentage is true" : "Zeroline in GraphSettings for Spare Respiratory Capacity Percentage is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.CheckNormalizationWithPlateMap = table.Rows[0]["Spare_Respiratory_Capacity$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");

                    _workFlow6.SpareRespiratoryCapacityPercentage.IsExportRequired = table.Rows[0]["Spare_Respiratory_Capacity$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow6.SpareRespiratoryCapacityPercentage.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.SpareRespiratoryCapacityPercentage.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow6.SpareRespiratoryCapacityPercentage.IsExportRequired);
                    }
                }

                #endregion

                //#region Testid -15

                //_workFlow6.DataTable = new WidgetItems();
                //_workFlow6.DataTable.GraphSettings = new GraphSettings();

                //if (tblcolumns.Contains("Data_Table$Default Oligo"))
                //{
                //    _workFlow6.DataTable.Oligo = table.Rows[0]["Data_Table$Default Oligo"].ToString();
                //    if (string.IsNullOrEmpty(_workFlow6.DataTable.Oligo))
                //    {
                //        _extentTest.Log(Status.Fail, "Default oligo for Data Table is missing");
                //        message += "Default oligo for Data Table&";
                //    }
                //    else
                //    {
                //        _extentTest.Log(Status.Pass, "Default oligo for Data Table  is " + _workFlow6.DataTable.Oligo);
                //    }
                //}

                //if (tblcolumns.Contains("Data_Table$Default Normalization"))
                //{
                //    _workFlow6.DataTable.Normalization = table.Rows[0]["Data_Table$Default Normalization"].ToString() == "ON";
                //    _extentTest.Log(Status.Pass, _workFlow6.DataTable.Normalization ? "Default normalization for Data Table is true" : "Default normalization for Data Table is false");
                //}

                //if (tblcolumns.Contains("Data_Table$Default Error Format"))
                //{
                //    _workFlow6.DataTable.ErrorFormat = table.Rows[0]["Data_Table$Default Error Format"].ToString();
                //    if (string.IsNullOrEmpty(_workFlow6.DataTable.ErrorFormat))
                //    {
                //        _extentTest.Log(Status.Fail, "Default error format for Data Table is missing");
                //        message += "Default error format for Data Table&";
                //    }
                //    else
                //    {
                //        _extentTest.Log(Status.Pass, "Default error format for Data Table  is " + _workFlow6.DataTable.ErrorFormat);
                //    }
                //}

                //if (tblcolumns.Contains("Data_Table$GraphSettingsRequired"))
                //{
                //    _workFlow6.DataTable.GraphSettingsVerify = table.Rows[0]["Data_Table$GraphSettingsRequired"].ToString() == "Yes";
                //    _extentTest.Log(Status.Pass, _workFlow6.DataTable.GraphSettingsVerify ? "GraphSettingsVerify for Data Table is true" : "GraphSettingsVerify for Data Table is false");

                //}
                //_workFlow6.DataTable.IsExportRequired = table.Rows[0]["Data_Table$IsExportRequired"].ToString() == "Yes" ? true : false;
                //#endregion
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

                if (tblcolumns.Contains("CellEnergy_Phenotype$Default Normalization"))
                {
                    _workFlow8.CellEnergyPhenotype.Normalization = table.Rows[0]["CellEnergy_Phenotype$Default Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow8.CellEnergyPhenotype.Normalization ? "Default normalization for CellEnergyPhenotype is true" : "Default normalization for CellEnergyPhenotype is false");
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$Default Error Format"))
                {
                    _workFlow8.CellEnergyPhenotype.ErrorFormat = table.Rows[0]["CellEnergy_Phenotype$Default Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow8.CellEnergyPhenotype.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Default error format for CellEnergyPhenotype is missing");
                        message += "Default error format for CellEnergyPhenotype&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Default error format for CellEnergyPhenotype  is " + _workFlow8.CellEnergyPhenotype.ErrorFormat);
                    }
                }


                if (tblcolumns.Contains("CellEnergy_Phenotype$GraphUnits-OCR"))
                {
                    var GraphUnits = table.Rows[0]["CellEnergy_Phenotype$GraphUnits-OCR"].ToString();
                    _workFlow8.CellEnergyPhenotype.GraphUnits = GraphUnits;
                    if (string.IsNullOrEmpty(GraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "GraphUnits for CellEnergyPhenotype value is missing");
                        message += "GraphUnits for CellEnergyPhenotype&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphUnits for CellEnergyPhenotype value is " + GraphUnits);
                    }
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$Normalized GraphUnits-OCR"))
                {
                    var NormGraphUnits = table.Rows[0]["CellEnergy_Phenotype$Normalized GraphUnits-OCR"].ToString();
                    _workFlow8.CellEnergyPhenotype.NormalizedGraphUnits = NormGraphUnits;
                    if (string.IsNullOrEmpty(NormGraphUnits))
                    {
                        _extentTest.Log(Status.Fail, "Normalized GraphUnits for CellEnergyPhenotype value is missing");
                        message += "Normalized GraphUnits for CellEnergyPhenotype&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Normalized GraphUnits for CellEnergyPhenotype value is " + NormGraphUnits);
                    }
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$IsExportRequired"))
                {
                    _workFlow8.CellEnergyPhenotype.IsExportRequired = table.Rows[0]["CellEnergy_Phenotype$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow8.CellEnergyPhenotype.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "Exports is required.");
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Exports is not required.");
                    }
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

