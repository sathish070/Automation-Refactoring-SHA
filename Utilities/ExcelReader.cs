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
        private readonly WorkFlow6Data _workFlow6;
        private readonly WorkFlow7Data _workFlow7;
        private readonly WorkFlow8Data _workFlow8;
        private readonly string _currentBuildPath;
        private readonly CurrentBrowser _currentBrowser;
        public string PerRate { get; private set; }
        public string PerGraphUnits { get; private set; }
        public string PerNormGraphUnits { get; private set; }
        public static List<string> testidList = new List<string>();
        private readonly FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public ExcelReader(LoginData loginData, FileUploadOrExistingFileData fileUploadOrExistingFileData, NormalizationData normalization, WorkFlow5Data workFlow5, WorkFlow6Data workFlow6, WorkFlow7Data workFlow7, WorkFlow8Data workFlow8,
            string currentBuildPath, CurrentBrowser currentBrowser, ExtentTest extentTest,FilesTabData filesTab)
        {
            _loginData = loginData;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _normalization = normalization;
            _workFlow5 = workFlow5;
            _workFlow6 = workFlow6;
            _workFlow7 = workFlow7;
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
            else if (sheetName == "Workflow-5")
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

                if (tblcolumns.Contains("Normalization_Icon$Normalized File Name"))
                {
                    _workFlow5.NormalizedFileName = table.Rows[0]["Normalization_Icon$Normalized File Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow5.NormalizedFileName))
                    {
                        _extentTest.Log(Status.Fail, "FileName status is false");
                        message += "FileName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The Normalizied file name is " + _workFlow5.NormalizedFileName);
                    }
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
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for Kinetic graph OCR value is " + _workFlow5.KineticGraphOcr.ExpectedGraphUnits);
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

                        if (tblcolumns.Contains("Kinetic_Graph$Remove Data Point Symbols"))
                        {
                            _workFlow5.KineticGraphOcr.GraphSettings.RemoveDataPointSymbols = table.Rows[0]["Kinetic_Graph$Remove Data Point Symbols"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.GraphSettings.RemoveDataPointSymbols ? "Remove Data Point Symbols in GraphSettings for Kinetic graph is true" : "Remove Data Point Symbols in GraphSettings for Kinetic graph is false");
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


                if (tblcolumns.Contains("Kinetic_Graph$IsExportRequired"))
                {
                    _workFlow5.KineticGraphOcr.IsExportRequired = table.Rows[0]["Kinetic_Graph$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.KineticGraphOcr.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                var EcarExpectedGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$Expected GraphUnits-ECAR"))
                {
                    EcarExpectedGraphUnits = table.Rows[0]["Kinetic_Graph$Expected GraphUnits-ECAR"].ToString();
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for Kinetic graph ECAR value is " + EcarExpectedGraphUnits);
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
                    PlateMapSynctoView = _workFlow5.KineticGraphOcr.PlateMapSynctoView,
                    IsExportRequired =  _workFlow5.KineticGraphOcr.IsExportRequired
                };

                var PerExpectedGraphUnits = "";
                if (tblcolumns.Contains("Kinetic_Graph$Expected GraphUnits-PER"))
                {
                    PerExpectedGraphUnits = table.Rows[0]["Kinetic_Graph$Expected GraphUnits-PER"].ToString();
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for Kinetic graph PER value is " + PerExpectedGraphUnits);
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
                    PlateMapSynctoView = _workFlow5.KineticGraphOcr.PlateMapSynctoView,
                    IsExportRequired =  _workFlow5.KineticGraphOcr.IsExportRequired
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
                if (tblcolumns.Contains("Bar_Chart$NonBoxPlotFile"))
                {
                    _workFlow5.Barchart.NonBoxPlotFile = table.Rows[0]["Bar_Chart$NonBoxPlotFile"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Bar Chart  is " + _workFlow5.Barchart.SortBy);
                }
                if (tblcolumns.Contains("Bar_Chart$Expected GraphUnits"))
                {
                    _workFlow5.Barchart.ExpectedGraphUnits = table.Rows[0]["Bar_Chart$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for Barchart is " + _workFlow5.Barchart.ExpectedGraphUnits);
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
                }

                if (tblcolumns.Contains("Bar_Chart$IsExportRequired"))
                {
                    _workFlow5.Barchart.IsExportRequired = table.Rows[0]["Bar_Chart$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow5.Barchart.IsExportRequired ? "Exports is required." : "Exports is not required.");
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
                    _extentTest.Log(Status.Pass, "Expected Graph Units for Energy Map is " + _workFlow5.EnergyMap.ExpectedGraphUnits);
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

                        if (tblcolumns.Contains("Heat_Map$Remove Data Point Symbols"))
                        {
                            _workFlow5.HeatMap.GraphSettings.RemoveDataPointSymbols = table.Rows[0]["Heat_Map$Remove Data Point Symbols"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.HeatMap.GraphSettings.RemoveDataPointSymbols ? "Remove Data Point Symbols in GraphSettings for HeatMap is true" : "Remove Data Point Symbols in GraphSettings for HeatMap is false");
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
                        _extentTest.Log(Status.Pass, "Expected Graph Units for Heat Map is " + _workFlow5.HeatMap.ExpectedGraphUnits);
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

                        if (tblcolumns.Contains("Dose_Response$Remove Data Point Symbols"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDataPointSymbols = table.Rows[0]["Dose_Response$Remove Data Point Symbols"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDataPointSymbols ? "Remove Data Point Symbols in GraphSettings for Dose Kinetic graph is true" : "Remove Data Point Symbols in GraphSettings for Dose Kinetic graph is false");
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

                        if (tblcolumns.Contains("Dose_Response$Remove Dose Data Point Symbols"))
                        {
                            _workFlow5.DoseResponse.GraphSettings.RemoveDoseDataPointSymbols = table.Rows[0]["Dose_Response$Remove Dose Data Point Symbols"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow5.DoseResponse.GraphSettings.RemoveDoseDataPointSymbols ? "Remove Dose Data Point Symbols in GraphSettings for Dose response is true" : "Remove Dose Data Point Symbols in GraphSettings for Dose Response is false");
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
                    _extentTest.Log(Status.Pass, "Expected Graph Units for Dose Response is " + _workFlow5.DoseResponse.ExpectedGraphUnits);
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

                #region TestId -2

                if (tblcolumns.Contains("Layout_Verification$Layout Verification"))
                {
                    _workFlow6.AnalysisLayoutVerification = table.Rows[0]["Layout_Verification$Layout Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AnalysisLayoutVerification ? "Analysis page layout verification is true" : "Analysis page layout verification is false");
                }

                #endregion

                #region TestId -3

                if (tblcolumns.Contains("Normalization_Icon$Normalization Verification"))
                {
                    _workFlow6.NormalizationVerification = table.Rows[0]["Normalization_Icon$Normalization Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NormalizationVerification ? "Normalization Verification is true" : "Normalization Verification is false");
                    ReadDataFromExcel("Normalization");
                }

                if (tblcolumns.Contains("Normalization_Icon$Apply to all widgets"))
                {
                    _workFlow6.ApplyToAllWidgets = table.Rows[0]["Normalization_Icon$Apply to all widgets"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, "ScaleFactor for Barchart is " + _workFlow6.ApplyToAllWidgets);
                }

                if (tblcolumns.Contains("Normalization_Icon$Normalized File Name"))
                {
                    _workFlow6.NormalizedFileName = table.Rows[0]["Normalization_Icon$Normalized File Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.NormalizedFileName))
                    {
                        _extentTest.Log(Status.Fail, "FileName status is false");
                        message += "FileName&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The Normalizied file name is " + _workFlow6.NormalizedFileName);
                    }
                }

                #endregion

                #region TestId -4

                if (tblcolumns.Contains("Modify_Assay$ModifyAssay Verification"))
                {
                    _workFlow6.ModifyAssay = table.Rows[0]["Modify_Assay$ModifyAssay Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ModifyAssay ? "ModifyAssay Verification is true" : "ModifyAssay Verification is false");
                }

                if (tblcolumns.Contains("Modify_Assay$Add Group Name"))
                {
                    _workFlow6.AddGroupName = table.Rows[0]["Modify_Assay$Add Group Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.AddGroupName))
                    {
                        _extentTest.Log(Status.Fail, " Group Name is empty :" + _workFlow6.AddGroupName);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Add group Name is : " + _workFlow5.AddGroupName);
                    }
                }

                if (tblcolumns.Contains("Modify_Assay$Select Controls"))
                {
                    _workFlow6.SelecttheControls = table.Rows[0]["Modify_Assay$Select Controls"].ToString();
                    _extentTest.Log(Status.Pass, "Select the control is : " + _workFlow6.SelecttheControls);
                }

                if (tblcolumns.Contains("Modify_Assay$Injection Name"))
                {
                    _workFlow6.InjectionName = table.Rows[0]["Modify_Assay$Injection Name"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.InjectionName))
                    {
                        _extentTest.Log(Status.Fail, "The Given injection name is :" + _workFlow6.InjectionName);
                        message += "&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "The given injection name is : " + _workFlow6.InjectionName);
                    }
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
                    _extentTest.Log(Status.Pass, "Ratetype for Mitochondrial Respiration is " + _workFlow6.MitochondrialRespiration.Rate);
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Display"))
                {
                    _workFlow6.MitochondrialRespiration.Display = table.Rows[0]["Mitochondrial_Respiration$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Mitochondrial Respiration  is " + _workFlow6.MitochondrialRespiration.Display);
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Y"))
                {
                    _workFlow6.MitochondrialRespiration.Y = table.Rows[0]["Mitochondrial_Respiration$Y"].ToString();
                    _extentTest.Log(Status.Pass, "Y-toggle for Mitochondrial Respiration  is " + _workFlow6.MitochondrialRespiration.Y);
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Normalization"))
                {
                    _workFlow6.MitochondrialRespiration.Normalization = table.Rows[0]["Mitochondrial_Respiration$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.Normalization ? "Normalization for Mitochiondrial Respiration is true" : "Normalization for Mitochiondrial Respiration is false");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Error Format"))
                {
                    _workFlow6.MitochondrialRespiration.ErrorFormat = table.Rows[0]["Mitochondrial_Respiration$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Mitochondrial Respiration  is " + _workFlow6.MitochondrialRespiration.ErrorFormat);
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Background Correction"))
                {
                    _workFlow6.MitochondrialRespiration.BackgroundCorrection = table.Rows[0]["Mitochondrial_Respiration$Background Correction"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.BackgroundCorrection ? "Background correction for Mitochondrial Respiration is true" : "Background for Mitochondrial Respiration is false");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Baseline"))
                {
                    _workFlow6.MitochondrialRespiration.Baseline = table.Rows[0]["Mitochondrial_Respiration$Baseline"].ToString();
                    _extentTest.Log(Status.Pass, "Baseline for Mitochiondrial Respiration  is " + _workFlow6.MitochondrialRespiration.Baseline);
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$Expected GraphUnits"))
                {
                    _workFlow6.MitochondrialRespiration.ExpectedGraphUnits = table.Rows[0]["Mitochondrial_Respiration$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Mitochondrial Respiration value is " + _workFlow6.MitochondrialRespiration.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$GraphSettingsRequired"))
                {
                    _workFlow6.MitochondrialRespiration.GraphSettingsVerify = table.Rows[0]["Mitochondrial_Respiration$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettingsVerify ? "GraphSettingsVerify for Mitochondrial Respiration is true" : "GraphSettingsVerify for Mitochondrial Respiration is false");
                    if (_workFlow6.MitochondrialRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Mitochondrial_Respiration$Remove Y AutoScale"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RemoveYAutoScale = table.Rows[0]["Mitochondrial_Respiration$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Mitochondrial Respiration is true" : "Remove Y AutoScale in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$Remove ZeroLine"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RemoveZeroLine = table.Rows[0]["Mitochondrial_Respiration$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RemoveZeroLine ? "RemoveZeroLine in GraphSettings for Mitochondrial Respiration is true" : "RemoveZeroLine in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$Remove Data Point Symbols"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RemoveDataPointSymbols = table.Rows[0]["Mitochondrial_Respiration$Remove Data Point Symbols"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RemoveDataPointSymbols ? "Data Points symbols in GraphSettings for Mitochondrial Respiration is true" : "Data Points symbols in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$Remove RateHighlight"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RemoveRateHighlight = table.Rows[0]["Mitochondrial_Respiration$Remove RateHighlight"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RemoveRateHighlight ? "RateHighlight in GraphSettings for Mitochondrial Respiration is true" : "RateHighlight in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$Remove InjectionMakers"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RemoveInjectionMarkers = table.Rows[0]["Mitochondrial_Respiration$Remove InjectionMakers"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RemoveInjectionMarkers ? "InjectionMakers in GraphSettings for Mitochondrial Respiration is true" : "InjectionMakers in GraphSettings for Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Mitochondrial_Respiration$Remove Zoom"))
                        {
                            _workFlow6.MitochondrialRespiration.GraphSettings.RemoveZoom = table.Rows[0]["Mitochondrial_Respiration$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Mitochondrial Respiration is true" : "Remove Zoom in GraphSettings for  Mitochondrial Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.MitochondrialRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["Mitochondrial_Respiration$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$PlateMap Sync to View"))
                {
                    _workFlow6.MitochondrialRespiration.PlateMapSynctoView = table.Rows[0]["Mitochondrial_Respiration$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$GraphSettings Sync to View"))
                {
                    _workFlow6.MitochondrialRespiration.GraphSettings.SynctoView = table.Rows[0]["Mitochondrial_Respiration$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Mitochondrial_Respiration$IsExportRequired"))
                {
                    _workFlow6.MitochondrialRespiration.IsExportRequired = table.Rows[0]["Mitochondrial_Respiration$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MitochondrialRespiration.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId -6

                _workFlow6.BasalRespiration = new WidgetItems();
                _workFlow6.BasalRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Basal_Respiration$Oligo"))
                {
                    _workFlow6.BasalRespiration.Oligo = table.Rows[0]["Basal_Respiration$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for Basal Respiration  is " + _workFlow6.BasalRespiration.Oligo);
                }

                if (tblcolumns.Contains("Basal_Respiration$Display"))
                {
                    _workFlow6.BasalRespiration.Display = table.Rows[0]["Basal_Respiration$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Basal Respiration  is " + _workFlow6.BasalRespiration.Display);
                }

                if (tblcolumns.Contains("Basal_Respiration$Normalization"))
                {
                    _workFlow6.BasalRespiration.Normalization = table.Rows[0]["Basal_Respiration$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.Normalization ? "Normalization for Basal Respiration is true" : "Normalization for Basal Respiration is false");
                }

                if (tblcolumns.Contains("Basal_Respiration$Error Format"))
                {
                    _workFlow6.BasalRespiration.ErrorFormat = table.Rows[0]["Basal_Respiration$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Basal Respiration  is " + _workFlow6.BasalRespiration.ErrorFormat);
                }

                if (tblcolumns.Contains("Basal_Respiration$Sort By"))
                {
                    _workFlow6.BasalRespiration.SortBy = table.Rows[0]["Basal_Respiration$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Basal Respiration  is " + _workFlow6.BasalRespiration.SortBy);
                }

                if (tblcolumns.Contains("Basal_Respiration$Expected GraphUnits"))
                {
                    _workFlow6.BasalRespiration.ExpectedGraphUnits = table.Rows[0]["Basal_Respiration$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Basal Respiration value is " + _workFlow6.BasalRespiration.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Basal_Respiration$GraphSettingsRequired"))
                {
                    _workFlow6.BasalRespiration.GraphSettingsVerify = table.Rows[0]["Basal_Respiration$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettingsVerify ? "GraphSettingsVerify for Basal Respiration is true" : "GraphSettingsVerify for Basal Respiration is false");
                    if (_workFlow6.BasalRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Basal_Respiration$Remove Y AutoScale"))
                        {
                            _workFlow6.BasalRespiration.GraphSettings.RemoveYAutoScale = table.Rows[0]["Basal_Respiration$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Basal Respiration is true" : "Remove Y AutoScale in GraphSettings for Basal Respiration is false");
                        }

                        if (tblcolumns.Contains("Basal_Respiration$Remove ZeroLine"))
                        {
                            _workFlow6.BasalRespiration.GraphSettings.RemoveZeroLine = table.Rows[0]["Basal_Respiration$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettings.RemoveZeroLine ? "RemoveZeroLine in GraphSettings for Basal Respiration is true" : "RemoveZeroLine in GraphSettings for Basal Respiration is false");
                        }

                        if (tblcolumns.Contains("Basal_Respiration$Remove Zoom"))
                        {
                            _workFlow6.BasalRespiration.GraphSettings.RemoveZoom = table.Rows[0]["Basal_Respiration$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Basal Respiration is true" : "Remove Zoom in GraphSettings for  Basal Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Basal_Respiration$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.BasalRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["Basal_Respiration$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Basal_Respiration$PlateMap Sync to View"))
                {
                    _workFlow6.BasalRespiration.PlateMapSynctoView = table.Rows[0]["Basal_Respiration$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Basal_Respiration$GraphSettings Sync to View"))
                {
                    _workFlow6.BasalRespiration.GraphSettings.SynctoView = table.Rows[0]["Basal_Respiration$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Basal_Respiration$IsExportRequired"))
                {
                    _workFlow6.BasalRespiration.IsExportRequired = table.Rows[0]["Basal_Respiration$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.BasalRespiration.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region Testid -7

                _workFlow6.AcuteResponse = new WidgetItems();
                _workFlow6.AcuteResponse.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Acute_Response$Display"))
                {
                    _workFlow6.AcuteResponse.Display = table.Rows[0]["Acute_Response$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Acute Response  is " + _workFlow6.AcuteResponse.Display);
                }

                if (tblcolumns.Contains("Acute_Response$Normalization"))
                {
                    _workFlow6.AcuteResponse.Normalization = table.Rows[0]["Acute_Response$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.Normalization ? "Normalization for Acute Response is true" : "Normalization for Acute Response is false");
                }

                if (tblcolumns.Contains("Acute_Response$Error Format"))
                {
                    _workFlow6.AcuteResponse.ErrorFormat = table.Rows[0]["Acute_Response$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Acute Response  is " + _workFlow6.AcuteResponse.ErrorFormat);
                }

                if (tblcolumns.Contains("Acute_Response$Sort By"))
                {
                    _workFlow6.AcuteResponse.SortBy = table.Rows[0]["Acute_Response$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Acute Response  is " + _workFlow6.AcuteResponse.SortBy);
                }

                if (tblcolumns.Contains("Acute_Response$Expected GraphUnits"))
                {
                    _workFlow6.AcuteResponse.ExpectedGraphUnits = table.Rows[0]["Acute_Response$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Acute Response value is " + _workFlow6.AcuteResponse.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Acute_Response$GraphSettingsRequired"))
                {
                    _workFlow6.AcuteResponse.GraphSettingsVerify = table.Rows[0]["Acute_Response$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettingsVerify ? "GraphSettingsVerify for Acute Response is true" : "GraphSettingsVerify for Acute Response is false");
                    if (_workFlow6.AcuteResponse.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Acute_Response$Remove Y AutoScale"))
                        {
                            _workFlow6.BasalRespiration.GraphSettings.RemoveYAutoScale = table.Rows[0]["Acute_Response$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Acute Response is true" : "Remove Y AutoScale in GraphSettings for Acute Response is false");
                        }

                        if (tblcolumns.Contains("Acute_Response$Remove ZeroLine"))
                        {
                            _workFlow6.AcuteResponse.GraphSettings.RemoveZeroLine = table.Rows[0]["Acute_Response$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettings.RemoveZeroLine ? "RemoveZeroLine in GraphSettings for Acute Response is true" : "RemoveZeroLine in GraphSettings for Acute Response is false");
                        }

                        if (tblcolumns.Contains("Acute_Response$Remove Zoom"))
                        {
                            _workFlow6.BasalRespiration.GraphSettings.RemoveZoom = table.Rows[0]["Acute_Response$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Acute Response is true" : "Remove Zoom in GraphSettings for Acute Response is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Acute_Response$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.AcuteResponse.CheckNormalizationWithPlateMap = table.Rows[0]["Acute_Response$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Acute_Response$PlateMap Sync to View"))
                {
                    _workFlow6.AcuteResponse.PlateMapSynctoView = table.Rows[0]["Acute_Response$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Acute_Response$GraphSettings Sync to View"))
                {
                    _workFlow6.AcuteResponse.GraphSettings.SynctoView = table.Rows[0]["Acute_Response$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Acute_Response$IsExportRequired"))
                {
                    _workFlow6.AcuteResponse.IsExportRequired = table.Rows[0]["Acute_Response$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.AcuteResponse.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region Testid-8

                _workFlow6.ProtonLeak = new WidgetItems();
                _workFlow6.ProtonLeak.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Proton_Leak$Oligo"))
                {
                    _workFlow6.ProtonLeak.Oligo = table.Rows[0]["Proton_Leak$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for Proton Leak is " + _workFlow6.ProtonLeak.Oligo);
                }

                if (tblcolumns.Contains("Proton_Leak$Display"))
                {
                    _workFlow6.ProtonLeak.Display = table.Rows[0]["Proton_Leak$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Proton Leak  is " + _workFlow6.ProtonLeak.Display);
                }

                if (tblcolumns.Contains("Proton_Leak$Normalization"))
                {
                    _workFlow6.ProtonLeak.Normalization = table.Rows[0]["Proton_Leak$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.Normalization ? "Normalization for Proton Leak is true" : "Normalization for Proton Leak is false");
                }

                if (tblcolumns.Contains("Proton_Leak$Error Format"))
                {
                    _workFlow6.ProtonLeak.ErrorFormat = table.Rows[0]["Proton_Leak$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Proton Leak  is " + _workFlow6.ProtonLeak.ErrorFormat);
                }

                if (tblcolumns.Contains("Proton_Leak$Sort By"))
                {
                    _workFlow6.ProtonLeak.SortBy = table.Rows[0]["Proton_Leak$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Proton Leak  is " + _workFlow6.ProtonLeak.SortBy);
                }

                if (tblcolumns.Contains("Proton_Leak$Expected GraphUnits"))
                {
                    _workFlow6.ProtonLeak.ExpectedGraphUnits = table.Rows[0]["Proton_Leak$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Proton Leak value is " + _workFlow6.ProtonLeak.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Proton_Leak$GraphSettingsRequired"))
                {
                    _workFlow6.ProtonLeak.GraphSettingsVerify = table.Rows[0]["Proton_Leak$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettingsVerify ? "GraphSettingsVerify for Proton Leak is true" : "GraphSettingsVerify for Proton Leak is false");
                    if (_workFlow6.ProtonLeak.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Proton_Leak$Remove Y AutoScale"))
                        {
                            _workFlow6.ProtonLeak.GraphSettings.RemoveYAutoScale = table.Rows[0]["Proton_Leak$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Proton Leak is true" : "Remove Y AutoScale in GraphSettings for Proton Leak is false");
                        }

                        if (tblcolumns.Contains("Proton_Leak$Remove ZeroLine"))
                        {
                            _workFlow6.ProtonLeak.GraphSettings.RemoveZeroLine = table.Rows[0]["Proton_Leak$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettings.RemoveZeroLine ? "Remove ZeroLine in GraphSettings for Proton Leak is true" : "RemoveZeroLine in GraphSettings for Proton Leak is false");
                        }

                        if (tblcolumns.Contains("Proton_Leak$Remove Zoom"))
                        {
                            _workFlow6.ProtonLeak.GraphSettings.RemoveZoom = table.Rows[0]["Proton_Leak$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Proton Leak is true" : "Remove Zoom in GraphSettings for Proton Leak is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Proton_Leak$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.ProtonLeak.CheckNormalizationWithPlateMap = table.Rows[0]["Proton_Leak$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Proton_Leak$PlateMap Sync to View"))
                {
                    _workFlow6.ProtonLeak.PlateMapSynctoView = table.Rows[0]["Proton_Leak$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Proton_Leak$GraphSettings Sync to View"))
                {
                    _workFlow6.ProtonLeak.GraphSettings.SynctoView = table.Rows[0]["Proton_Leak$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Proton_Leak$IsExportRequired"))
                {
                    _workFlow6.ProtonLeak.IsExportRequired = table.Rows[0]["Proton_Leak$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ProtonLeak.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region Testid -9

                _workFlow6.MaximalRespiration = new WidgetItems();
                _workFlow6.MaximalRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Maximal_Respiration$Oligo"))
                {
                    _workFlow6.MaximalRespiration.Oligo = table.Rows[0]["Maximal_Respiration$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for  Maximal Respiration  is " + _workFlow6.MaximalRespiration.Oligo);
                }

                if (tblcolumns.Contains("Maximal_Respiration$Display"))
                {
                    _workFlow6.MaximalRespiration.Display = table.Rows[0]["Maximal_Respiration$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Maximal Respiration Capacity  is " + _workFlow6.MaximalRespiration.Display);
                }

                if (tblcolumns.Contains("Maximal_Respiration$Normalization"))
                {
                    _workFlow6.MaximalRespiration.Normalization = table.Rows[0]["Maximal_Respiration$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.Normalization ? "Normalization for Maximal Respiration is true" : "Normalization for Maximal Respiration is false");
                }

                if (tblcolumns.Contains("Maximal_Respiration$Error Format"))
                {
                    _workFlow6.MaximalRespiration.ErrorFormat = table.Rows[0]["Maximal_Respiration$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Maximal Respiration  is " + _workFlow6.MaximalRespiration.ErrorFormat);
                }

                if (tblcolumns.Contains("Maximal_Respiration$Sort By"))
                {
                    _workFlow6.MaximalRespiration.SortBy = table.Rows[0]["Maximal_Respiration$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Maximal Respiration  is " + _workFlow6.MaximalRespiration.SortBy);
                }

                if (tblcolumns.Contains("Maximal_Respiration$Expected GraphUnits"))
                {
                    _workFlow6.MaximalRespiration.ExpectedGraphUnits = table.Rows[0]["Maximal_Respiration$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Basal Respiration value is " + _workFlow6.MaximalRespiration.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Maximal_Respiration$GraphSettingsRequired"))
                {
                    _workFlow6.MaximalRespiration.GraphSettingsVerify = table.Rows[0]["Maximal_Respiration$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettingsVerify ? "GraphSettingsVerify for Maximal Respiration is true" : "GraphSettingsVerify for Maximal Respiration is false");
                    if (_workFlow6.MaximalRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Maximal_Respiration$Remove Y AutoScale"))
                        {
                            _workFlow6.MaximalRespiration.GraphSettings.RemoveYAutoScale = table.Rows[0]["Maximal_Respiration$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for  Maximal Respiration is true" : "Remove Y AutoScale in GraphSettings for  Maximal Respiration is false");
                        }

                        if (tblcolumns.Contains("Maximal_Respiration$Remove ZeroLine"))
                        {
                            _workFlow6.MaximalRespiration.GraphSettings.RemoveZeroLine = table.Rows[0]["Maximal_Respiration$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for  Maximal Respiration is true" : "Zeroline in GraphSettings for  Maximal Respiration is false");
                        }

                        if (tblcolumns.Contains("Maximal_Respiration$Remove Zoom"))
                        {
                            _workFlow6.MaximalRespiration.GraphSettings.RemoveZoom = table.Rows[0]["Maximal_Respiration$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for  Maximal Respiration is true" : "Remove Zoom in GraphSettings for  Maximal Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Maximal_Respiration$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.MaximalRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["Maximal_Respiration$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Maximal_Respiration$PlateMap Sync to View"))
                {
                    _workFlow6.MaximalRespiration.PlateMapSynctoView = table.Rows[0]["Maximal_Respiration$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Maximal_Respiration$GraphSettings Sync to View"))
                {
                    _workFlow6.MaximalRespiration.GraphSettings.SynctoView = table.Rows[0]["Maximal_Respiration$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Maximal_Respiration$IsExportRequired"))
                {
                    _workFlow6.MaximalRespiration.IsExportRequired = table.Rows[0]["Maximal_Respiration$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.MaximalRespiration.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region Testid -10

                _workFlow6.SpareRespiratoryCapacity = new WidgetItems();
                _workFlow6.SpareRespiratoryCapacity.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Spare_Respiratory$Oligo"))
                {
                    _workFlow6.SpareRespiratoryCapacity.Oligo = table.Rows[0]["Spare_Respiratory$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.Oligo);
                }

                if (tblcolumns.Contains("Spare_Respiratory$Display"))
                {
                    _workFlow6.SpareRespiratoryCapacity.Display = table.Rows[0]["Spare_Respiratory$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.Display);
                }

                if (tblcolumns.Contains("Spare_Respiratory$Normalization"))
                {
                    _workFlow6.SpareRespiratoryCapacity.Normalization = table.Rows[0]["Spare_Respiratory$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.Normalization ? "Dormalization for SpareRespiratory Capacity is true" : "Default normalization for SpareRespiratory Capacity is false");
                }

                if (tblcolumns.Contains("Spare_Respiratory$Error Format"))
                {
                    _workFlow6.SpareRespiratoryCapacity.ErrorFormat = table.Rows[0]["Spare_Respiratory$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.ErrorFormat);
                }

                if (tblcolumns.Contains("Spare_Respiratory$Sort By"))
                {
                    _workFlow6.SpareRespiratoryCapacity.SortBy = table.Rows[0]["Spare_Respiratory$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for SpareRespiratory Capacity  is " + _workFlow6.SpareRespiratoryCapacity.SortBy);
                }

                if (tblcolumns.Contains("Spare_Respiratory$Expected GraphUnits"))
                {
                    _workFlow6.SpareRespiratoryCapacity.ExpectedGraphUnits = table.Rows[0]["Spare_Respiratory$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for SpareRespiratory Capacity value is " + _workFlow6.SpareRespiratoryCapacity.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Spare_Respiratory$GraphSettingsRequired"))
                {
                    _workFlow6.SpareRespiratoryCapacity.GraphSettingsVerify = table.Rows[0]["Spare_Respiratory$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettingsVerify ? "GraphSettingsVerify for SpareRespiratory Capacity is true" : "GraphSettingsVerify for SpareRespiratory Capacity is false");
                    if (_workFlow6.SpareRespiratoryCapacity.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Spare_Respiratory$Remove Y AutoScale"))
                        {
                            _workFlow6.SpareRespiratoryCapacity.GraphSettings.RemoveYAutoScale = table.Rows[0]["Spare_Respiratory$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for SpareRespiratory Capacity is true" : "Remove Y AutoScale in GraphSettings for SpareRespiratory Capacity is false");
                        }

                        if (tblcolumns.Contains("Spare_Respiratory$Remove ZeroLine"))
                        {
                            _workFlow6.SpareRespiratoryCapacity.GraphSettings.RemoveZeroLine = table.Rows[0]["Spare_Respiratory$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for SpareRespiratory Capacity is true" : "Zeroline in GraphSettings for SpareRespiratory Capacity is false");
                        }

                        if (tblcolumns.Contains("Spare_Respiratory$Remove Zoom"))
                        {
                            _workFlow6.SpareRespiratoryCapacity.GraphSettings.RemoveZoom = table.Rows[0]["Spare_Respiratory$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for SpareRespiratory Capacity is true" : "Remove Zoom in GraphSettings for SpareRespiratory Capacity is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.SpareRespiratoryCapacity.CheckNormalizationWithPlateMap = table.Rows[0]["Spare_Respiratory$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Spare_Respiratory$PlateMap Sync to View"))
                {
                    _workFlow6.SpareRespiratoryCapacity.PlateMapSynctoView = table.Rows[0]["Spare_Respiratory$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Spare_Respiratory$GraphSettings Sync to View"))
                {
                    _workFlow6.SpareRespiratoryCapacity.GraphSettings.SynctoView = table.Rows[0]["Spare_Respiratory$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Spare_Respiratory$IsExportRequired"))
                {
                    _workFlow6.SpareRespiratoryCapacity.IsExportRequired = table.Rows[0]["Spare_Respiratory$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacity.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region Testid -11

                _workFlow6.NonmitoO2Consumption = new WidgetItems();
                _workFlow6.NonmitoO2Consumption.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Non_Mitochondrial$Oligo"))
                {
                    _workFlow6.NonmitoO2Consumption.Oligo = table.Rows[0]["Non_Mitochondrial$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.Oligo);
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Display"))
                {
                    _workFlow6.NonmitoO2Consumption.Display = table.Rows[0]["Non_Mitochondrial$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.Display);
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Normalization"))
                {
                    _workFlow6.NonmitoO2Consumption.Normalization = table.Rows[0]["Non_Mitochondrial$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.Normalization ? "Normalization for Non Mitochondrial Respiration is true" : "Normalization for Non Mitochondrial Respiration is false");
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Error Format"))
                {
                    _workFlow6.NonmitoO2Consumption.ErrorFormat = table.Rows[0]["Non_Mitochondrial$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.ErrorFormat);
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Sort By"))
                {
                    _workFlow6.NonmitoO2Consumption.SortBy = table.Rows[0]["Non_Mitochondrial$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for  Non Mitochondrial Respiration  is " + _workFlow6.NonmitoO2Consumption.SortBy);
                }

                if (tblcolumns.Contains("Non_Mitochondrial$Expected GraphUnits"))
                {
                    _workFlow6.NonmitoO2Consumption.ExpectedGraphUnits = table.Rows[0]["Non_Mitochondrial$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Non Mitochondrial Respiration value is " + _workFlow6.NonmitoO2Consumption.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Non_Mitochondrial$GraphSettingsRequired"))
                {
                    _workFlow6.NonmitoO2Consumption.GraphSettingsVerify = table.Rows[0]["Non_Mitochondrial$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettingsVerify ? "GraphSettingsVerify for Non Mitochondrial Respiration is true" : "GraphSettingsVerify for Non Mitochondrial Respiration is false");
                    if (_workFlow6.NonmitoO2Consumption.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Non_Mitochondrial$Remove Y AutoScale"))
                        {
                            _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveYAutoScale = table.Rows[0]["Non_Mitochondrial$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Non Mitochondrial Respiration is true" : "Remove Y AutoScale in GraphSettings for Non Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Non_Mitochondrial$Remove ZeroLine"))
                        {
                            _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveZeroLine = table.Rows[0]["Non_Mitochondrial$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for Non Mitochondrial Respiration is true" : "Zeroline in GraphSettings for Non Mitochondrial Respiration is false");
                        }

                        if (tblcolumns.Contains("Non_Mitochondrial$Remove Zoom"))
                        {
                            _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveZoom = table.Rows[0]["Non_Mitochondrial$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Non Mitochondrial Respiration is true" : "Remove Zoom in GraphSettings for Non Mitochondrial Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Non_Mitochondrial$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.NonmitoO2Consumption.CheckNormalizationWithPlateMap = table.Rows[0]["Non_Mitochondrial$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Non_Mitochondrial$PlateMap Sync to View"))
                {
                    _workFlow6.NonmitoO2Consumption.PlateMapSynctoView = table.Rows[0]["Non_Mitochondrial$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Non_Mitochondrial$GraphSettings Sync to View"))
                {
                    _workFlow6.NonmitoO2Consumption.GraphSettings.SynctoView = table.Rows[0]["Non_Mitochondrial$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Non_Mitochondrial$IsExportRequired"))
                {
                    _workFlow6.NonmitoO2Consumption.IsExportRequired = table.Rows[0]["Non_Mitochondrial$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region Testid -12

                _workFlow6.ATPProductionCoupledRespiration = new WidgetItems();
                _workFlow6.ATPProductionCoupledRespiration.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("ATP_Production$Oligo"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.Oligo = table.Rows[0]["ATP_Production$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.Oligo);
                }

                if (tblcolumns.Contains("ATP_Production$Display"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.Display = table.Rows[0]["ATP_Production$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.Display);
                }

                if (tblcolumns.Contains("ATP_Production$Normalization"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.Normalization = table.Rows[0]["ATP_Production$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.Normalization ? "Normalization for ATP-Production Coupled Respiration is true" : "Normalization for ATP-Production Coupled Respiration is false");
                }

                if (tblcolumns.Contains("ATP_Production$Error Format"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.ErrorFormat = table.Rows[0]["ATP_Production$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.ErrorFormat);
                }

                if (tblcolumns.Contains("ATP_Production$Sort By"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.SortBy = table.Rows[0]["ATP_Production$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for ATP-Production Coupled Respiration  is " + _workFlow6.ATPProductionCoupledRespiration.SortBy);
                }

                if (tblcolumns.Contains("ATP_Production$Expected GraphUnits"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.ExpectedGraphUnits = table.Rows[0]["ATP_Production$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for ATP-Production Coupled Respiration value is " + _workFlow6.ATPProductionCoupledRespiration.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("ATP_Production$GraphSettingsRequired"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.GraphSettingsVerify = table.Rows[0]["ATP_Production$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.GraphSettingsVerify ? "GraphSettingsVerify for ATP-Production Coupled Respiration is true" : "GraphSettingsVerify for ATP-Production Coupled Respiration is false");
                    if (_workFlow6.ATPProductionCoupledRespiration.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("ATP_Production$Remove Y AutoScale"))
                        {
                            _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveYAutoScale = table.Rows[0]["ATP_Production$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for ATP-Production Coupled Respiration is true" : "Remove Y AutoScale in GraphSettings for ATP-Production Coupled Respiration is false");
                        }

                        if (tblcolumns.Contains("ATP_Production$Remove ZeroLine"))
                        {
                            _workFlow6.ATPProductionCoupledRespiration.GraphSettings.RemoveZeroLine = table.Rows[0]["ATP_Production$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for ATP-Production Coupled Respiration is true" : "Zeroline in GraphSettings for ATP-Production Coupled Respiration is false");
                        }

                        if (tblcolumns.Contains("ATP_Production$Remove Zoom"))
                        {
                            _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveZoom = table.Rows[0]["ATP_Production$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.NonmitoO2Consumption.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for ATP-Production Coupled Respiration is true" : "Remove Zoom in GraphSettings for ATP-Production Coupled Respiration is false");
                        }
                    }
                }

                if (tblcolumns.Contains("ATP_Production$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.CheckNormalizationWithPlateMap = table.Rows[0]["ATP_Production$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_Production$PlateMap Sync to View"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.PlateMapSynctoView = table.Rows[0]["ATP_Production$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_Production$GraphSettings Sync to View"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.GraphSettings.SynctoView = table.Rows[0]["ATP_Production$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_Production$IsExportRequired"))
                {
                    _workFlow6.ATPProductionCoupledRespiration.IsExportRequired = table.Rows[0]["ATP_Production$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.ATPProductionCoupledRespiration.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region Testid-13

                _workFlow6.CouplingEfficiency = new WidgetItems();
                _workFlow6.CouplingEfficiency.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Coupling_Efficiency$Oligo"))
                {
                    _workFlow6.CouplingEfficiency.Oligo = table.Rows[0]["Coupling_Efficiency$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for Coupling Efficiency is " + _workFlow6.CouplingEfficiency.Oligo);
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Display"))
                {
                    _workFlow6.CouplingEfficiency.Display = table.Rows[0]["Coupling_Efficiency$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for Coupling Efficiency  is " + _workFlow6.CouplingEfficiency.Display);
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Normalization"))
                {
                    _workFlow6.CouplingEfficiency.Normalization = table.Rows[0]["Coupling_Efficiency$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.Normalization ? "Normalization for Coupling Efficiency is true" : "Normalization for Coupling Efficiency is false");
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Error Format"))
                {
                    _workFlow6.CouplingEfficiency.ErrorFormat = table.Rows[0]["Coupling_Efficiency$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Coupling Efficiency  is " + _workFlow6.CouplingEfficiency.ErrorFormat);
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Sort By"))
                {
                    _workFlow6.CouplingEfficiency.SortBy = table.Rows[0]["Coupling_Efficiency$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Coupling Efficiency  is " + _workFlow6.CouplingEfficiency.SortBy);
                }

                if (tblcolumns.Contains("Coupling_Efficiency$Expected GraphUnits"))
                {
                    _workFlow6.CouplingEfficiency.ExpectedGraphUnits = table.Rows[0]["Coupling_Efficiency$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Coupling Efficiency value is " + _workFlow6.CouplingEfficiency.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Coupling_Efficiency$GraphSettingsRequired"))
                {
                    _workFlow6.CouplingEfficiency.GraphSettingsVerify = table.Rows[0]["Coupling_Efficiency$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettingsVerify ? "GraphSettingsVerify for Coupling Efficiency is true" : "GraphSettingsVerify for Coupling Efficiency is false");
                    if (_workFlow6.CouplingEfficiency.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Coupling_Efficiency$Remove Y AutoScale"))
                        {
                            _workFlow6.CouplingEfficiency.GraphSettings.RemoveYAutoScale = table.Rows[0]["Coupling_Efficiency$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Coupling Efficiency is true" : "Remove Y AutoScale in GraphSettings for Coupling Efficiency is false");
                        }

                        if (tblcolumns.Contains("Coupling_Efficiency$Remove ZeroLine"))
                        {
                            _workFlow6.CouplingEfficiency.GraphSettings.RemoveZeroLine = table.Rows[0]["Coupling_Efficiency$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for Coupling Efficiency is true" : "Zeroline in GraphSettings for Coupling Efficiency is false");
                        }

                        if (tblcolumns.Contains("Coupling_Efficiency$Remove Zoom"))
                        {
                            _workFlow6.CouplingEfficiency.GraphSettings.RemoveZoom = table.Rows[0]["Coupling_Efficiency$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Coupling Efficiency is true" : "Remove Zoom in GraphSettings for Coupling Efficiency is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Coupling_Efficiency$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.CouplingEfficiency.CheckNormalizationWithPlateMap = table.Rows[0]["Coupling_Efficiency$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Coupling_Efficiency$PlateMap Sync to View"))
                {
                    _workFlow6.CouplingEfficiency.PlateMapSynctoView = table.Rows[0]["Coupling_Efficiency$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Coupling_Efficiency$GraphSettings Sync to View"))
                {
                    _workFlow6.CouplingEfficiency.GraphSettings.SynctoView = table.Rows[0]["Coupling_Efficiency$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Coupling_Efficiency$IsExportRequired"))
                {
                    _workFlow6.CouplingEfficiency.IsExportRequired = table.Rows[0]["Coupling_Efficiency$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.CouplingEfficiency.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region Testid-14
                _workFlow6.SpareRespiratoryCapacityPercentage = new WidgetItems();
                _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Oligo"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.Oligo = table.Rows[0]["Spare_Respiratory_Capacity$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for Spare Respiratory Capacity Percentage is " + _workFlow6.SpareRespiratoryCapacityPercentage.Oligo);
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Display"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.Display = table.Rows[0]["Spare_Respiratory_Capacity$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Default displaymode for Spare Respiratory Capacity Percentage  is " + _workFlow6.SpareRespiratoryCapacityPercentage.Display);
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Normalization"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.Normalization = table.Rows[0]["Spare_Respiratory_Capacity$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.Normalization ? "Normalization for Coupling Efficiency is true" : "Normalization for Coupling Efficiency is false");
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Error Format"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.ErrorFormat = table.Rows[0]["Spare_Respiratory_Capacity$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for Spare Respiratory Capacity Percentage  is " + _workFlow6.SpareRespiratoryCapacityPercentage.ErrorFormat);
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Sort By"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.SortBy = table.Rows[0]["Spare_Respiratory_Capacity$Sort By"].ToString();
                    _extentTest.Log(Status.Pass, "Sort By for Spare Respiratory Capacity Percentage  is " + _workFlow6.SpareRespiratoryCapacityPercentage.SortBy);
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$Expected GraphUnits"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.ExpectedGraphUnits = table.Rows[0]["Spare_Respiratory_Capacity$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "GraphUnits for Spare Respiratory Capacity Percentage value is " + _workFlow6.SpareRespiratoryCapacityPercentage.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$GraphSettingsRequired"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettingsVerify = table.Rows[0]["Spare_Respiratory_Capacity$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettingsVerify ? "GraphSettingsVerify for Spare Respiratory Capacity Percentage is true" : "GraphSettingsVerify for Spare Respiratory Capacity Percentage is false");
                    if (_workFlow6.SpareRespiratoryCapacityPercentage.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Spare_Respiratory_Capacity$Remove Y AutoScale"))
                        {
                            _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.RemoveYAutoScale = table.Rows[0]["Spare_Respiratory_Capacity$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for Spare Respiratory Capacity Percentage is true" : "Remove Y AutoScale in GraphSettings for Spare Respiratory Capacity Percentage is false");
                        }
                        if (tblcolumns.Contains("Spare_Respiratory_Capacity$Remove ZeroLine"))
                        {
                            _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.RemoveZeroLine = table.Rows[0]["Spare_Respiratory_Capacity$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for Spare Respiratory Capacity Percentage is true" : "Zeroline in GraphSettings for Spare Respiratory Capacity Percentage is false");
                        }

                        if (tblcolumns.Contains("Spare_Respiratory_Capacity$Remove Zoom"))
                        {
                            _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.RemoveZoom = table.Rows[0]["Spare_Respiratory_Capacity$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for Spare Respiratory Capacity Percentage is true" : "Remove Zoom in GraphSettings for Spare Respiratory Capacity Percentage is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$CheckNormalizationWithPlateMap"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.CheckNormalizationWithPlateMap = table.Rows[0]["Spare_Respiratory_Capacity$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$PlateMap Sync to View"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.PlateMapSynctoView = table.Rows[0]["Spare_Respiratory_Capacity$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$GraphSettings Sync to View"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.SynctoView = table.Rows[0]["Spare_Respiratory_Capacity$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Spare_Respiratory_Capacity$IsExportRequired"))
                {
                    _workFlow6.SpareRespiratoryCapacityPercentage.IsExportRequired = table.Rows[0]["Spare_Respiratory_Capacity$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.SpareRespiratoryCapacityPercentage.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }

                #endregion

                #region TestId-15

                _workFlow6.DataTable = new WidgetItems();
                _workFlow6.DataTable.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Data_Table$Default Oligo"))
                {
                    _workFlow6.DataTable.Oligo = table.Rows[0]["Data_Table$Default Oligo"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.DataTable.Oligo))
                    {
                        _extentTest.Log(Status.Fail, "Default oligo for Data Table is missing");
                        message += "Default oligo for Data Table&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Default oligo for Data Table  is " + _workFlow6.DataTable.Oligo);
                    }
                }

                if (tblcolumns.Contains("Data_Table$Default Normalization"))
                {
                    _workFlow6.DataTable.Normalization = table.Rows[0]["Data_Table$Default Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow6.DataTable.Normalization ? "Default normalization for Data Table is true" : "Default normalization for Data Table is false");
                }

                if (tblcolumns.Contains("Data_Table$Default Error Format"))
                {
                    _workFlow6.DataTable.ErrorFormat = table.Rows[0]["Data_Table$Default Error Format"].ToString();
                    if (string.IsNullOrEmpty(_workFlow6.DataTable.ErrorFormat))
                    {
                        _extentTest.Log(Status.Fail, "Default error format for Data Table is missing");
                        message += "Default error format for Data Table&";
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "Default error format for Data Table  is " + _workFlow6.DataTable.ErrorFormat);
                    }
                }

                if (tblcolumns.Contains("Data_Table$DataTableSettingsRequired"))
                {
                    _workFlow6.DataTable.GraphSettingsVerify = table.Rows[0]["Data_Table$DataTableSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow6.DataTable.GraphSettingsVerify ? "Data Table Settings Verify for Data Table is true" : "Data Table Settings Verify for Data Table is false");

                }
                _workFlow6.DataTable.IsExportRequired = table.Rows[0]["Data_Table$IsExportRequired"].ToString() == "Yes" ? true : false;

                #endregion
            }
            else if (sheetName == "Workflow-7")
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
                if (_fileUploadOrExistingFileData.IsTitrationFile)
                    _extentTest.Log(Status.Pass, "File Titration type is " + _fileUploadOrExistingFileData.IsTitrationFile);
                else
                    _extentTest.Log(Status.Pass, "File Titration type is " + _fileUploadOrExistingFileData.IsTitrationFile);

                _fileUploadOrExistingFileData.IsNormalized = table.Rows[0]["Upload_File$IsNormalized"].ToString() == "Yes";
                if (_fileUploadOrExistingFileData.IsNormalized)
                    _extentTest.Log(Status.Pass, "File Normalization status is " + _fileUploadOrExistingFileData.IsNormalized);
                else
                    _extentTest.Log(Status.Pass, "File Normalization status is " + _fileUploadOrExistingFileData.IsNormalized);


                _fileUploadOrExistingFileData.OligoInjection = table.Rows[0]["Upload_File$Oligo Injection"].ToString();
                _extentTest.Log(Status.Pass, "Oligo Injection is selected as - " + _fileUploadOrExistingFileData.OligoInjection + " Injection");

                var filetype = table.Rows[0]["Upload_File$FileType"].ToString();
                if (String.IsNullOrEmpty(filetype))
                {
                    _extentTest.Log(Status.Fail, "File Type is empty");
                    message += "File Type&";
                }
                else
                {
                    _fileUploadOrExistingFileData.FileType = filetype == "XFe24" ? FileType.Xfe24 : filetype == "XFe96" ? FileType.Xfe96 : filetype == "XFp" ? FileType.Xfp : filetype == "HsMini" ? FileType.XfHsMini : filetype == "XFPro" ? FileType.XFPro : FileType.XFPro;
                    _extentTest.Log(Status.Pass, "File Type is " + filetype);
                }


                _fileUploadOrExistingFileData.SelectedWidgets = new List<WidgetTypes>();
                if (table.Rows[0]["Upload_File$mitoATP Production Rate"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.MitoAtpProductionRate);

                if (table.Rows[0]["Upload_File$glycoATP Production Rate "].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.GlycoAtpProductionRate);

                if (table.Rows[0]["Upload_File$ATP Production Rate Data "].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.AtpProductionRateData);

                if (table.Rows[0]["Upload_File$ATP Production Rate (Basal)"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.AtpProductionRateBasal);

                if (table.Rows[0]["Upload_File$ATP production Rate (Induced)"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.AtpProductionRateInduced);

                if (table.Rows[0]["Upload_File$Energetic Map (Basal)"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.EnergeticMapBasal);

                if (table.Rows[0]["Upload_File$Energetic Map (Induced)"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.EnergeticMapInduced);

                if (table.Rows[0]["Upload_File$XF ATP Rate Index"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.XfAtpRateIndex);

                if (table.Rows[0]["Upload_File$Data Table"].ToString() == "Yes")
                    _fileUploadOrExistingFileData.SelectedWidgets.Add(WidgetTypes.DataTable);

                #endregion

                #region TestId -2

                if (tblcolumns.Contains("Layout_Verification$Layout Verification"))
                {
                    _workFlow7.AnalysisLayoutVerification = table.Rows[0]["Layout_Verification$Layout Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.AnalysisLayoutVerification ? "Analysis page layout verification is true" : "Analysis page layout verification is false");
                }
                #endregion

                #region TestId -3

                if (tblcolumns.Contains("Normalization_Icon$Normalization Verification"))
                {
                    _workFlow7.Normalization = table.Rows[0]["Normalization_Icon$Normalization Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.Normalization ? "Normalization Verification is true" : "Normalization Verification is false");
                    ReadDataFromExcel("Normalization");

                    if (_workFlow7.Normalization)
                    {

                        if (tblcolumns.Contains("Normalization_Icon$Apply to all widgets"))
                        {
                            _workFlow7.ApplyToAllWidgets = table.Rows[0]["Normalization_Icon$Apply to all widgets"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, "ScaleFactor for Barchart is " + _workFlow7.ApplyToAllWidgets);
                        }

                        if (tblcolumns.Contains("Normalization_Icon$Normalized File Name"))
                        {
                            _workFlow7.NormalizedFileName = table.Rows[0]["Normalization_Icon$Normalized File Name"].ToString();
                            if (string.IsNullOrEmpty(_workFlow7.NormalizedFileName))
                            {
                                _extentTest.Log(Status.Fail, "FileName status is false");
                                message += "FileName&";
                            }
                            else
                            {
                                _extentTest.Log(Status.Pass, "The Normalizied file name is " + _workFlow7.NormalizedFileName);
                            }
                        }
                    }
                }
                #endregion

                #region TestId -4

                if (tblcolumns.Contains("Modify_Assay$ModifyAssay Verification"))
                {
                    _workFlow7.ModifyAssay = table.Rows[0]["Modify_Assay$ModifyAssay Verification"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ModifyAssay ? "ModifyAssay Verification is true" : "ModifyAssay Verification is false");
                }

                #endregion

                #region TestId -5

                _workFlow7.MitoATPProductionRate = new WidgetItems();
                _workFlow7.MitoATPProductionRate.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("mitoATP_Production_Rate$Oligo"))
                {
                    _workFlow7.MitoATPProductionRate.Oligo = table.Rows[0]["mitoATP_Production_Rate$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "oligo for mitoATP Production Rate  is " + _workFlow7.MitoATPProductionRate.Oligo);
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$Induced"))
                {
                    _workFlow7.MitoATPProductionRate.Induced = table.Rows[0]["mitoATP_Production_Rate$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "induced for mitoATP Production Rate  is " + _workFlow7.MitoATPProductionRate.Induced);
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$Normalization"))
                {
                    _workFlow7.MitoATPProductionRate.Normalization = table.Rows[0]["mitoATP_Production_Rate$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.Normalization ? "normalization for mitoATP Production Rate is true" : "normalizaion for mitoATP Production Rate is false");
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$Error Format"))
                {
                    _workFlow7.MitoATPProductionRate.ErrorFormat = table.Rows[0]["mitoATP_Production_Rate$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "error format for mitoATP Production Rate  is " + _workFlow7.MitoATPProductionRate.ErrorFormat);
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$Expected GraphUnits"))
                {
                    _workFlow7.MitoATPProductionRate.ExpectedGraphUnits = table.Rows[0]["mitoATP_Production_Rate$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits for mitoATP Production Rate value is " + _workFlow7.MitoATPProductionRate.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$GraphSettingsRequired"))
                {
                    _workFlow7.MitoATPProductionRate.GraphSettingsVerify = table.Rows[0]["mitoATP_Production_Rate$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.GraphSettingsVerify ? "GraphSettingsVerify for mito ATP Production Rate is true" : "GraphSettingsVerify for mito ATP Production Rate is false");
                    if (_workFlow7.MitoATPProductionRate.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("mitoATP_Production_Rate$Remove Y AutoScale"))
                        {
                            _workFlow7.MitoATPProductionRate.GraphSettings.RemoveYAutoScale = table.Rows[0]["mitoATP_Production_Rate$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for mito ATP Production Rate is true" : "Remove Y AutoScale in GraphSettings for mito ATP Production Rate is false");
                        }

                        if (tblcolumns.Contains("mitoATP_Production_Rate$Remove ZeroLine"))
                        {
                            _workFlow7.MitoATPProductionRate.GraphSettings.RemoveZeroLine = table.Rows[0]["mitoATP_Production_Rate$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for mito ATP Production Rate is true" : "Zeroline in GraphSettings for mitoATP Production Rate is false");
                        }

                        if (tblcolumns.Contains("mitoATP_Production_Rate$Remove Zoom"))
                        {
                            _workFlow7.MitoATPProductionRate.GraphSettings.RemoveZoom = table.Rows[0]["mitoATP_Production_Rate$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for mito ATP Production Rate is true" : "Remove Zoom in GraphSettings for mito ATP Production Rate is false");
                        }
                    }
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.MitoATPProductionRate.CheckNormalizationWithPlateMap = table.Rows[0]["mitoATP_Production_Rate$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$PlateMap Sync to View"))
                {
                    _workFlow7.MitoATPProductionRate.PlateMapSynctoView = table.Rows[0]["mitoATP_Production_Rate$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$GraphSettings Sync to View"))
                {
                    _workFlow7.MitoATPProductionRate.GraphSettings.SynctoView = table.Rows[0]["mitoATP_Production_Rate$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("mitoATP_Production_Rate$IsExportRequired"))
                {
                    _workFlow7.MitoATPProductionRate.IsExportRequired = table.Rows[0]["mitoATP_Production_Rate$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.MitoATPProductionRate.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId - 6

                _workFlow7.GlycoATPProductionRate = new WidgetItems();
                _workFlow7.GlycoATPProductionRate.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("glycoATP_Production _Rate$Oligo"))
                {
                    _workFlow7.GlycoATPProductionRate.Oligo = table.Rows[0]["glycoATP_Production _Rate$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "oligo for glycoATP Production Rate  is " + _workFlow7.GlycoATPProductionRate.Oligo);
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$Induced"))
                {
                    _workFlow7.GlycoATPProductionRate.Induced = table.Rows[0]["glycoATP_Production _Rate$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "Induced for glycoATP Production Rate  is " + _workFlow7.GlycoATPProductionRate.Induced);
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$Normalization"))
                {
                    _workFlow7.GlycoATPProductionRate.Normalization = table.Rows[0]["glycoATP_Production _Rate$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.Normalization ? "normalizationfor ATP Production Rate Data is true" : "normalizaionfor ATP Production Rate Data is false");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$Error Format"))
                {
                    _workFlow7.GlycoATPProductionRate.ErrorFormat = table.Rows[0]["glycoATP_Production _Rate$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "error formatfor ATP Production Rate Data  is " + _workFlow7.GlycoATPProductionRate.ErrorFormat);
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$Expected GraphUnits"))
                {
                    _workFlow7.GlycoATPProductionRate.ExpectedGraphUnits = table.Rows[0]["glycoATP_Production _Rate$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits for glycoATP Production Rate value is " + _workFlow7.GlycoATPProductionRate.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$GraphSettingsRequired"))
                {
                    _workFlow7.GlycoATPProductionRate.GraphSettingsVerify = table.Rows[0]["glycoATP_Production _Rate$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.GraphSettingsVerify ? "GraphSettingsVerify for glyco ATP Production Rate is true" : "GraphSettingsVerify for glyco ATP Production Rate is false");
                    if (_workFlow7.GlycoATPProductionRate.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("glycoATP_Production _Rate$Remove Y AutoScale"))
                        {
                            _workFlow7.GlycoATPProductionRate.GraphSettings.RemoveYAutoScale = table.Rows[0]["glycoATP_Production _Rate$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for glyco ATP Production Rate is true" : "Remove Y AutoScale in GraphSettings for glyco ATP Production Rate is false");
                        }
                        if (tblcolumns.Contains("glycoATP_Production _Rate$Remove ZeroLine"))
                        {
                            _workFlow7.GlycoATPProductionRate.GraphSettings.RemoveZeroLine = table.Rows[0]["glycoATP_Production _Rate$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for glyco ATP Production Rate is true" : "Zeroline in GraphSettings for glyco ATP Production Rate is false");
                        }

                        if (tblcolumns.Contains("glycoATP_Production _Rate$Remove Zoom"))
                        {
                            _workFlow7.GlycoATPProductionRate.GraphSettings.RemoveZoom = table.Rows[0]["glycoATP_Production _Rate$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for glyco ATP Production Rate is true" : "Remove Zoom in GraphSettings for glyco ATP Production Rate is false");
                        }
                    }
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.GlycoATPProductionRate.CheckNormalizationWithPlateMap = table.Rows[0]["glycoATP_Production _Rate$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$PlateMap Sync to View"))
                {
                    _workFlow7.GlycoATPProductionRate.PlateMapSynctoView = table.Rows[0]["glycoATP_Production _Rate$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$GraphSettings Sync to View"))
                {
                    _workFlow7.GlycoATPProductionRate.GraphSettings.SynctoView = table.Rows[0]["glycoATP_Production _Rate$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$IsExportRequired"))
                {
                    _workFlow7.GlycoATPProductionRate.IsExportRequired = table.Rows[0]["glycoATP_Production _Rate$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.GlycoATPProductionRate.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-7

                _workFlow7.ATPProductionRateData = new WidgetItems();
                _workFlow7.ATPProductionRateData.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("ATP_Production_Rate_Data$Oligo"))
                {
                    _workFlow7.ATPProductionRateData.Oligo = table.Rows[0]["ATP_Production_Rate_Data$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "oligo forATP Production Rate Data  is " + _workFlow7.ATPProductionRateData.Oligo);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Data$Induced"))
                {
                    _workFlow7.ATPProductionRateData.Induced = table.Rows[0]["ATP_Production_Rate_Data$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "induced for ATP Production Rate Data  is " + _workFlow7.ATPProductionRateData.Induced);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Data$Normalization"))
                {
                    _workFlow7.ATPProductionRateData.Normalization = table.Rows[0]["ATP_Production_Rate_Data$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.Normalization ? "normalizationfor ATP Production Rate Data is true" : "normalizaionfor ATP Production Rate Data is false");
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Data$Error Format"))
                {
                    _workFlow7.ATPProductionRateData.ErrorFormat = table.Rows[0]["ATP_Production_Rate_Data$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "error format for ATP Production Rate Data  is " + _workFlow7.ATPProductionRateData.ErrorFormat);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Data$Expected GraphUnits"))
                {
                    _workFlow7.ATPProductionRateData.ExpectedGraphUnits = table.Rows[0]["ATP_Production_Rate_Data$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits forATP Production Rate Data value is " + _workFlow7.ATPProductionRateData.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Data$GraphSettingsRequired"))
                {
                    _workFlow7.ATPProductionRateData.GraphSettingsVerify = table.Rows[0]["ATP_Production_Rate_Data$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.GraphSettingsVerify ? "GraphSettingsVerify for ATP Production Rate Data is true" : "GraphSettingsVerify for ATP Production Rate Data is false");
                    if (_workFlow7.ATPProductionRateData.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("ATP_Production_Rate_Data$Remove Y AutoScale"))
                        {
                            _workFlow7.ATPProductionRateData.GraphSettings.RemoveYAutoScale = table.Rows[0]["ATP_Production_Rate_Data$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for ATP Production Rate Data is true" : "Remove Y AutoScale in GraphSettings for ATP Production Rate Data is false");
                        }

                        if (tblcolumns.Contains("ATP_Production_Rate_Data$Remove ZeroLine"))
                        {
                            _workFlow7.ATPProductionRateData.GraphSettings.RemoveZeroLine = table.Rows[0]["ATP_Production_Rate_Data$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings for ATP Production Rate Data is true" : "Zeroline in GraphSettings for ATP Production Rate Data is false");
                        }

                        if (tblcolumns.Contains("ATP_Production_Rate_Data$Remove Zoom"))
                        {
                            _workFlow7.ATPProductionRateData.GraphSettings.RemoveZoom = table.Rows[0]["ATP_Production_Rate_Data$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for  ATP Production Rate Data is true" : "Remove Zoom in GraphSettings for ATP Production Rate Data is false");
                        }
                    }
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Data$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.ATPProductionRateData.CheckNormalizationWithPlateMap = table.Rows[0]["ATP_Production_Rate_Data$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$PlateMap Sync to View"))
                {
                    _workFlow7.ATPProductionRateData.PlateMapSynctoView = table.Rows[0]["ATP_Production_Rate_Data$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$GraphSettings Sync to View"))
                {
                    _workFlow7.ATPProductionRateData.GraphSettings.SynctoView = table.Rows[0]["ATP_Production_Rate_Data$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("glycoATP_Production _Rate$IsExportRequired"))
                {
                    _workFlow7.ATPProductionRateData.IsExportRequired = table.Rows[0]["ATP_Production_Rate_Data$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRateData.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-8

                _workFlow7.ATPProductionRate_Basal = new WidgetItems();
                _workFlow7.ATPProductionRate_Basal.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Display"))
                {
                    _workFlow7.ATPProductionRate_Basal.Display = table.Rows[0]["ATP_Production_Rate_Basal$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for ATPProduction Rate (Basal)  is " + _workFlow7.ATPProductionRate_Basal.Display);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Oligo"))
                {
                    _workFlow7.ATPProductionRate_Basal.Oligo = table.Rows[0]["ATP_Production_Rate_Basal$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for ATPProduction Rate(Basal)  is " + _workFlow7.ATPProductionRate_Basal.Oligo);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Induced"))
                {
                    _workFlow7.ATPProductionRate_Basal.Induced = table.Rows[0]["ATP_Production_Rate_Basal$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "Induced for ATPProduction Rate (Basal)  is " + _workFlow7.ATPProductionRate_Basal.Induced);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Display"))
                {
                    _workFlow7.ATPProductionRate_Basal.Display = table.Rows[0]["ATP_Production_Rate_Basal$Display"].ToString();
                    _extentTest.Log(Status.Pass, "Displaymode for ATPProduction Rate (Basal)  is " + _workFlow7.ATPProductionRate_Basal.Display);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Normalization"))
                {
                    _workFlow7.ATPProductionRate_Basal.Normalization = table.Rows[0]["ATP_Production_Rate_Basal$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.Normalization ? "Normalizationfor ATP Production Rate (Basal) is true" : "Normalizaionfor ATP Production Rate (Basal) is false");
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Error Format"))
                {
                    _workFlow7.ATPProductionRate_Basal.ErrorFormat = table.Rows[0]["ATP_Production_Rate_Basal$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for  ATPProduction Rate (Basal) is " + _workFlow7.ATPProductionRate_Basal.ErrorFormat);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$Expected GraphUnits"))
                {
                    _workFlow7.ATPProductionRate_Basal.ExpectedGraphUnits = table.Rows[0]["ATP_Production_Rate_Basal$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits forATPProduction Rate (Basal) value is " + _workFlow7.ATPProductionRate_Basal.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$GraphSettingsRequired"))
                {
                    _workFlow7.ATPProductionRate_Basal.GraphSettingsVerify = table.Rows[0]["ATP_Production_Rate_Basal$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.GraphSettingsVerify ? "GraphSettings Verify for ATP Production Rate (Basal) is true" : "GraphSettings Verify for ATP Production Rate (Basal) is false");
                    if (_workFlow7.ATPProductionRate_Basal.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("ATP_Production_Rate_Basal$Remove Y AutoScale"))
                        {
                            _workFlow7.ATPProductionRate_Basal.GraphSettings.RemoveYAutoScale = table.Rows[0]["ATP_Production_Rate_Basal$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for ATP Production Rate (Basal) is true" : "Remove Y AutoScale in GraphSettings for ATP Production Rate (Basal) is false");
                        }

                        if (tblcolumns.Contains("ATP_Production_Rate_Basal$Remove ZeroLine"))
                        {
                            _workFlow7.ATPProductionRate_Basal.GraphSettings.RemoveZeroLine = table.Rows[0]["ATP_Production_Rate_Basal$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.GraphSettings.RemoveZeroLine ? "Reomve Zeroline in GraphSettings for ATP Production Rate (Basal) is true" : "Reomve Zeroline in GraphSettings for ATPProduction Rate (Basal) is false");
                        }

                        if (tblcolumns.Contains("ATP_Production_Rate_Basal$Remove Zoom"))
                        {
                            _workFlow7.ATPProductionRate_Basal.GraphSettings.RemoveZoom = table.Rows[0]["ATP_Production_Rate_Basal$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for ATP Production Rate (Basal) is true" : "Remove Zoom in GraphSettings for ATP Production Rate (Basal) is false");
                        }
                    }
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.ATPProductionRate_Basal.CheckNormalizationWithPlateMap = table.Rows[0]["ATP_Production_Rate_Basal$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$PlateMap Sync to View"))
                {
                    _workFlow7.ATPProductionRate_Basal.PlateMapSynctoView = table.Rows[0]["ATP_Production_Rate_Basal$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$GraphSettings Sync to View"))
                {
                    _workFlow7.ATPProductionRate_Basal.GraphSettings.SynctoView = table.Rows[0]["ATP_Production_Rate_Basal$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_Production_Rate_Basal$IsExportRequired"))
                {
                    _workFlow7.ATPProductionRate_Basal.IsExportRequired = table.Rows[0]["ATP_Production_Rate_Basal$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPProductionRate_Basal.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-9

                _workFlow7.ATPproductionRate_Induced = new WidgetItems();
                _workFlow7.ATPproductionRate_Induced.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("ATP_production_Rate_Induced$Display"))
                {
                    _workFlow7.ATPproductionRate_Induced.Display = table.Rows[0]["ATP_production_Rate_Induced$Display"].ToString();
                    _extentTest.Log(Status.Pass, "displaymode for ATP Production Rate (Induced)  is " + _workFlow7.ATPproductionRate_Induced.Display);
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$Oligo"))
                {
                    _workFlow7.ATPproductionRate_Induced.Oligo = table.Rows[0]["ATP_production_Rate_Induced$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for ATP Production Rate (Induced)  is " + _workFlow7.ATPproductionRate_Induced.Oligo);
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$Induced"))
                {
                    _workFlow7.ATPproductionRate_Induced.Induced = table.Rows[0]["ATP_production_Rate_Induced$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "Induced for ATP Production Rate (Induced)  is " + _workFlow7.ATPproductionRate_Induced.Induced);
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$Normalization"))
                {
                    _workFlow7.ATPproductionRate_Induced.Normalization = table.Rows[0]["ATP_production_Rate_Induced$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.Normalization ? "Normalizationfor ATP Production Rate (Induced) is true" : "Normalizaionfor ATP Production Rate (Induced) is false");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$Error Format"))
                {
                    _workFlow7.ATPproductionRate_Induced.ErrorFormat = table.Rows[0]["ATP_production_Rate_Induced$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "error format for  ATP Production Rate (Induced )is " + _workFlow7.ATPproductionRate_Induced.ErrorFormat);
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$Expected GraphUnits"))
                {
                    _workFlow7.ATPproductionRate_Induced.ExpectedGraphUnits = table.Rows[0]["ATP_production_Rate_Induced$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits for ATP Production Rate (Induced) value is " + _workFlow7.ATPproductionRate_Induced.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$GraphSettingsRequired"))
                {
                    _workFlow7.ATPproductionRate_Induced.GraphSettingsVerify = table.Rows[0]["ATP_production_Rate_Induced$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.GraphSettingsVerify ? "GraphSettings Verify for ATP Production Rate (Induced) is true" : "GraphSettings Verify for ATP Production Rate (Induced) is false");
                    if (_workFlow7.ATPproductionRate_Induced.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("ATP_production_Rate_Induced$Remove Y AutoScale"))
                        {
                            _workFlow7.ATPproductionRate_Induced.GraphSettings.RemoveYAutoScale = table.Rows[0]["ATP_production_Rate_Induced$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for ATP Production Rate (Induced) is true" : "Remove Y AutoScale in GraphSettings for ATP Production Rate (Induced) is false");
                        }
                        if (tblcolumns.Contains("ATP_production_Rate_Induced$Remove ZeroLine"))
                        {
                            _workFlow7.ATPproductionRate_Induced.GraphSettings.RemoveZeroLine = table.Rows[0]["ATP_production_Rate_Induced$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.GraphSettings.RemoveZeroLine ? "Remove Zeroline in GraphSettings for ATPProduction Rate (Induced) is true" : "Remove Zeroline in GraphSettings for ATPProduction Rate (Induced) is false");
                        }

                        if (tblcolumns.Contains("ATP_production_Rate_Induced$Remove Zoom"))
                        {
                            _workFlow7.ATPproductionRate_Induced.GraphSettings.RemoveZoom = table.Rows[0]["ATP_production_Rate_Induced$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for  ATP Production Rate (Induced) is true" : "Remove Zoom in GraphSettings for ATP Production Rate (Induced) is false");
                        }
                    }
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.ATPproductionRate_Induced.CheckNormalizationWithPlateMap = table.Rows[0]["ATP_production_Rate_Induced$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$PlateMap Sync to View"))
                {
                    _workFlow7.ATPproductionRate_Induced.PlateMapSynctoView = table.Rows[0]["ATP_production_Rate_Induced$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$GraphSettings Sync to View"))
                {
                    _workFlow7.ATPproductionRate_Induced.GraphSettings.SynctoView = table.Rows[0]["ATP_production_Rate_Induced$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$IsExportRequired"))
                {
                    _workFlow7.ATPproductionRate_Induced.IsExportRequired = table.Rows[0]["ATP_production_Rate_Induced$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.ATPproductionRate_Induced.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-10

                _workFlow7.EnergeticMap_Basal = new WidgetItems();
                _workFlow7.EnergeticMap_Basal.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Energetic_Map_Basal$Oligo"))
                {
                    _workFlow7.EnergeticMap_Basal.Oligo = table.Rows[0]["Energetic_Map_Basal$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo EnergeticMap Rate (Basal)  is " + _workFlow7.EnergeticMap_Basal.Oligo);
                }

                if (tblcolumns.Contains("Energetic_Map_Basal$Induced"))
                {
                    _workFlow7.EnergeticMap_Basal.Induced = table.Rows[0]["Energetic_Map_Basal$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "Induced EnergeticMap Rate (Basal) is " + _workFlow7.EnergeticMap_Basal.Induced);
                }

                if (tblcolumns.Contains("Energetic_Map_Basal$Normalization"))
                {
                    _workFlow7.EnergeticMap_Basal.Normalization = table.Rows[0]["Energetic_Map_Basal$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.Normalization ? "Normalization for EnergeticMap Rate (Basal) is true" : "Normalizaion for EnergeticMap Rate (Basal) is false");
                }

                if (tblcolumns.Contains("Energetic_Map_Basal$Error Format"))
                {
                    _workFlow7.EnergeticMap_Basal.ErrorFormat = table.Rows[0]["Energetic_Map_Basal$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error formatfor EnergeticMap Rate (Basal)  is " + _workFlow7.EnergeticMap_Basal.ErrorFormat);
                }

                if (tblcolumns.Contains("Energetic_Map_Basal$Expected GraphUnits"))
                {
                    _workFlow7.EnergeticMap_Basal.ExpectedGraphUnits = table.Rows[0]["Energetic_Map_Basal$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits EnergeticMap Rate (Basal) value is " + _workFlow7.EnergeticMap_Basal.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Energetic_Map_Basal$GraphSettingsRequired"))
                {
                    _workFlow7.EnergeticMap_Basal.GraphSettingsVerify = table.Rows[0]["Energetic_Map_Basal$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.GraphSettingsVerify ? "GraphSettingsVerify for EnergeticMap Rate (Basal) is true" : "GraphSettingsVerify for EnergeticMap Rate (Basal) is false");
                    if (_workFlow7.EnergeticMap_Basal.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Energetic_Map_Basal$Remove Y AutoScale"))
                        {
                            _workFlow7.EnergeticMap_Basal.GraphSettings.RemoveYAutoScale = table.Rows[0]["Energetic_Map_Basal$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for EnergeticMap Rate (Basal) is true" : "Remove Y AutoScale in GraphSettings forEnergeticMap Rate (Basal) is false");
                        }

                        if (tblcolumns.Contains("Energetic_Map_Basal$Remove ZeroLine"))
                        {
                            _workFlow7.EnergeticMap_Basal.GraphSettings.RemoveZeroLine = table.Rows[0]["Energetic_Map_Basal$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.GraphSettings.RemoveZeroLine ? "Zeroline in GraphSettings EnergeticMap (Basal) Rate is true" : "Zeroline in GraphSettings EnergeticMap (Basal) Rate is false");
                        }

                        if (tblcolumns.Contains("Energetic_Map_Basal$Remove Zoom"))
                        {
                            _workFlow7.EnergeticMap_Basal.GraphSettings.RemoveZoom = table.Rows[0]["Energetic_Map_Basal$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for EnergeticMap Rate (Basal) is true" : "Remove Zoom in GraphSettings for EnergeticMap Rate (Basal) is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Energetic_Map_Basal$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.EnergeticMap_Basal.CheckNormalizationWithPlateMap = table.Rows[0]["Energetic_Map_Basal$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$PlateMap Sync to View"))
                {
                    _workFlow7.EnergeticMap_Basal.PlateMapSynctoView = table.Rows[0]["ATP_production_Rate_Induced$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$GraphSettings Sync to View"))
                {
                    _workFlow7.EnergeticMap_Basal.GraphSettings.SynctoView = table.Rows[0]["ATP_production_Rate_Induced$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("ATP_production_Rate_Induced$IsExportRequired"))
                {
                    _workFlow7.EnergeticMap_Basal.IsExportRequired = table.Rows[0]["Energetic_Map_Basal$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Basal.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-11

                _workFlow7.EnergeticMap_Induced = new WidgetItems();
                _workFlow7.EnergeticMap_Induced.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Energetic_Map_Induced$Oligo"))
                {
                    _workFlow7.EnergeticMap_Induced.Oligo = table.Rows[0]["Energetic_Map_Induced$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "Oligo for EnergeticMap Rate (Induced) is " + _workFlow7.EnergeticMap_Induced.Oligo);
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$Induced"))
                {
                    _workFlow7.EnergeticMap_Induced.Induced = table.Rows[0]["Energetic_Map_Induced$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "Induced for EnergeticMap Rate (Induced) is " + _workFlow7.EnergeticMap_Induced.Induced);
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$Normalization"))
                {
                    _workFlow7.EnergeticMap_Induced.Normalization = table.Rows[0]["Energetic_Map_Induced$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.Normalization ? "Normalization for EnergeticMap Rate (Induced) is true" : "Normalizaionfor EnergeticMap Rate (Induced) is false");
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$Error Format"))
                {
                    _workFlow7.EnergeticMap_Induced.ErrorFormat = table.Rows[0]["Energetic_Map_Induced$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for EnergeticMap Rate (Induced)   is " + _workFlow7.EnergeticMap_Induced.ErrorFormat);
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$Expected GraphUnits"))
                {
                    _workFlow7.EnergeticMap_Induced.ExpectedGraphUnits = table.Rows[0]["Energetic_Map_Induced$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits for EnergeticMap Rate (Induced)value is " + _workFlow7.EnergeticMap_Induced.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$GraphSettingsRequired"))
                {
                    _workFlow7.EnergeticMap_Induced.GraphSettingsVerify = table.Rows[0]["Energetic_Map_Induced$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettingsVerify ? "GraphSettingsVerify for EnergeticMap Rate (Induced) is true" : "GraphSettingsVerify for EnergeticMap Rate (Induced) is false");
                    if (_workFlow7.EnergeticMap_Induced.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("Energetic_Map_Induced$Remove Y AutoScale"))
                        {
                            _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveYAutoScale = table.Rows[0]["Energetic_Map_Induced$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for EnergeticMap Rate (Induced) is true" : "Remove Y AutoScale in GraphSettings for EnergeticMap Rate (Induced) is false");
                        }

                        if (tblcolumns.Contains("Energetic_Map_Induced$Remove ZeroLine"))
                        {
                            _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveZeroLine = table.Rows[0]["Energetic_Map_Induced$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveZeroLine ? "Remove Zeroline in GraphSettings for EnergeticMap Rate (Induced) is true" : "Remove Zeroline in GraphSettings for EnergeticMap Rate (Induced) is false");
                        }

                        if (tblcolumns.Contains("Energetic_Map_Induced$Remove Zoom"))
                        {
                            _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveZoom = table.Rows[0]["Energetic_Map_Induced$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for EnergeticMap Rate (Induced) is true" : "Remove Zoom in GraphSettings for EnergeticMap Rate (Induced) is false");
                        }
                    }
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.EnergeticMap_Induced.CheckNormalizationWithPlateMap = table.Rows[0]["Energetic_Map_Induced$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$PlateMap Sync to View"))
                {
                    _workFlow7.EnergeticMap_Induced.PlateMapSynctoView = table.Rows[0]["Energetic_Map_Induced$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$GraphSettings Sync to View"))
                {
                    _workFlow7.EnergeticMap_Induced.GraphSettings.SynctoView = table.Rows[0]["Energetic_Map_Induced$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("Energetic_Map_Induced$IsExportRequired"))
                {
                    _workFlow7.EnergeticMap_Induced.IsExportRequired = table.Rows[0]["Energetic_Map_Induced$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-12

                _workFlow7.XFATPRateIndex = new WidgetItems();
                _workFlow7.XFATPRateIndex.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("XF_ATP_Rate_Index$Oligo"))
                {
                    _workFlow7.XFATPRateIndex.Oligo = table.Rows[0]["XF_ATP_Rate_Index$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "oligo for XF ATP Rate Index is " + _workFlow7.XFATPRateIndex.Oligo);
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$Induced"))
                {
                    _workFlow7.XFATPRateIndex.Induced = table.Rows[0]["XF_ATP_Rate_Index$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "induced for XF ATP Rate Index is " + _workFlow7.XFATPRateIndex.Induced);
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$Normalization"))
                {
                    _workFlow7.XFATPRateIndex.Normalization = table.Rows[0]["XF_ATP_Rate_Index$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.Normalization ? "normalizationfor ATP Production Rate Data is true" : "normalizaionfor ATP Production Rate Data is false");
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$Error Format"))
                {
                    _workFlow7.XFATPRateIndex.ErrorFormat = table.Rows[0]["XF_ATP_Rate_Index$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "error format for XF ATP Rate Index   is " + _workFlow7.XFATPRateIndex.ErrorFormat);
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$Expected GraphUnits"))
                {
                    _workFlow7.XFATPRateIndex.ExpectedGraphUnits = table.Rows[0]["XF_ATP_Rate_Index$Expected GraphUnits"].ToString();
                    _extentTest.Log(Status.Pass, "Normalized GraphUnits for XF ATP Rate Indexvalue is " + _workFlow7.XFATPRateIndex.ExpectedGraphUnits);
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$GraphSettingsRequired"))
                {
                    _workFlow7.XFATPRateIndex.GraphSettingsVerify = table.Rows[0]["XF_ATP_Rate_Index$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.GraphSettingsVerify ? "GraphSettingsVerify for XF ATP Rate Index is true" : "GraphSettingsVerify for XF ATP Rate Index is false");
                    if (_workFlow7.XFATPRateIndex.GraphSettingsVerify)
                    {
                        if (tblcolumns.Contains("XF_ATP_Rate_Index$Remove Y AutoScale"))
                        {
                            _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveYAutoScale = table.Rows[0]["XF_ATP_Rate_Index$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for XF ATP Rate Index is true" : "Remove Y AutoScale in GraphSettings for XF ATP Rate Index is false");
                        }
                        if (tblcolumns.Contains("XF_ATP_Rate_Index$Remove ZeroLine"))
                        {
                            _workFlow7.XFATPRateIndex.GraphSettings.RemoveZeroLine = table.Rows[0]["XF_ATP_Rate_Index$Remove ZeroLine"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.GraphSettings.RemoveZeroLine ? "Remove Zeroline in GraphSettings for  XF ATP Rate Index is true" : "Remove Zeroline in GraphSettings for XF ATP Rate Index is false");
                        }

                        if (tblcolumns.Contains("XF_ATP_Rate_Index$Remove Zoom"))
                        {
                            _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveZoom = table.Rows[0]["XF_ATP_Rate_Index$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow7.EnergeticMap_Induced.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for XF ATP Rate Index is true" : "Remove Zoom in GraphSettings for XF ATP Rate Index is false");
                        }
                    }
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$CheckNormalizationWithPlateMap"))
                {
                    _workFlow7.XFATPRateIndex.CheckNormalizationWithPlateMap = table.Rows[0]["XF_ATP_Rate_Index$CheckNormalizationWithPlateMap"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.CheckNormalizationWithPlateMap ? "Normalization needs to be verified with platemap" : "Normalization need not be verified with platemap");
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$PlateMap Sync to View"))
                {
                    _workFlow7.XFATPRateIndex.PlateMapSynctoView = table.Rows[0]["XF_ATP_Rate_Index$PlateMap Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.PlateMapSynctoView ? "PlateMap Sync to View needs to be verified with platemap" : "PlateMap Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$GraphSettings Sync to View"))
                {
                    _workFlow7.XFATPRateIndex.GraphSettings.SynctoView = table.Rows[0]["XF_ATP_Rate_Index$GraphSettings Sync to View"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.GraphSettings.SynctoView ? "GraphSettings Sync to View needs to be verified with platemap" : "GraphSettings Sync to View need not be verified with platemap");
                }

                if (tblcolumns.Contains("XF_ATP_Rate_Index$IsExportRequired"))
                {
                    _workFlow7.XFATPRateIndex.IsExportRequired = table.Rows[0]["XF_ATP_Rate_Index$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.XFATPRateIndex.IsExportRequired ? "Exports is required." : "Exports is not required.");
                }
                #endregion

                #region TestId-13

                _workFlow7.DataTable = new WidgetItems();
                _workFlow7.DataTable.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Data_Table$Oligo"))
                {
                    _workFlow7.DataTable.Oligo = table.Rows[0]["Data_Table$Oligo"].ToString();
                    _extentTest.Log(Status.Pass, "oligo for Data Table Widget  is " + _workFlow7.DataTable.Oligo);
                }

                if (tblcolumns.Contains("Data_Table$Induced"))
                {
                    _workFlow7.DataTable.Induced = table.Rows[0]["Data_Table$Induced"].ToString();
                    _extentTest.Log(Status.Pass, "induced for Data Table Widget  is " + _workFlow7.DataTable.Induced);
                }

                if (tblcolumns.Contains("Data_Table$Normalization"))
                {
                    _workFlow7.DataTable.Normalization = table.Rows[0]["Data_Table$Normalization"].ToString() == "ON";
                    _extentTest.Log(Status.Pass, _workFlow7.DataTable.Normalization ? "normalizationfor ATP Production Rate Data is true" : "normalizaionfor ATP Production Rate Data is false");
                }

                if (tblcolumns.Contains("Data_Table$Error Format"))
                {
                    _workFlow7.DataTable.ErrorFormat = table.Rows[0]["Data_Table$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "error format for Data Table Widget    is " + _workFlow7.DataTable.ErrorFormat);
                }

                if (tblcolumns.Contains("Data_Table$GraphSettingsRequired"))
                {
                    _workFlow7.DataTable.GraphSettingsVerify = table.Rows[0]["Data_Table$GraphSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow7.DataTable.GraphSettingsVerify ? "GraphSettingsVerify for Data Table Widget  is true" : "GraphSettingsVerify for Data Table Widget  is false");
                }

                if (tblcolumns.Contains("Data_Table$IsExportRequired"))
                {
                    _workFlow7.DataTable.IsExportRequired = table.Rows[0]["Data_Table$IsExportRequired"].ToString() == "Yes";
                    if (_workFlow7.XFATPRateIndex.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow7.DataTable.IsExportRequired);
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "File Normalization status is " + _workFlow7.DataTable.IsExportRequired);
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

                if (tblcolumns.Contains("CellEnergy_Phenotype$GraphSettingsRequired"))
                {
                    _workFlow8.CellEnergyPhenotype.GraphSettingsVerify = table.Rows[0]["CellEnergy_Phenotype$GraphSettingsRequired"].ToString() == "Yes";
                    if (_workFlow8.CellEnergyPhenotype.IsExportRequired)
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is required.");

                        if (tblcolumns.Contains("CellEnergy_Phenotype$Remove X AutoScale"))
                        {
                            _workFlow8.CellEnergyPhenotype.GraphSettings.RemoveXAutoScale = table.Rows[0]["CellEnergy_Phenotype$Remove X AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.CellEnergyPhenotype.GraphSettings.RemoveXAutoScale ? "Remove X AutoScale in GraphSettings for CellEnergyPhenotype graph is true" : "Remove X AutoScale in GraphSettings for CellEnergyPhenotype graph is false");
                        }

                        if (tblcolumns.Contains("CellEnergy_Phenotype$Remove Y AutoScale"))
                        {
                            _workFlow8.CellEnergyPhenotype.GraphSettings.RemoveYAutoScale = table.Rows[0]["CellEnergy_Phenotype$Remove Y AutoScale"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.CellEnergyPhenotype.GraphSettings.RemoveYAutoScale ? "Remove Y AutoScale in GraphSettings for CellEnergyPhenotype graph is true" : "Remove Y AutoScale in GraphSettings for CellEnergyPhenotype graph is false");
                        }

                        if (tblcolumns.Contains("CellEnergy_Phenotype$Remove Zoom"))
                        {
                            _workFlow8.CellEnergyPhenotype.GraphSettings.RemoveZoom = table.Rows[0]["CellEnergy_Phenotype$Remove Zoom"].ToString() == "Yes";
                            _extentTest.Log(Status.Pass, _workFlow8.CellEnergyPhenotype.GraphSettings.RemoveZoom ? "Remove Zoom in GraphSettings for CellEnergyPhenotype graph is true" : "Remove Zoom in GraphSettings for CellEnergyPhenotype graph is false");
                        }
                    }
                    else
                    {
                        _extentTest.Log(Status.Pass, "GraphSettings is not required.");
                    }
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$Expected GraphUnits"))
                {
                    var GraphUnits = table.Rows[0]["CellEnergy_Phenotype$Expected GraphUnits"].ToString();
                    _workFlow8.CellEnergyPhenotype.ExpectedGraphUnits = GraphUnits;
                    _extentTest.Log(Status.Pass, "Expected GraphUnits for CellEnergyPhenotype value is " + GraphUnits);
                }

                if (tblcolumns.Contains("CellEnergy_Phenotype$IsExportRequired"))
                {
                    _workFlow8.CellEnergyPhenotype.IsExportRequired = table.Rows[0]["CellEnergy_Phenotype$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.CellEnergyPhenotype.IsExportRequired ? "Exports is required." : "Exports is not required.");
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
                    _workFlow8.MetabolicPotentialOCR.ExpectedGraphUnits = GraphUnits;
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
                    _workFlow8.MetabolicPotentialECAR.ExpectedGraphUnits = GraphUnits;
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
                    _workFlow8.MetabolicPotentialECAR.IsExportRequired = table.Rows[0]["MetabolicPotential_ECAR$IsExportRequired"].ToString() == "Yes";
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
                    _workFlow8.BaselineOCR.ExpectedGraphUnits = GraphUnits;
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
                    _workFlow8.BaselineECAR.ExpectedGraphUnits = GraphUnits;
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
                    _workFlow8.StressedOCR.ExpectedGraphUnits = GraphUnits;
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
                    _workFlow8.StressedECAR.ExpectedGraphUnits = GraphUnits;
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

                #region TestId -10

                _workFlow8.DataTable = new WidgetItems();
                _workFlow8.DataTable.GraphSettings = new GraphSettings();

                if (tblcolumns.Contains("Data_Table$Normalization"))
                {
                    _workFlow8.DataTable.Normalization = table.Rows[0]["Data_Table$Normalization"].ToString().ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.DataTable.Normalization ? "Normalization for DataTable is true" : "Normalization for DataTable is false");
                }

                if (tblcolumns.Contains("Data_Table$Error Format"))
                {
                    _workFlow8.DataTable.ErrorFormat = table.Rows[0]["Data_Table$Error Format"].ToString();
                    _extentTest.Log(Status.Pass, "Error format for DataTable  is " + _workFlow8.StressedECAR.ErrorFormat);
                }

                if (tblcolumns.Contains("Data_Table$DataTableSettingsRequired"))
                {
                    _workFlow8.DataTable.DataTableSettingsVerify = table.Rows[0]["Data_Table$DataTableSettingsRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.DataTable.DataTableSettingsVerify ? "DataTable Settings is required." : "DataTable Settings is not required.");
                }

                if (tblcolumns.Contains("Data_Table$IsExportRequired"))
                {
                    _workFlow8.DataTable.IsExportRequired = table.Rows[0]["Data_Table$IsExportRequired"].ToString() == "Yes";
                    _extentTest.Log(Status.Pass, _workFlow8.DataTable.IsExportRequired ? "Exports is required." : "Exports is not required.");
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

