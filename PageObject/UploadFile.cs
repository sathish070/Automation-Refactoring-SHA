using AventStack.ExtentReports;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.PageObject
{
    public class UploadFile
    {

        public IWebDriver? _driver;
        public FindElements? _findElements;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;

        public UploadFile(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData)
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements =findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            PageFactory.InitElements(_driver, this);
        }

        [FindsBy(How = How.XPath, Using = "//a[@title='Upload a file']")]
        public IWebElement? uploadFileButton;

        [FindsBy(How = How.XPath, Using = "//input[@type='file']")]
        public IWebElement? browseButton;

        [FindsBy(How = How.ClassName, Using = "box__button")]
        public IWebElement? uploadButton;

        [FindsBy(How = How.XPath, Using = "//button[@class='box__button btn btn-primary']")]
        public IWebElement? Donebutton;

        [FindsBy(How = How.CssSelector, Using = "#myGrid .ag-row:first-child .ag-cell:first-child")]
        public IWebElement? firstFile;

        [FindsBy(How = How.Id, Using = "Files")]
        public IWebElement? filesTab;

        [FindsBy(How = How.CssSelector, Using = "#filter-text-box")]
        public IWebElement? searchTextBox;

        [FindsBy(How = How.XPath, Using = "(//span[@class=\"ag-cell-value\"]/span)[1]")]
        public IWebElement? selectFirstResultedFile;

        [FindsBy(How = How.CssSelector, Using = ".file-count")]
        public IWebElement? UploadFileCount;

        public bool HomePageFileUpload()
        {
            try
            {
                _findElements.ClickElement(uploadFileButton, _currentPage, $"Home Page -Upload File Button");

                string folderPath = Environment.OSVersion.Platform == PlatformID.MacOSX || Environment.OSVersion.Platform == PlatformID.Unix ? _fileUploadOrExistingFileData.FileUploadPath + "/" + _fileUploadOrExistingFileData.FileName + "." + _fileUploadOrExistingFileData.FileExtension :
                _fileUploadOrExistingFileData.FileUploadPath + "\\" + _fileUploadOrExistingFileData.FileName + "." + _fileUploadOrExistingFileData.FileExtension;

                browseButton?.SendKeys(folderPath);

                _findElements.ClickElement(uploadButton, _currentPage, $"Upload File Popup - Upload Button");

                _findElements.WaitForElementVisible(UploadFileCount);

                _findElements.ClickElement(Donebutton, _currentPage, $"Upload File Popup-Done Button");

                ScreenShot.ScreenshotNow(_driver, _currentPage, "Given File", ScreenshotType.Info, firstFile);

                _findElements.ClickElement(firstFile, _currentPage, $"Uploaded as recent File ");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"File upload successfully");
                return true;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"File not upload successfully. The error is {e.Message}");

                ScreenShot.ScreenshotNow(_driver, _currentPage, "Error Screenshot", ScreenshotType.Error);
                return false;
            }
        }

        public bool SearchFilesInFileTab(string FileName)
        {
            try
            {
                _findElements.ClickElement(filesTab, _currentPage,"Files Tab" );

                _findElements.SendKeys(FileName, searchTextBox, _currentPage, $"Given file name is - {_fileUploadOrExistingFileData.FileName}");

                _findElements.ClickElement(selectFirstResultedFile, _currentPage, $"Files Tab - First file");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Existing file selected successfully");
                return true;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Existing file not selected.The error is {e.Message}");
                return false;
            }
        }
    }
}
