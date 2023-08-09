using AngleSharp.Dom;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Model;
using OpenQA.Selenium;
using OpenQA.Selenium.DevTools.V112.Debugger;
using OpenQA.Selenium.Support.UI;
using RazorEngine.Compilation.ImpromptuInterface;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using SHAProject.Workflows;
using SharpCompress.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace SHAProject.Create_Widgets
{
    public class FilesPage
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData? _fileUploadOrExistingFileData;
        public FilesTabData? _FilesTabData;

        public FilesPage(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, FilesTabData FilesTabData)
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _FilesTabData = FilesTabData;
            PageFactory.InitElements(_driver, this);
        }

        #region SearchBox

        [FindsBy(How = How.Id, Using = "Files")]
        public IWebElement? filesTab;

        [FindsBy(How = How.CssSelector, Using = "#filter-text-box")]
        public IWebElement? searchTextBox;

        //[FindsBy(How = How.XPath, Using = "(//div[@col-id=\"Name\"])[2]")]
        //public IWebElement? selectFirstResultedFile;

        //[FindsBy(How = How.XPath, Using = "(//span[@class=\"ag-cell-value\"]/span)[1]")]
        //public IWebElement? selectFirstResultedFile;

        [FindsBy(How = How.XPath, Using = "(//span[@class=\"ag-cell-value\"])[1]")]
        public IWebElement? selectFirstResultedFile;
        #endregion

        #region LayoutVerification Elements

        [FindsBy(How = How.XPath, Using = "//a[@id='Files']")]
        private IWebElement Filetabnavlinkbutton { get; set; }

        [FindsBy(How = How.Id, Using = "idmyfileblock")]
        private IWebElement folderList { get; set; }

        [FindsBy(How = How.Id, Using = "filestab-breadcrumbs")]
        private IWebElement breadcrumsBar { get; set; }

        // header Icons
        [FindsBy(How = How.XPath, Using = "//div[@class='sortbyicons-files']")]
        private IWebElement headerIcons { get; set; }

        [FindsBy(How = How.Id, Using = "divassaylist")]
        private IWebElement filesList { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".ag-paging-panel.ag-unselectable")]
        private IWebElement pagination { get; set; }

        // Header Icon  - Verification

        [FindsBy(How = How.XPath, Using = "//a[@title='Upload a file']")]
        private IWebElement uploadAFile { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Download']")]
        private IWebElement downloadFile { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Make a Copy']")]
        private IWebElement makeACopy { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Move to folder']")]
        private IWebElement moveToFolder { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Delete']")]
        private IWebElement delete { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Add Assay Kit']")]
        private IWebElement addAssayKit { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Export']")]
        private IWebElement export { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Send To']")]
        private IWebElement sendTo { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Rename']")]
        private IWebElement rename { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Favorite a file']")]
        private IWebElement favouriteAFile { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='CatFile']")]
        private IWebElement categories { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='New Assay']")]
        private IWebElement newAssay { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Create a new project']")]
        private IWebElement newProject { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Create a new folder']")]
        private IWebElement newFolder { get; set; }

        //Ag - Grid Header Icons
        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[1]")]
        private IWebElement name { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[2]")]
        private IWebElement listCategories { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[3]")]
        private IWebElement runOnDate { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[4]")]
        private IWebElement lastModified { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[5]")]
        private IWebElement size { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[6]")]
        private IWebElement instrument { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[7]")]
        private IWebElement license { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-cell-label-container'])[8]")]
        private IWebElement favorite { get; set; }

        #endregion

        #region Pagenation Elements

        [FindsBy(How = How.Id, Using = "page-size")]
        private IWebElement pageDropdown { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".ag-paging-description")]
        private IWebElement paginationList { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@Class='ag-paging-page-summary-panel']/div[1]")]
        private IWebElement firstPageIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@Class='ag-paging-page-summary-panel']/div[2]")]
        private IWebElement previousPageIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@Class='ag-paging-page-summary-panel']/span[1]")]
        private IWebElement paginationNumber { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@Class='ag-paging-page-summary-panel']/div[3]")]
        private IWebElement nextPageIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@Class='ag-paging-page-summary-panel']/div[4]")]
        private IWebElement lastPageIcon { get; set; }

        [FindsBy(How = How.Id, Using = "gotopageTextbox-aggrid")]
        private IWebElement textBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='ag-paging-description']/span[last()]")]
        private IWebElement paginationLastNumber { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='ag-paging-description']/span[1]")]
        private IWebElement paginationFirstNumber { get; set; }
        
        [FindsBy(How = How.XPath, Using = "//*[@id='ag-30']/span[2]/div[5]/span/img")]
        private IWebElement textBoxEnter { get; set; }
        #endregion

        #region Search Box Elements
        [FindsBy(How = How.Id, Using = "filter-text-box")]
        private IWebElement searchBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@class='ag-center-cols-viewport']")]
        private IWebElement searchedFiles { get; set; }
        #endregion

        #region New_Assay & Project Elements

        [FindsBy(How = How.XPath, Using = "(//div[@class='modal-content'])[25]")]
        private IWebElement assayTemplate { get; set; }

        [FindsBy(How = How.XPath, Using = "//img[@src='/images/svg/Close-X.svg?v=Dkw16oRIk7HbVG3HvxturCiC0_Wq7NsUv_MQ1dfPJ50']")]
        private IWebElement closeButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id ='btnCreateAssayTemplate']")]
        private IWebElement createButton { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class ='btn btn-default'])[18]")]
        private IWebElement cancelButton { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='modal-content'])[22]")]
        private IWebElement newProjectTemplate { get; set; }

        [FindsBy(How = How.XPath, Using = "(//img[@src='/images/svg/Close-X.svg'])[27]")]
        private IWebElement closeBtn { get; set; }

        #endregion

        #region Create Newfolder Elements

        [FindsBy(How = How.XPath, Using = "(//div[@class='modal-content'])[12]")]
        private IWebElement newFolderTemplate { get; set; }

        [FindsBy(How = How.XPath, Using = "//select[@id='ddlfolder']")]
        private IWebElement folderDropdown { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='txtfoldername']")]
        private IWebElement folderName { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class ='btn btn-default'])[10]")]
        private IWebElement cancelbtn { get; set; }

        [FindsBy(How = How.XPath, Using = "(//input[@class='btn btn-primary'])[2]")]
        private IWebElement createbtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@title='Create a new folder']")]
        private IWebElement lastFolder { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@id='createfolderModal']")]
        private IWebElement CreateFolderPopup { get; set; }

        [FindsBy(How = How.Id, Using = "assayFolderview")]
        private IWebElement createdFolder { get; set; }

        #endregion

        #region FileUpload Elements

        [FindsBy(How = How.XPath, Using = "//span[@class='iconspace -upload']")]
        private IWebElement uploadFileIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//img[@src ='/images/svg/Close-X.svg'])[13]")]
        private IWebElement closeIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@type='file']")]
        private IWebElement filePath { get; set; }

        [FindsBy(How = How.XPath, Using = "//ul[@class='fileList']")]
        private IWebElement fileList { get; set; }

        [FindsBy(How = How.XPath, Using = "//img[@class='closeFileListIcon']")]
        private IWebElement removeFile { get; set; }

        [FindsBy(How = How.XPath, Using = "(//select[@id='ddlfolder_upload'])[1]")]
        private IWebElement uploadFilesInFolder { get; set; }

        [FindsBy(How = How.XPath, Using = "(//input[@class='select2-search__field'])[1]")]
        private IWebElement addCategories { get; set; }
        //select2-search__field valid
        [FindsBy(How = How.XPath, Using = "//input[@class='select2-search__field valid']")]
        private IWebElement enterCategories { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id='importAssayModal']/div/div/div[3]/button[2]")]
        private IWebElement uploadButton { get; set; }
        
        [FindsBy(How = How.XPath, Using = "//span[@id='spanFileCount']")]
        private IWebElement uploadFileCount { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id='importAssayModal']/div/div/div[3]/button[2]")]
        private IWebElement doneButton { get; set; }

        #endregion

        #region Folder Elements

        [FindsBy(How = How.XPath, Using = "//span[@class='ic open-arrow arrowclas']")]
        private IWebElement myFilesArrow { get; set; }

        [FindsBy(How = How.XPath, Using = "(//span[@class='ic opened-arrow'])[2]")] 
        private IWebElement folderArrows { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='ic opened-arrow']")] //span[@class=\"ic open-arrow pointer\"]
        private IWebElement folderArrow { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@id='divassaylist']")]
        private IWebElement filesInFilesList { get; set; }

        #endregion

        #region Rename and Delete Elements

        [FindsBy(How = How.XPath, Using = "//span[@class='ic open-arrow pointer']")]
        private IWebElement lastFolderArrow { get; set; }

        [FindsBy(How = How.XPath, Using = "(//label[@class='tree-toggler nav-header nosubfolder'])[last()]")] 
        private IWebElement renameText { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class ='list-options treeview'])[last()]")]  
        private IWebElement? lastOption { get; set; }

        [FindsBy(How = How.XPath, Using = "(//*[@class='options-popup-folder'])[1]/ul/li[2]")] 
        private IWebElement folderRename { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".rename_event.folder-rename")]
        private IWebElement renameTextBox { get; set; }

        [FindsBy(How = How.XPath, Using = "(//*[@class='options-popup-folder'])[1]/ul/li[1]")]
        private IWebElement folderdelete { get; set; }
        
        [FindsBy(How = How.XPath, Using = "(//div[@class='modal-content'])[21]")]
        private IWebElement confirmationPopUp { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='close'])[22]")]
        private IWebElement confirmationCloseIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@id='btncancel'])[3]")]
        private IWebElement confirmationNoButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='btndelete']")]
        private IWebElement confirmationYesButton { get; set; }

        #endregion

        #region Header Icon Elements

        [FindsBy(How = How.CssSelector, Using = "#filter-text-box")]
        private IWebElement searchbox { get; set; }

        //[FindsBy(How = How.CssSelector, Using = ".filetabAssayimage")]
        //private IWebElement selectFirstResultedFile { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-selection-checkbox'])[1]")] 
        private IWebElement firstCheckbox { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='DowFile']")]
        private IWebElement elements {get; set;}

        [FindsBy(How = How.XPath, Using = "(//span[@class='iconspace -download'])[1]")]
        private IWebElement downloadIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='btn btn-primary'])[11]")]
        private IWebElement okButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='MakeFile']")]
        private IWebElement copy { get; set; }

        [FindsBy(How = How.XPath, Using = "(//span[@class='iconspace -clipboard'])[1]")]
        private IWebElement makeaCopy { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='close'])[7]")]
        private IWebElement makeACopyCloseIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//select[@id='ddlcopyfolder']")]
        private IWebElement copyDropDown { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='btncopyassay']")]
        private IWebElement copyFile { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='MoveToFold']")]
        private IWebElement Movefolder { get; set; }

        [FindsBy(How = How.XPath, Using = "(//span[@class='iconspace -clipboard'])[2]")]
        private IWebElement moveTofolder { get; set; }

        [FindsBy(How = How.XPath, Using = "//select[@id='ddlmovefolder_upload']")]
        private IWebElement foldername { get; set; }

        [FindsBy(How = How.XPath, Using = "(//input[@class='btn btn-primary'])[4]")]
        private IWebElement moveButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@id='moveConfirmModal']")]
        private IWebElement moveConfirmModal { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id='btnmoverename']")]
        private IWebElement renameBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id='btnmovereplace']")]
        private IWebElement replaceBtn { get; set; }

        [FindsBy(How = How.CssSelector, Using = "button#btncancel.btn.btn-default.btn-primary")]
        private IWebElement cancelBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='DelFile']")]
        private IWebElement delelement { get; set; }

        [FindsBy(How =How.XPath, Using = "//span[@class='iconspace -trash']")]
        private IWebElement deleteIcon { get; set; }

        [FindsBy(How =How.XPath, Using = "(//button[@class='btn btn-primary'])[9]")]
        private IWebElement delYesicon { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='AddAssKey']")]
        private IWebElement Assayelement { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='iconspace assykit-Icon']")]
        private IWebElement AssaykitIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='assay-license-part-number-span']")]
        private IWebElement catNumber { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='assay-license-lot-number-span']")]
        private IWebElement lotNumber { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='assay-license-swid-span']")]
        private IWebElement SWID { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@onclick='fnValidateAssayLicense()']")]
        private IWebElement validatebtn { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='close'])[24]")]
        private IWebElement assaykitCloseIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='ExpFile']")]
        private IWebElement Exportelement { get; set; }

        [FindsBy(How = How.XPath, Using = "//li[@title='Excel']")]
        private IWebElement Excelicon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='btn btn-primary'])[10]")]
        private IWebElement excelOkButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//li[@title='Prism']")]
        private IWebElement prismIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='btn btn-primary'])[10]")]
        private IWebElement prismOkButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//li[@title='DQR']")]
        private IWebElement DQRicon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='btn btn-primary'])[10]")]
        private IWebElement DQRokButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='SendTo']")]
        private IWebElement SendToElement { get; set; }

        [FindsBy(How = How.CssSelector, Using = "span.select2-container input.select2-search__field")]
        private IWebElement sendToTextBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@onclick='fnfilestabMailsSend()']")]
        private IWebElement sendButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='RenameFile']")]
        private IWebElement renameIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='newassay']")]
        private IWebElement renameFileName { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='btnrenameassay']")]
        private IWebElement upadteButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='FavFile']")]
        private IWebElement favElement { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='iconspace  favorite-headericon']")]
        private IWebElement favoriteIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "(//div[@class='ag-selection-checkbox'])[1]")]
        private IWebElement CheckBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@id='CatFile']")]
        private IWebElement categoryElement { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='categorySeach']")]
        private IWebElement categoriesTextBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@class='input-group-append']")]
        private IWebElement categoriesFilterIcon { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@id='createlabel']")]
        private IWebElement createNewCategories { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@id='managelabel']")]
        private IWebElement manageLabels { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@onclick='fnCreateNewCategory()']")]
        private IWebElement createNewLabel { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@id='frmcreatecategory']")]
        private IWebElement createNewPopUp { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@onclick='fnCloseCategory()'])[1]")]
        private IWebElement popUpCloseBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@onclick='fnCloseCategory()'])[2]")]
        private IWebElement popUpCancelBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@onkeyup='fnNewCtegoryTextBox(this)']")]
        private IWebElement popUpTextBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='btnnewcategory']")]
        private IWebElement createBtn { get; set; }

        [FindsBy(How =How.XPath, Using = "(//div[@class='ag-selection-checkbox'])[1]")]
        private IWebElement firstChkBox { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[@href='/Manage/MyCategories']")]
        private IWebElement manageLabel { get; set; }

        [FindsBy(How =How.XPath, Using = "(//span[@class='custom-view-edit-icon'])[1]")]
        private IWebElement editIcon { get; set; }

        [FindsBy(How =How.XPath, Using = "//div[@id='editTagOption']")]
        private IWebElement editTag { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='close'])[2]")]
        private IWebElement editTagCloseBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "(//button[@class='btn btn-default'])[2]")]
        private IWebElement editTagCancelBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='customeditinput']")]
        private IWebElement editTagTextBox { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".btn.btn-primary.confirm-update")]
        private IWebElement updateBtn { get; set; }
        #endregion

        #region Test Functions

        public bool FilesPageRedirect()
        {
            try
            {
                _findElements?.ClickElement(Filetabnavlinkbutton, _currentPage, "Home Page -Files tab Navigate link Button");
                //Thread.sleep(5000);
                return true;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Files tab button is not clicked. The error is {e.Message}");
                ScreenShot.ScreenshotNow(_driver, _currentPage, "Error Screenshot", ScreenshotType.Error);
                return false;
            }
        }
        //public bool SearchFilesInFileTab(string fileName)
        //{
        //    try
        //    {
        //        _findElements.ClickElement(filesTab, _currentPage, "Files Tab");

        //        _findElements.SendKeys(fileName, searchTextBox, _currentPage, $"Given file name is - {fileName}");

        //        IReadOnlyCollection<IWebElement> FileList = _driver.FindElements(By.CssSelector(".ag-cell-value"));

        //        foreach (IWebElement File in FileList)
        //        {
        //            if (File.Text.Equals(fileName))
        //            {
        //                Thread.Sleep(2000);
        //                _findElements.ClickElement(File, _currentPage, $"Files Tab - First file");
        //                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Existing file selected successfully");
        //                return true;
        //            }
        //        }

        //        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The given existing file name is not present in the files tab. Upload a new file");
        //        Assert.Fail();
        //        throw new Exception("Existing file was not selected.");
        //        return false;
        //    }
        //    catch (Exception e)
        //    {
        //        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Existing file was not selected.The error is {e.Message}");
        //        return false;
        //    }
        //}

        public bool SearchFilesInFileTab(string fileName)
        {
            try
            {
                _findElements.ClickElement(filesTab, _currentPage, "Files Tab");

                Thread.Sleep(3000);

                _findElements.SendKeys(fileName, searchTextBox, _currentPage, $"Given file name is - {fileName}");

                Thread.Sleep(2000);

                _findElements.ClickElementByJavaScript(selectFirstResultedFile, _currentPage, $"Files Tab - First file");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Existing file selected successfully");
                return true;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Existing file not selected.The error is {e.Message}");
                return false;
            }
        }

        public void LayoutIconsVerification() //Test ID - 1 Layout_Verification
        {
            try
            {
                _findElements.VerifyElement(Filetabnavlinkbutton, _currentPage,"Files Page - Files tab Navigate link Button");

                _findElements.VerifyElement(folderList, _currentPage, "Files Page - Folder List");

                _findElements.VerifyElement(breadcrumsBar, _currentPage, "Files Page - Breadcrums Bar");

                _findElements.VerifyElement(headerIcons, _currentPage, "Files Page - Header Icons");

                _findElements.VerifyElement(filesList, _currentPage, "Files Page - Files List");

                _findElements.VerifyElement(pagination, _currentPage, "Files Page - Pagination");

                _findElements.VerifyElement(uploadAFile, _currentPage, "Files Page - Upload File");

                _findElements.VerifyElement(downloadFile, _currentPage, "Files Page - Download File");

                _findElements.VerifyElement(makeACopy, _currentPage, "Files Page - Copy Files");

                _findElements.VerifyElement(moveToFolder, _currentPage, "Files Page - Move Files to Folder");

                _findElements.VerifyElement(delete, _currentPage, "Files Page - Delete Files");

                _findElements.VerifyElement(addAssayKit, _currentPage, "Files Page - Add Assay Kit to Files");

                _findElements.VerifyElement(export, _currentPage, "Files Page - Export Files");

                _findElements.VerifyElement(sendTo, _currentPage, "Files Page - Send Files");

                _findElements.VerifyElement(rename, _currentPage, "Files Page - Rename Files");

                _findElements.VerifyElement(favouriteAFile, _currentPage, "Files Page - Add Favourites to the Files");

                _findElements.VerifyElement(categories, _currentPage, "Files Page - Categories");

                _findElements.VerifyElement(newAssay, _currentPage, "Files Page - Add new Assay");

                _findElements.VerifyElement(newProject, _currentPage, "Files Page - Create new Project");

                _findElements.VerifyElement(newFolder, _currentPage, "Files Page - Add new Folder");

                _findElements.VerifyElement(name, _currentPage, "Files Page - File Name");

                _findElements.VerifyElement(listCategories, _currentPage, "Files Page - Categories");

                _findElements.VerifyElement(runOnDate, _currentPage, "Files Page - Run-On-Date");

                _findElements.VerifyElement(lastModified, _currentPage, "Files Page - last Modified");

                _findElements.VerifyElement(size, _currentPage, "Files Page - Size");

                _findElements.VerifyElement(instrument, _currentPage, "Files Page - Instrument Type");

                _findElements.VerifyElement(license, _currentPage, "Files Page - License");

                _findElements.VerifyElement(favorite, _currentPage, "Files Page - Favorite");

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Error Occured While verifiying the icons.The error is {e.Message}");
            }
        }

        public void PagenationVerificattion() //Test ID - 2 Pagenation
        {
            try
            {
                _findElements.VerifyElement(pageDropdown, _currentPage, "Files Page - Page Dropdown");

                _findElements.VerifyElement(paginationList, _currentPage, "Files Page - Pagination List");

                _findElements.VerifyElement(firstPageIcon, _currentPage, "Files Page - First Page Icon");

                _findElements.VerifyElement(previousPageIcon, _currentPage, "Files Page - Previous Page Icon");

                _findElements.VerifyElement(paginationNumber, _currentPage, "Files Page - Pagination Number");

                _findElements.VerifyElement(nextPageIcon, _currentPage, "Files Page - Favorite");

                _findElements.VerifyElement(lastPageIcon, _currentPage, "Files Page - Favorite");

                _findElements.VerifyElement(textBox, _currentPage, "Files Page - Favorite");

                _findElements.ClickElementByJavaScript(lastPageIcon, _currentPage, "Files page - Last page icon");
                //Thread.sleep(1000);

                _findElements.ClickElementByJavaScript(firstPageIcon, _currentPage, "Files page - First Page Icon");
                //Thread.sleep(1000);

                _findElements.ClickElementByJavaScript(nextPageIcon, _currentPage, "Files page - Next Page Icon");
                //Thread.sleep(1000);

                _findElements.ClickElementByJavaScript(previousPageIcon, _currentPage, "Files page - Previous Page Icon");
                //Thread.sleep(1000);

                _findElements.VerifyElement(paginationLastNumber, _currentPage, "Files Page - pagination Last page Number");

                _findElements.ClickElementByJavaScript(paginationLastNumber, _currentPage, "Files Page - pagination Last page Number");
                //Thread.sleep(1000);

                _findElements.VerifyElement(paginationFirstNumber, _currentPage, "Files Page - pagination First page Number");

                _findElements.ClickElementByJavaScript(paginationFirstNumber, _currentPage, "Files Page - Pagination first Number");
                textBox?.SendKeys(_FilesTabData?.PageNumber);

                _findElements.VerifyElement(textBoxEnter, _currentPage, "Files Page - Pagination First page Number");

                _findElements.ClickElementByJavaScript(textBoxEnter, _currentPage, "Files Page - pagination text box");

                //SelectElement select = new(pageDropdown);
                //select.SelectByText(_FilesTabData?.FilesList); 

                _findElements.SelectFromDropdown(pageDropdown, _currentPage, "text", _FilesTabData?.FilesList, $"File list - {_FilesTabData?.FilesList}");

            }
            catch(Exception e) 
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Error Occured While verifiying the Pagenation icons.The error is {e.Message}");
            }
        }

        public void SearchboxVerification() //Test ID - 3 SearchBox
        {
            try
            {
                var list = _FilesTabData.searchBoxDataList;
                for (int i = 0; i < list.Count; i++)
                {
                    searchBox.SendKeys(list[i]);
                    string message;
                    switch (i)
                    {
                        case 0:
                            message = "File First Name";
                            break;
                        case 1:
                            message = "File Middle Name";
                            break;
                        case 2:
                            message = "File Last Name";
                            break;
                        case 3:
                            message = "File Full Name";
                            break;
                        case 4:
                            message = "Categories";
                            break;
                        case 5:
                            message = "Date";
                            break;
                        case 6:
                            message = "Instrument";
                            break;
                        case 7:
                            message = "License";
                            break;
                        default:
                            message = "File First Name"; 
                            break;
                    }
                   _findElements.VerifyElement(searchedFiles, _currentPage, "Files Page - Files Searched by " + message);
                    searchBox.Clear();
                }
            }
            catch(Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Error Occured While verifiying the Searchbox.The error is {e.Message}");
            }
        }

        public void CreateNewAssayandProject()
        {
            try
            {
                // Create Assay Template
                _findElements?.VerifyElement(newAssay, _currentPage, "Files Page - Add new Assay");

                _findElements?.ClickElementByJavaScript(newAssay, _currentPage, "Files Page - pagination Last page Number");
                //Thread.sleep(2000);

                _findElements?.VerifyElement(assayTemplate, _currentPage, "Files Page - Assay Template");

                _findElements?.VerifyElement(closeButton, _currentPage, "Files Page - Assay Template Close Button");

                _findElements?.VerifyElement(createButton, _currentPage, "Files Page - Assay Template Create Button");

                _findElements?.VerifyElement(cancelButton, _currentPage, "Files Page - Assay Template Cancel Button");
                //Thread.sleep(1000);

                _findElements?.ClickElementByJavaScript(cancelButton, _currentPage, "Cancel Button");

                //Create New Project
                _findElements?.VerifyElement(newProject, _currentPage, "Files Page - New Project");
                //Thread.sleep(1000);

                _findElements?.ClickElementByJavaScript(newProject, _currentPage, "New Project");
                //Thread.sleep(2000);

                _findElements?.VerifyElement(newProjectTemplate, _currentPage, "Files Page - New Project Template");

                _findElements?.VerifyElement(closeBtn, _currentPage, "Files Page - New Project Template Close Button");

                _findElements?.ClickElementByJavaScript(closeBtn, _currentPage, "Files Page - Close Button");

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Error Occured While verifiying the Create New Assay and Project Popup.The error is {e.Message}");
            }
        }

        public void CreateNewFolder()
        {
            try
            {
                _findElements?.VerifyElement(newFolder, _currentPage, "Files Page - Add new Folder");

                _findElements?.ClickElementByJavaScript(newFolder, _currentPage,"New Folder");

                //Thread.sleep(2000);
                _findElements?.VerifyElement(newFolderTemplate, _currentPage, "Files Page - Add new Folder Template Popup");

                _findElements?.VerifyElement(folderDropdown, _currentPage, "Files Page - Add new Folder Template folder Dropdown");

                _findElements?.VerifyElement(folderName, _currentPage, "Files Page - Add new Folder Template folder name text box");

                _findElements?.VerifyElement(cancelbtn, _currentPage, "Files Page - Add new Folder Template Cancel button");

                _findElements?.VerifyElement(createbtn, _currentPage, "Files Page - Add new Folder Template Create button");

                string[] folderNames = { _FilesTabData.FolderName, _FilesTabData.SubFolderName, _FilesTabData.LastFolderName };
                bool isSubFolderCreated = false;
                bool duplicatefolder = false;
                foreach (string name in folderNames)
                {
                    string display = CreateFolderPopup.GetCssValue("display");
                    if (display == "none" && !isSubFolderCreated)
                    {
                        _findElements?.ClickElementByJavaScript(newFolder, _currentPage, "New Folder");
                        //Thread.sleep(2000);
                        folderName.SendKeys(_FilesTabData.SubFolderName);
                        _findElements?.VerifyElement(folderName, _currentPage, "Files Page - Add new Folder Template new Foldre Name is Entered");
                        //Thread.sleep(1000);
                        _findElements?.ClickElementByJavaScript(createbtn, _currentPage, "Create Button");
                        try
                        {
                            IAlert alert = _driver.SwitchTo().Alert();
                            alert.Accept();
                            duplicatefolder = true;
                            //Thread.sleep(2000);
                        }
                        catch (Exception ex)
                        {

                        }
                        isSubFolderCreated = true;
                    }
                    else if (display == "block" || isSubFolderCreated)
                    {
                        if (isSubFolderCreated)
                        {
                            _findElements?.ClickElementByJavaScript(lastFolder, _currentPage, "Last Folder");
                        }
                        //Thread.sleep(3000);
                        string foldername = display == "block" ? _FilesTabData.FolderName : _FilesTabData.LastFolderName;
                        //Create the new folder
                        folderName.SendKeys(foldername);
                        _findElements?.VerifyElement(folderName, _currentPage, "Files Page - Add new Folder Template new Folder Name is Entered");
                        //Thread.sleep(1000);
                        _findElements?.ClickElementByJavaScript(createbtn, _currentPage, "Create Button");
                        try
                        {
                            IAlert alert = _driver.SwitchTo().Alert();
                            alert.Accept();
                            duplicatefolder = true;
                            //Thread.sleep(2000);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    if (!duplicatefolder) {
                        _findElements?.VerifyElement(createdFolder, _currentPage, "Files Page - Newly Created Folder" + name);
                    }
                    else
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Folder and Sub folders is already created in the given name " + name);
                    }
                    //Thread.sleep(3000);
                }

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Error, $"Error Occured While Creating a new Folder .The error is {e.Message}");
            }
        }

        public void fileUpload()
        {
            try
            {
                _findElements?.VerifyElement(uploadFileIcon, _currentPage, "Files Page - Add new Folder");

                _findElements?.ClickElementByJavaScript(uploadFileIcon, _currentPage, "Upload File Icon");
                //Thread.sleep(3000);

                _findElements?.VerifyElement(closeIcon, _currentPage,"Files Page - Upload File popup Close button");
               
                string fileNames = _FilesTabData.FileName;

                string[] singleFileName = fileNames.Split(',');
                for (int i = 0; i < singleFileName.Length; i++)
                {
                    string fileName = singleFileName[i].Trim();
                    if (fileName != "")
                    {
                        singleFileName[i] = _FilesTabData.FileUploadPath + "\\" + fileName;
                    }
                }

                string newFileNames = string.Join(",", singleFileName);

                string[] filePaths = newFileNames.Split(',');
                string newPath = string.Join("\n", filePaths);
                filePath?.SendKeys(newPath);
                //Thread.sleep(1000);

                _findElements?.VerifyElement(fileList,_currentPage, "Files Page - Selected File added in the file list");

                _findElements?.VerifyElement(removeFile, _currentPage, "Files Page - Remove File Icon");

                _findElements?.ClickElementByJavaScript(removeFile, _currentPage, "Files page - First Page Icon");

                _findElements?.VerifyElement(fileList, _currentPage, "Files Page - Selected file is Deleted");

                _findElements?.VerifyElement(uploadFilesInFolder, _currentPage, "Files Page - Selected files are in folder");

                uploadFilesInFolder.SendKeys(_FilesTabData.FileLocatedFolderPath);
                //Thread.sleep(1000);

                //IWebElement addCategories = driver.FindElement(By.XPath("(//input[@class=\"select2-search__field\"])[1]"));
                addCategories.SendKeys(Keys.Enter);
                for(int i =0; i< _FilesTabData.AddCategories.Length; i++)
                {
                    string category = _FilesTabData.AddCategories[i].ToString();
                    enterCategories.SendKeys(category);
                }
                enterCategories.SendKeys(Keys.Enter);
                //Thread.sleep(3000);

                _findElements?.VerifyElement(uploadButton, _currentPage, "Files Page - Upload File folder");

                _findElements.ClickElementByJavaScript(uploadButton, _currentPage, "Upload Button");
                //Thread.sleep(5000);

                //Thread.sleep(5000);
                _findElements?.VerifyElement(uploadFileCount, _currentPage, "Files Page - Upload File Count");

                _findElements?.VerifyElement(fileList, _currentPage, "Files Page - Valid and Invalid Files");

                _findElements?.VerifyElement(doneButton, _currentPage, " Files Page - Upload File Template - Done Button");

                _findElements?.ClickElementByJavaScript(doneButton, _currentPage, "Done Button"); 

                _findElements?.VerifyElement(searchedFiles, _currentPage, "Files Page - Valid files are added in the files list");
                //IWebElement searchedFiles = driver.FindElement(By.XPath("//div[@class='ag-center-cols-viewport']"));
                //ScreenshotNow(ScreenshotPath, currentPage, "Valid files are added in the files list", ScreenshotType.Info, searchedFiles);
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Error, $"Error Occured While Creating a new Folder .The error is {ex.Message}");
            }
        }

        public void folderFunctions()
        {
            try
            {
                _findElements?.VerifyElement(myFilesArrow, _currentPage, "Files Page - My File Arrow Icon");
                //Thread.sleep(3000);

                _findElements?.VerifyElement(folderArrow, _currentPage, "Files Page - Folder Arrow Icon");

                _findElements?.ClickElementByJavaScript(folderArrow, _currentPage, "Folder Arrow");

                _findElements?.VerifyElement(folderArrows, _currentPage, "Files Page - Sub Folder Arrow Icon");
                //Thread.sleep(3000);

                IWebElement folder = _driver.FindElement(By.XPath("//label[@title='" + _FilesTabData.FolderName + "']"));

                _findElements?.ClickElementByJavaScript(folder, _currentPage, "Folder");

                _findElements?.VerifyElement(filesInFilesList, _currentPage, "Files Page - Files in the folder");

                IWebElement subFolder = _driver.FindElement(By.XPath("//label[@title='" + _FilesTabData.SubFolderName + "']"));
                if (!(subFolder.Displayed))
                {
                    _findElements?.ClickElementByJavaScript(folderArrow, _currentPage, "Folder Arrow");
                }
                ExtentReport.ExtentTest("ExtentTestNode", subFolder.Displayed ? Status.Pass : Status.Fail, subFolder.Displayed ? "Sub Folder is displayed" : "Sub Folder is not displayed");

                _findElements?.ClickElementByJavaScript(subFolder, _currentPage, "Sub Folder");

                _findElements?.VerifyElement(filesInFilesList, _currentPage, "Files Page - Files in the sub folder");

            }
            catch (Exception ex)
            { 
                ExtentReport.ExtentTest("ExtentTestNode", Status.Error, $"Error Occured While Creating a new Folder .The error is {ex.Message}");
            }
        }

        public void renameAnddelete()
        {
            try
            {
                if (!myFilesArrow.Displayed)
                {
                    _findElements?.ClickElementByJavaScript(myFilesArrow, _currentPage, "My FIles arrow");
                }
                ExtentReport.ExtentTest("ExtentTestNode", myFilesArrow.Displayed ? Status.Pass : Status.Fail, myFilesArrow.Displayed ? "Folder arrow is displayed" : "Folder Arrow is not displayed");
                //Thread.sleep(3000);

                if (!folderArrow.Displayed)
                {
                    _findElements?.ClickElementByJavaScript(folderArrow, _currentPage, "Files Page - Folder Arrow Icon");
                }
                ExtentReport.ExtentTest("ExtentTestNode", folderArrow.Displayed ? Status.Pass : Status.Fail, folderArrow.Displayed ? "Sub Folder arrow is displayed" : "Sub Folder Arrow is not displayed");
                _findElements?.ClickElementByJavaScript(folderArrow, _currentPage, "Files Page - Folder Arrow Icon");

                //Thread.sleep(3000);
                if (!lastFolderArrow.Displayed)
                {
                    _findElements?.ClickElementByJavaScript(lastFolderArrow, _currentPage, "Files Page -Last Folder Arrow Icon");
                }
                ExtentReport.ExtentTest("ExtentTestNode", lastFolderArrow.Displayed ? Status.Pass : Status.Fail, lastFolderArrow.Displayed ? "Last Folder arrow is displayed" : "Last Folder Arrow is not displayed");

                _findElements?.ClickElementByJavaScript(renameText, _currentPage, "Rename Text");

                _findElements?.VerifyElement(lastOption, _currentPage,"Files Page - Three dot icon");

                lastOption?.Click();

                _findElements?.VerifyElement(folderRename, _currentPage, "Files Page - Rename Option");

                folderRename?.Click();

                renameTextBox.SendKeys(Keys.Control + "a"); // Select all text
                renameTextBox.SendKeys(Keys.Backspace);
                renameTextBox.SendKeys(_FilesTabData.Rename);

                lastOption.Click();
                //Thread.sleep(4000);
                lastOption.Click();

                _findElements?.VerifyElement(folderdelete, _currentPage, "Files Page - Delete Option");
              
                folderdelete.Click();

                //Thread.sleep(1000);

                _findElements?.VerifyElement(confirmationPopUp, _currentPage, "Files Page - Confirmation PopUp");

                _findElements?.VerifyElement(confirmationCloseIcon, _currentPage, "Files Page - Confirmation PopUp close icon");

                _findElements?.VerifyElement(confirmationNoButton, _currentPage, "Files Page - Confirmation PopUp No Button");

                _findElements?.VerifyElement(confirmationNoButton, _currentPage, "Files Page - Confirmation PopUp Yes Button");

                confirmationYesButton.Click();
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Error, $"Error Occured While Creating a new Folder .The error is {ex.Message}");

            }
        }
        public void HeaderIconFunction()
        {
            try
            {
                searchbox.Clear();  /*Clear the search box*/
                searchbox.SendKeys(_FilesTabData.FileFullName);  /*Enter the file name in the search box*/
                //Thread.sleep(2000);
              
                if (selectFirstResultedFile != null)
                {
                    _FilesTabData.FileFullName = selectFirstResultedFile.Text;
                    firstCheckbox.Click();
                    //Thread.sleep(5000);
                    ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "File was selected successfully");
                }

                if (_FilesTabData.DownloadFileVerification)
                {
                    try
                    {
                        //Download Icon Functionality
                        string Opacity = elements.GetCssValue("opacity");

                        // Check if the CSS property has the expected value
                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and downloaded header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and downloaded header icon is already enabled ");
                        }
                        _findElements?.ClickElementByJavaScript(downloadIcon, _currentPage, "Download");
                        //Thread.sleep(3000);

                        _findElements?.ClickElementByJavaScript(okButton, _currentPage,"Ok Button");
                        //Thread.sleep(3000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "The download file icon verification has been failed. The error is -" + ex.Message);
                       // ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.MakeACopy)
                {
                    try
                    {
                        string Opacity = copy.GetCssValue("opacity");

                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and Make a copy header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing file was selected and Make a copy header icon is already enabled ");
                        }
                        //Make a Copy Icon Functionality
                        _findElements?.ClickElementByJavaScript(makeaCopy, _currentPage, "Make a copy");
                        //Thread.sleep(2000);

                        _findElements?.VerifyElement(makeACopyCloseIcon, _currentPage, "Files Page - Make A Copy Close Icon");

                        //SelectElement dropdown = new(copyDropDown);
                        //dropdown.SelectByText(_FilesTabData.CopyFilePath);
                        _findElements.SelectFromDropdown(copyDropDown, _currentPage, "text", _FilesTabData.CopyFilePath, $"Copied File Path - {_FilesTabData.CopyFilePath}" );

                        //Thread.sleep(2000);

                        _findElements?.ClickElementByJavaScript(copyFile, _currentPage, "Copy File");

                        _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                        //Thread.sleep(20000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "The make a copy icon verification has been failed. The error is -" + ex.Message);
                        // ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.MoveToFolder)
                {
                    try
                    {
                        string opacity = Movefolder.GetCssValue("opacity");

                        if (opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files were selected, and the Move to Folder header icon is enabled.");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files were selected, and the Move to Folder header icon is already enabled.");
                        }

                        // Move to Folder Icon Functionality
                        _findElements?.ClickElementByJavaScript(moveTofolder, _currentPage, "Move to folder");
                        //Thread.sleep(3000);

                        SelectElement dropdown = new SelectElement(foldername);
                        dropdown.SelectByText(_FilesTabData.FolderPath);

                        _findElements.SelectFromDropdown(foldername, _currentPage, "text", _FilesTabData.FolderPath, $"Folder Path - {_FilesTabData.FolderPath}");
                        //Thread.sleep(5000);

                        _findElements.ClickElementByJavaScript(moveButton, _currentPage, "Move Button");

                        // Need to verify the existing file name - Replace or Rename
                        _findElements?.VerifyElement(moveConfirmModal, _currentPage, "Files Page - Rename or Replace file popup");

                        _findElements?.VerifyElement(renameBtn, _currentPage, "Files Page - Rename Button");

                        _findElements?.VerifyElement(replaceBtn, _currentPage, "Files Page - Replace Button");

                        _findElements?.VerifyElement(cancelBtn, _currentPage, "Files Page - Cancel Button");

                        if (_FilesTabData.ReplaceOrRename == "Rename")
                        {
                            _findElements?.ClickElementByJavaScript(renameBtn, _currentPage, "Rename Button");
                        }
                        else
                        {
                            _findElements?.ClickElementByJavaScript(replaceBtn, _currentPage, "Replace Button");
                        }

                        //Thread.sleep(10000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Move to the Folder Icon verification has failed. The error is - " + ex.Message);
                        //ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.DeleteFile)
                {
                    try
                    {
                        string Opacity = delelement.GetCssValue("opacity");

                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and delete file header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and delete file header icon is already enabled ");
                        }

                        //Delete Icon Functionality
                        _findElements.ClickElementByJavaScript(deleteIcon, _currentPage, "Delete Icon");
                        //Thread.sleep(3000);

                        _findElements.ClickElementByJavaScript(delYesicon, _currentPage, "Delete Yes Icon");
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Delete The File Icon verification has been Passed ");
                        //Thread.sleep(15000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Delete The File Icon verification has been failed. The error is -" + ex.Message);
                        //ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.AssayKitVerification)
                {
                    try
                    {
                        string Opacity = Assayelement.GetCssValue("opacity");

                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files was selected and Assay kit - header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and Assay kit - header icon is already enabled ");
                        }

                        //Add Assay kit functionality
                        _findElements?.ClickElementByJavaScript(AssaykitIcon, _currentPage, "Assay Kit Icon");

                        _findElements?.VerifyElement(catNumber, _currentPage, "Files Page - Assay Kit validation - Cat Number");
                        //Thread.sleep(2000);

                        /*Apply Lot Number*/
                        _findElements?.VerifyElement(lotNumber, _currentPage, "Files Page - Assay Kit validation - Lot Number");
                        //Thread.sleep(2000);

                        /*Apply SWID Number*/
                        _findElements?.VerifyElement(SWID, _currentPage, "Files Page -Assay Kit validation - SWID Number");
                        //Thread.sleep(2000);

                        /*Click on the Validate Button*/
                        _findElements?.VerifyElement(validatebtn, _currentPage, "Files Page -Assay Kit validation - Validation Button");
                        //Thread.sleep(2000);

                        _findElements?.ClickElementByJavaScript(assaykitCloseIcon, _currentPage, "Assay kit close icon");
                        //Thread.sleep(3000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Assay Kit Icon verification has been failed. The error is -" + ex.Message);
                       // ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.ExportFilesVerification)
                {
                    try
                    {
                        string Opacity = Exportelement.GetCssValue("opacity");

                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files was selected and Export - header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files was selected and Export - header icon is already enabled ");
                        }

                        //Export Icon Functionality
                        _findElements?.ClickElementByJavaScript(Exportelement, _currentPage, "Export Element");

                        _findElements?.ClickElementByJavaScript(Excelicon, _currentPage, "Excel Icon");
                        //Thread.sleep(1000);

                        _findElements?.ClickElementByJavaScript(excelOkButton, _currentPage,"Excel Ok Button");
                        //Thread.sleep(1000);

                        _findElements?.ClickElementByJavaScript(prismIcon, _currentPage, "Prism Icon");

                        _findElements?.ClickElementByJavaScript(prismOkButton, _currentPage, "Prism Ok Button");
                        //Thread.sleep(1000);

                        _findElements?.ClickElementByJavaScript(Exportelement, _currentPage, "Export Element");

                        _findElements?.ClickElementByJavaScript(DQRicon, _currentPage, "DQR Icon");

                        _findElements?.ClickElementByJavaScript(DQRokButton, _currentPage, "DQR Ok button");
                        //Thread.sleep(3000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Export The Files Icon verification has been failed. The error is -" + ex.Message);
                       // ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.SendToVerfication)
                {
                    try
                    {
                        string Opacity = SendToElement.GetCssValue("opacity");

                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files was selected and Send To - header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files was selected and Send To - header icon is already enabled ");
                        }

                        //SendTo Functionality
                        //Thread.sleep(1000);
                        _findElements?.ClickElementByJavaScript(SendToElement, _currentPage, "Send to Element");
                        //Thread.sleep(2000);

                        //sendToTextBox.Clear();
                        string emailId = _FilesTabData.FirstMailRecepient;
                        // IWebElement mailtext = driver.FindElement(By.CssSelector(".select2-selection"));
                        IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                        jScript.ExecuteScript("arguments[0].value = arguments[1];", sendToTextBox, emailId);
                        sendToTextBox.SendKeys(" ");
                        sendToTextBox.SendKeys(" ");
                        sendToTextBox.SendKeys(Keys.Enter);
                        jScript.ExecuteScript("arguments[0].dispatchEvent(new KeyboardEvent('keyup', {key: 'Enter'}));", sendToTextBox);
                        //Thread.sleep(3000);

                        _findElements?.ClickElementByJavaScript(sendButton, _currentPage, "Send button");
                        //Thread.sleep(5000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Send To Icon verification has been failed. The error is -" + ex.Message);
                       // ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.RenameVerification)
                {
                    try
                    {
                        string Opacity = renameIcon.GetCssValue("opacity");

                        if (Opacity == "1")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files was selected and Rename - header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and Rename - header icon is already enabled ");
                        }

                        //RenameFile Icon Functionality
                        _findElements?.ClickElementByJavaScript(renameIcon, _currentPage, "Rename Icon");
                        //Thread.sleep(2000);

                        //Thread.sleep(2000);
                        renameFileName.Clear();
                        renameFileName.SendKeys("Titration SmokeA 10_0_1_53");
                        //Thread.sleep(1000);

                        _findElements?.ClickElementByJavaScript(upadteButton, _currentPage, "Update Button");
                        //Thread.sleep(5000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Rename the Files Icon verification has been failed. The error is -" + ex.Message);
                        //ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.AddFavorite)
                {
                    try
                    {
                        string Opacity = favElement.GetCssValue("opacity");

                        if (Opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and Favorite header icon is enabled ");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "The existing files was selected and Favorite - header icon is already enabled ");
                        }

                        //FavoriteIcon Functionality
                        _findElements?.ClickElementByJavaScript(favoriteIcon, _currentPage, "Fav Icon");
                        //Thread.sleep(3000);
                        CheckBox.Click();
                        //Thread.sleep(3000);
                        _findElements?.ClickElementByJavaScript(favoriteIcon, _currentPage, "Fav Icon");
                        //Thread.sleep(15000);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Fail, "Add Favorite Icon verification has been failed. The error is -" + ex.Message);
                       // ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

                if (_FilesTabData.AddCategory)
                {
                    try
                    {
                        string opacity = categoryElement.GetCssValue("opacity");

                        if (opacity == "0.5")
                        {
                            firstCheckbox.Click();
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files were selected, and Category - header icon is enabled");
                        }
                        else
                        {
                            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The existing files were selected, and Category - header icon is already enabled");
                        }
                        // Scroll to the categoriesIcon element
                        IJavaScriptExecutor jScript = (IJavaScriptExecutor)_driver;
                        jScript.ExecuteScript("arguments[0].scrollIntoView(true);", categoryElement);

                        // Wait for the obscuring element to become invisible
                        WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
                        wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//input[@id='categorySeach']"))); // Replace ... with the locator of the obscuring element

                        // Click on the categoriesIcon element
                        _findElements?.ClickElementByJavaScript(categoryElement, _currentPage, "Category Element");

                        _findElements?.VerifyElement(categoriesTextBox, _currentPage, "Files Page - Categories Text Box");
                        //Thread.sleep(3000);

                        _findElements?.VerifyElement(categoriesFilterIcon, _currentPage, "Files Page - Categories Filter Icon");

                        _findElements?.VerifyElement(createNewCategories, _currentPage, "Files Page - Categories Create New label Popup");

                        _findElements?.VerifyElement(manageLabels, _currentPage, "Files Page - Categories Manage labels");

                        createNewLabel.Click();

                        _findElements?.VerifyElement(createNewPopUp, _currentPage, "Files Page - Create New Label Popup");

                        _findElements?.VerifyElement(popUpCloseBtn, _currentPage, "Files Page - Create New Label Popup Close button");

                        _findElements?.VerifyElement(popUpCancelBtn, _currentPage, "Files Page - Create New Label Popup Cancel Button");

                        _findElements?.VerifyElement(popUpTextBox, _currentPage, "Files Page - Create New Label Popup Text Box");

                        popUpTextBox.SendKeys(_FilesTabData.AddCategoryName);

                        ScreenShot.ScreenshotNow(_driver, _currentPage, "Category Name - " + _FilesTabData.AddCategoryName, ScreenshotType.Info, popUpTextBox);
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The Given Category Name is - " + _FilesTabData.AddCategoryName);

                        _findElements?.ClickElementByJavaScript(createBtn, _currentPage, "Create Button");
                        //Thread.sleep(15000);

                        firstChkBox.Click();

                        _findElements?.ClickElementByJavaScript(categoryElement, _currentPage, "Category Element");

                        _findElements?.ClickElementByJavaScript(manageLabel, _currentPage, "Manage Label");
                        //Thread.sleep(3000);

                        _findElements?.ClickElementByJavaScript(editIcon, _currentPage, "Edit Icon");
                        //Thread.sleep(2000);

                        _findElements?.VerifyElement(editTag, _currentPage, "Files Page - Edit Tag PopUp");

                        _findElements?.VerifyElement(editTagCloseBtn, _currentPage, "Files Page - Edit Tag Close button");

                        _findElements?.VerifyElement(editTagCancelBtn, _currentPage, "Files Page - Edit Tag Cancel Button");

                        _findElements?.VerifyElement(editTagTextBox, _currentPage, "Files Page - Edit Tag Text Box");
                        editTagTextBox.Clear();
                        editTagTextBox.SendKeys(_FilesTabData.EditCategoryName);
                        ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, "The Given Edit Category Name is - " + _FilesTabData.EditCategoryName);
                        //Thread.sleep(2000);

                        _findElements?.ClickElementByJavaScript(updateBtn, _currentPage, "Update Button");
                        //Thread.sleep(2000);

                        _driver.SwitchTo().Alert().Accept();

                        IWebElement updatedCategoryName = _driver.FindElement(By.XPath("//span[@title='" + _FilesTabData.EditCategoryName + "')]"));
                        ExtentReport.ExtentTest("ExtentTestNode.Log",updatedCategoryName.Displayed ? Status.Pass : Status.Fail, $"Updated Category Name is {(updatedCategoryName.Displayed ? "displayed" : "not displayed")}");
                        ScreenShot.ScreenshotNow(_driver, _currentPage, "Updated Category Name - " + _FilesTabData.EditCategoryName, ScreenshotType.Info, updatedCategoryName);
                    }
                    catch (Exception ex)
                    {
                        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Add Categories Icon verification has been failed. The error is -" + ex.Message);
                        //ScreenshotNow(ScreenshotPath, currentPage, "Error Screenshot", ScreenshotType.Error);
                    }
                }

            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, "Add Categories Icon verification has been failed. The error is -" + ex.Message);
            }
        }
        #endregion
    }
}
