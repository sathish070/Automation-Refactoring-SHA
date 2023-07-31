using AventStack.ExtentReports;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.PageObject
{
    public class ModifyAssay
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;
        public CommonFunctions? _commonFunc;
        public ModifyAssay(string currentPage, IWebDriver driver, FindElements findElements, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc) 
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements = findElements;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _commonFunc = commonFunc;
            PageFactory.InitElements(_driver, this);
        }

        #region Modify Assay Header Tabs

        [FindsBy(How = How.XPath, Using = "//a[@href=\"#groups\"]")]
        public IWebElement GroupTab;

        [FindsBy(How = How.XPath, Using = "//a[@href=\"#plateMap\"]")]
        public IWebElement PlateMap;

        [FindsBy(How = How.XPath, Using = "//a[@href=\"#media_buffer\"]")]
        public IWebElement AssayMedia;

        [FindsBy(How = How.XPath, Using = "//a[@href=\"#tab3\"]")]
        public IWebElement BackgroundBuffer;

        [FindsBy(How = How.XPath, Using = "//a[@href=\"#Injection_Names\"]")]
        public IWebElement InjectionNames;

        [FindsBy(How = How.XPath, Using = "//a[@href=\"#tab4\"]")]
        public IWebElement GeneralInfo;

        [FindsBy(How = How.XPath, Using = "//a[@title='Modify Assay']")]
        public IWebElement ModifyAssayTab;

        #endregion

        #region Group Tab Elements

        [FindsBy(How = How.XPath, Using = "//*[@title ='Expand/Collapse All']")]
        public IWebElement ExpandIcon;

        [FindsBy(How = How.XPath, Using = "(//*[@class ='row grouprow imagetoggle'])[2]")]
        public IWebElement GroupExpansion;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"row1\"]/div/div[1]/div[1]")]
        public IWebElement InjectionStrategy;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"row1\"]/div/div[1]/div[2]")]
        public IWebElement Pretreatment;

        [FindsBy(How = How.XPath, Using = "//*[@id='row1']/div/div[2]/div[1]")]
        public IWebElement Assaymedia;

        [FindsBy(How = How.XPath, Using = "//*[@id='row1']/div/div[2]/div[2]")]
        public IWebElement CellType;

        [FindsBy(How = How.XPath, Using = "(//span[@class=\"imgarea\"])[2]")]
        public IWebElement ExpansionIcon;

        [FindsBy(How = How.XPath, Using = "//*[@title ='Move Selection Down']")]
        public IWebElement MoveSelectionDown;

        [FindsBy(How = How.XPath, Using = "//*[@title ='Move Selection Up']")]
        public IWebElement MoveSelectionUp;

        [FindsBy(How = How.Id, Using = "addgrp-btn")]
        public IWebElement AddGroupBtn;

        [FindsBy(How = How.XPath, Using = "(//div[@class ='row grouprow selected'])[last()]")]
        public IWebElement AddedGroup;

        [FindsBy(How = How.XPath, Using = "//*[@class=\"row grouprow selected\"]/span[3]/a/img")]
        public IWebElement DotIcon;

        [FindsBy(How = How.XPath, Using = "(//*[@onclick=\"grouplistrename(this)\"])[last()]")]
        public IWebElement RenameButton;

        [FindsBy(How = How.XPath, Using = "(//label[@onblur=\"ChangeGroupName(this)\"])[last()]")]
        public IWebElement GroupRename;

        [FindsBy(How = How.XPath, Using = "//*[@onclick=\"fnModifyDialog()\"]")]
        public IWebElement SaveButton;

        [FindsBy(How = How.CssSelector, Using = "[src=\"/images/svg/Modify.svg\"]")]
        public IWebElement Modifyassay;

        [FindsBy(How = How.XPath, Using = "(//*[@onblur=\"ChangeGroupName(this)\"])[last()]")]
        public IWebElement DuplicateGroupName;

        [FindsBy(How = How.XPath, Using = "(//*[@onclick=\"grouplistdelete(this)\"])[last()]")]
        public IWebElement DeleteBtn;

        [FindsBy(How = How.XPath, Using = "(//*[@onclick=\"grouplistrename(this)\"])[last()]")]
        public IWebElement ButtonLast;

        [FindsBy(How = How.XPath, Using = "(//*[@onclick=\"grouplistdelete(this)\"])[last()]")]
        public IWebElement DeleteButton;

        #endregion

        #region Plate Map Elements

        [FindsBy(How = How.XPath, Using ="//*[@class =\"col-md-5 platemap-groups\"]")]
        public IWebElement GroupList;

        [FindsBy(How = How.Id, Using = "plate-map-table")]
        public IWebElement PlateMapTable;

        [FindsBy(How = How.XPath, Using ="(//*[@Class=\"list-options groupoption\"])[last()]")]
        public IWebElement LastGroupList;

        [FindsBy(How = How.XPath, Using ="(//select[@Class=\"set-group-ctrl\"])[last()]")]
        public IWebElement DropdownControlGroups;

        [FindsBy(How = How.XPath, Using ="//*[@class =\"col-md-12 platemapArea\"]")]
        public IWebElement PlateMapArea;

        [FindsBy(How = How.XPath, Using ="(//td[@data-wellnum=\"4\"])[2]")]
        public IWebElement WellDataPopup;

        #endregion

        #region Assay Media Elements

        [FindsBy(How = How.XPath, Using ="(//*[@class ='row form-group']/label)[2]")]
        public IWebElement Name;

        [FindsBy(How = How.XPath, Using ="(//*[@class =\"col-md-7 assy-medianame\"])[1]")]
        public IWebElement NameTextBox;

        [FindsBy(How = How.XPath, Using ="(//*[@class=\"form-group col-md-7\"])[1]")]
        public IWebElement MediaType;

        [FindsBy(How = How.XPath, Using ="(//*[@class=\"form-group col-md-5\"])[1]")]
        public IWebElement BufferFactor;

        [FindsBy(How = How.XPath, Using ="(//*[@id =\"btnApplytoallGroup1\"])[1]")]
        public IWebElement ApplyToAllGroups;

        [FindsBy(How = How.XPath, Using ="(//*[@class ='imgarea'])[2]")]
        public IWebElement Groupexpansion;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"ddlassaymedia1\"]")]
        public IWebElement AssayMediaDropdown;

        #endregion

        #region Background Buffer

        [FindsBy(How = How.XPath, Using = "//*[@id=\"bufferTable\"]/thead/tr/th[1]")]
        public IWebElement Well;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"bufferTable\"]/thead/tr/th[2]")]
        public IWebElement UseDefaultBF;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"bufferTable\"]/thead/tr/th[3]")]
        public IWebElement BufferFactorSetting;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"bufferTable\"]/thead/tr/th[4]")]
        public IWebElement SelectAllText;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"0_0\"]/td[1]")]
        public IWebElement BackgroundWell1;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"0_11\"]/td[1]")]
        public IWebElement BackgroundWell2;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"7_0\"]/td[1]")]
        public IWebElement BackgroundWell3;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"7_11\"]/td[1]")]
        public IWebElement BackgroundWell4;

        [FindsBy(How = How.XPath, Using = "//*[@id='listgroup_1']")]
        public IWebElement BackgroundSelection;

        [FindsBy(How = How.XPath, Using = "(//div[@class='boxes'])[1]")]
        public IWebElement UnselectDefaultBF;

        [FindsBy(How = How.XPath, Using ="(//*[@class=\"boxes\"])[1]")]
        public IWebElement UnselectDefaultBFChkBox;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"chkselectall\"]")]
        public IWebElement UnselectAllCheckBox;


        #endregion

        #region Injection Names

        [FindsBy(How = How.CssSelector, Using = "[src=\"/images/svg/Modify.svg\"]")]
        public IWebElement ModifyAssayIcon;

        [FindsBy(How = How.CssSelector, Using = "//div[@class='col-md-9']/table]")]
        public IWebElement InjTable;

        [FindsBy(How = How.CssSelector, Using = "//*[@class=\"ClassInjectionNames\"]")]
        public IWebElement InjectionCount;

        [FindsBy(How = How.XPath, Using = "(//input[@class='ClassInjectionNames'])[1]")]
        public IWebElement InjectionRename;

        //[FindsBy(How = How.XPath, Using ="//button[@onclick='fnModifyDialog()']")]
        //public IWebElement SaveBtn;

        [FindsBy(How = How.XPath, Using = "(//button[@class=\"btn btn-primary\"])[12]")]
        public IWebElement SaveBtn;

        [FindsBy(How = How.CssSelector, Using = "//button[@id=\"btnIsLinkedToProjects\"]")]
        public IWebElement ContinueBtn;

        #endregion

        #region General Info

        [FindsBy(How = How.XPath, Using = "//*[@id=\"tab4\"]/div/div[1]/div[1]/label[1]")]
        public IWebElement ProjectName;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"tab4\"]/div/div[1]/div[1]/label[2]")]
        public IWebElement PrincipalInvestigator;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"tab4\"]/div/div[1]/div[1]/label[3]")]
        public IWebElement ProjectNumber;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"tab4\"]/div/div[1]/div[1]/label[4]")]
        public IWebElement WellVolume;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"tab4\"]/div/div[1]/div[1]/label[5]")]
        public IWebElement PlatedBy;

        [FindsBy(How = How.XPath, Using ="//*[@id=\"tab4\"]/div/div[1]/div[1]/label[6]")]
        public IWebElement PlatedOn;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"tab4\"]/div/div[1]/div[2]/label")]
        public IWebElement Notes;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"tab4\"]/div/div[2]")]
        public IWebElement AssayInformation;

        #endregion

        public void ModifyAssayHeaderTabs()
        {
            try
            {
                _findElements?.ClickElementByJavaScript(ModifyAssayTab, _currentPage, $"Modify Assay - Icon Button");

                _findElements.ElementTextVerify(GroupTab, "Groups", _currentPage, $"Modify Assay - Group Tab");

                _findElements.ElementTextVerify(PlateMap, "Plate Map", _currentPage, $"Modify Assay - Plate Map");

                _findElements.ElementTextVerify(AssayMedia, "Assay Media", _currentPage, $"Modify Assay - Assay Media");

                _findElements.ElementTextVerify(BackgroundBuffer, "Background Buffer", _currentPage, $"Modify Assay - Background Buffer");

                _findElements.ElementTextVerify(InjectionNames, "Injection Names", _currentPage, $"Modify Assay - Injection Names");

                _findElements.ElementTextVerify(GeneralInfo, "General Info", _currentPage, $"Modify Assay - General Info");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay header tabs has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Modify assay header tabs has not been verified. The error is {e.Message}");
            }
        }

        public void GroupTabElements(string groupName)
        {
            try
            {
                _findElements.ClickElementByJavaScript(ExpandIcon, _currentPage, $"Group Tab - Expand Icon");

                _findElements.VerifyElement(GroupExpansion, _currentPage, $"Group Tab - Group Expansion");

                _findElements.ElementTextVerify(InjectionStrategy, "Injection Strategy", _currentPage, $"Group Expansion - Injection Strategy");

                _findElements.ElementTextVerify(Pretreatment, "Pretreatment", _currentPage, $"Group Expansion - Pretreatment");

                _findElements.ElementTextVerify(Assaymedia, "Assay Media", _currentPage, $"Group Expansion - Assay Media");

                _findElements.ElementTextVerify(CellType, "Cell Type", _currentPage, $"Group Expansion - Cell Type");

                _findElements.ClickElementByJavaScript(ExpansionIcon, _currentPage, $"Group Tab - Group Expansion back to position"); // Expand/Collapse tab is back to normal

                _findElements.ClickElementByJavaScript(MoveSelectionDown, _currentPage, $"Group Tab - Move Selection Down");

                _findElements.ClickElementByJavaScript(MoveSelectionUp, _currentPage, $"Group Tab - Move Selection Up");

                _findElements.ElementTextVerify(AddGroupBtn, "Add Group", _currentPage, $"Group Tab - Add Group Button");

                if (_fileUploadOrExistingFileData.IsTitrationFile == false)
                {
                    AddGroupBtn.Click();

                    _findElements.VerifyElement(AddedGroup, _currentPage, $"New Group Added");

                    _findElements.ClickElementByJavaScript(DotIcon, _currentPage, $"Add Group - Three Dot Icon");

                    _findElements.ActionsClassClick(RenameButton);

                    GroupRename.SendKeys(Keys.End);
                    while (GroupRename.Text.Length > 0)
                    {
                        GroupRename.SendKeys(Keys.Backspace);
                    }
                    Thread.Sleep(2000);

                    _findElements.SendKeys(groupName, GroupRename, _currentPage, $"Given group name is  {groupName}");

                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Group Name added in the list");

                    _findElements.ClickElementByJavaScript(SaveButton, _currentPage, $"Add Group - Save Button");

                    ScreenShot.ScreenshotNow(_driver, _currentPage, "Group Added in the Analysis Page", ScreenshotType.Info);
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, "Group Name added is verified in the Analysis Page");

                    _findElements.ClickElementByJavaScript(Modifyassay, _currentPage, $"Modify Assay Button");

                    AddGroupBtn.Click();

                    _findElements.ActionsClassClick(RenameButton);

                    _findElements.ClickElementByJavaScript(DotIcon, _currentPage, $"Add Group - Three Dot Icon");

                    DuplicateGroupName.SendKeys(Keys.End);
                    while (DuplicateGroupName.Text.Length > 0)
                    {
                        DuplicateGroupName.SendKeys(Keys.Backspace);
                    }
                    _findElements.SendKeys(groupName, GroupRename, _currentPage, $"Given group name is {groupName}");

                    DuplicateGroupName.SendKeys(Keys.Tab);

                    _driver.SwitchTo().Alert().Accept();

                    _findElements.ScrollIntoViewAndClickElementByJavaScript(DeleteBtn, _currentPage, $"Delete Button");

                    _findElements.ScrollIntoViewAndClickElementByJavaScript(ButtonLast, _currentPage, $"Last Button");

                    _findElements.ClickElementByJavaScript(DotIcon, _currentPage, $"Add Group - Three Dot Icon");

                    _findElements.ScrollIntoViewAndClickElementByJavaScript(DeleteButton, _currentPage, $"Delete Button");
                }

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay group tab has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay group tab has not been verified .The error is {e.Message}");
            }
        }

        public void PlateMapElements(string selectTheControls)
        {
            try
            {
                _findElements.ClickElement(PlateMap, _currentPage, $"Plate Map Tab");

                _findElements.VerifyElement(GroupList, _currentPage, $"Plate Map tab - Group List");

                _findElements.VerifyElement(PlateMapArea, _currentPage, $"Plate Map tab - Plate Map table");

                _findElements.ScrollIntoViewAndClickElementByJavaScript(LastGroupList, _currentPage, $"Last Group List");

                _findElements.ClickElement(DropdownControlGroups, _currentPage, $"Select the control popup");

                int selectedIndex = selectTheControls == "Set Group as Positive Control" ? 1 : selectTheControls == "Set Group as Negative Control" ? 2 :
                                    selectTheControls == "Set Group as Vehicle Control" ? 3 : 0;

                //_findElements.SelectByIndex(DropdownControlGroups, selectedIndex);

                _findElements.SelectFromDropdown(DropdownControlGroups, _currentPage, "index", selectedIndex.ToString() , "Select the Controls" );

                _findElements.VerifyElement(PlateMapArea, _currentPage, $"Selected controls in Plate Map Area ");

                string elementId = "Groudetail7";
                string script = $"document.getElementById('{elementId}').style.display = 'block';";
                ((IJavaScriptExecutor)_driver).ExecuteScript(script);

                ScreenShot.ScreenshotNow(_driver, _currentPage, "WellData popup in the platemap tab", ScreenshotType.Info);

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - plate map tab has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - plate map tab has not been verified. The error is {e.Message}");
            }
        }

        public void AssayMediaElements()
        {
            try
            {
                _findElements.ClickElement(AssayMedia, _currentPage, $"Modify Assay - Assay Media");

                _findElements.VerifyElement(Name, _currentPage, $"Assay Media - Name");

                _findElements.VerifyElement(MediaType, _currentPage, $"Assay Media - Media Type");

                _findElements.ClickElementByJavaScript(ApplyToAllGroups, _currentPage, $"Modify Assay - Apply To All groups");

                _findElements.ClickElement(GroupTab, _currentPage, $"Modify Assay - Group Tab");

                _findElements.ScrollIntoViewAndClickElementByJavaScript(Groupexpansion, _currentPage, $"Group Tab - Group Expansion");

                _findElements.ClickElementByJavaScript(AssayMediaDropdown, _currentPage, $"Group Tab - Assay Media drop down");

                Groupexpansion.Click();

                AssayMedia.Click();

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - assay media tab has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - assay media tab has not been verified. The error is {e.Message}");
            }
        }

        public void BackgroundBufferElements()
        {
            try
            {
                _findElements.ClickElement(BackgroundBuffer, _currentPage, $"Modify Assay - Background Buffer");

                _findElements.ElementTextVerify(Well, "Well", _currentPage, $"BackgroundBuffer - Well");

                _findElements.ElementTextVerify(UseDefaultBF, "Use Default BF", _currentPage, $"BackgroundBuffer - Use Default BF");

                _findElements.ElementTextVerify(BufferFactorSetting, "Buffer Factor Setting", _currentPage, $"BackgroundBuffer - Buffer Factor Setting");

                /*Select All Check Box*/

                _findElements.ElementTextVerify(SelectAllText, "Select all", _currentPage, $"Select all Check box");

                _findElements.VerifyElement(BackgroundWell1, _currentPage, $"Background well -1");

                _findElements.VerifyElement(BackgroundWell2, _currentPage, $"Background well -2");

                _findElements.VerifyElement(BackgroundWell3, _currentPage, $"Background well -3");

                _findElements.VerifyElement(BackgroundWell4, _currentPage, $"Background well -4");

                //extentTestNode.Log(Status.Pass, "Background Assigend Well Names are displayed");

                /*Click on the background group in the plate map tab*/
                _findElements.ClickElement(PlateMap, _currentPage, $"Plate Map Tab");

                _findElements.ClickElementByJavaScript(BackgroundSelection, _currentPage, $"Modify Assay - Background Selection");

                IWebElement selectionBackground = null;
                for (int i = 1; i < 3; i++)
                {
                    selectionBackground = _driver.FindElement(By.Id("ctrl_" + i));
                    _findElements.ClickElementByJavaScript(selectionBackground, _currentPage, $"Modify Assay - Background Selections");
                }

                //extentTestNode.Log(selectionBackground.Displayed ? Status.Pass : Status.Fail, selectionBackground.Displayed ? "New background well names are added" : "New background well names are not added");

                _findElements.ClickElement(BackgroundBuffer, _currentPage, $"Modify Assay - Background Selections"); /* Back to Background Buffer Tab*/


                _findElements.ClickElement(UnselectDefaultBF, _currentPage, $"Unselect Default BF");

                //extentTestNode.Log(!unselectDefaultBF.Selected ? Status.Pass : Status.Fail, !unselectDefaultBF.Selected ? "DefaultBF checkbox is unselected" : "DefaultBF checkbox is not unselected");

                //jScript.ExecuteScript("arguments[0].click();", unselectDefaultBF); /* Select the DefaultBF*/

                _findElements.ClickElementByJavaScript(UnselectAllCheckBox, _currentPage, $"Unselect All Chck Box");

                //extentTestNode.Log(unselectAllCheckBox.Displayed ? Status.Pass : Status.Fail, unselectAllCheckBox.Displayed ? "All DefaultBF checkbox are unselected" : "All DefaultBF checkbox are not unselected");

                _findElements.ClickElementByJavaScript(UnselectAllCheckBox, _currentPage, $"Unselect All Chck Box");/* Select all the Checkbox*/

                //IWebElement closeIcon = driver.FindElement(By.XPath("(//img[@src=\"/images/svg/Close-X.svg\"])[8]"));
                //jScript.ExecuteScript("arguments[0].click();", closeIcon);

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - background buffer tab has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - background buffer tab has not been verified. The error is {e.Message}");
            }
        }

        public void InjectionNamesElements(string injectionName)
        {
            try
            {
                //_findElements.ClickElementByJavaScript(ModifyAssayIcon, _currentPage, $"Modify Assay Icon ");

                _findElements.ClickElement(InjectionNames, _currentPage, $"Modify Assay - Injection Names ");

                InjectionRename.Clear();

                _findElements.SendKeys(injectionName, InjectionRename, _currentPage,$"Modify Assay - Injection Count ");

                _findElements.ClickElementByJavaScript(SaveBtn, _currentPage, $"Injection Name -Save button");

                Thread.Sleep(2000);

                if (_driver.PageSource.Contains("Continue"))
                {
                    _findElements.ClickElementByJavaScript(ContinueBtn, _currentPage, $"Injection Name - Continue Button");

                }

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - injection name tab has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - injection name tab has not been verified. The error is {e.Message}");
            }
        }

        public void GeneralInfoElements()
        {
            try
            {
                _findElements.ClickElementByJavaScript(ModifyAssayIcon, _currentPage, $"Modify Assay Icon ");

                /*Click on the general info tab*/
                Thread.Sleep(2000);

                _findElements.ClickElementByJavaScript(GeneralInfo, _currentPage, $"Modify Assay -General Info ");

                _findElements.VerifyElement(ProjectName, _currentPage, $"General Info - project Name");

                _findElements.VerifyElement(PrincipalInvestigator, _currentPage, $"General Info - Principal Investigator");

                _findElements.VerifyElement(ProjectNumber, _currentPage, $"General Info - Project Number");

                _findElements.VerifyElement(WellVolume, _currentPage, $"General Info - Well Volume");

                _findElements.VerifyElement(PlatedBy, _currentPage, $"General Info - PlateBy");

                _findElements.VerifyElement(PlatedOn, _currentPage, $"General Info - PlateOn");

                _findElements.VerifyElement(Notes, _currentPage, $"General Info - Notes");

                _findElements.VerifyElement(AssayInformation, _currentPage, $"General Info - Assay Information");

                _findElements.ClickElementByJavaScript(SaveBtn, _currentPage, $"General Info - Save Button");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Modify assay - general info tab has been verified");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Modify assay - general info tab has not been verified. The error is {e.Message}");
            }
        }
    }
}
