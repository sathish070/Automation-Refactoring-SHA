using AventStack.ExtentReports;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.Page_Object
{
    public class Normalization : Tests
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public string _currentPage = string.Empty;
        public NormalizationData _normalizationData;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;
        public CommonFunctions? _commonFunc;

        public Normalization(string currentPage, IWebDriver driver, FindElements findElements, NormalizationData normalizationData, FileUploadOrExistingFileData fileUploadOrExistingFileData, CommonFunctions commonFunc)
        {
            _driver = driver;
            _currentPage = currentPage;
            _findElements = findElements;
            _normalizationData = normalizationData;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            _commonFunc = commonFunc;
            PageFactory.InitElements(_driver, this);
        }

        #region Normalization PopUp Elements

        [FindsBy(How = How.CssSelector, Using = "[title='Normalize']")]
        public IWebElement NormalizationIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"ModalNormalizeSetting\"]/div")]
        public IWebElement NormalizationPopup;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"ModalNormalizeSetting\"]/div/div/span")]
        public IWebElement EditNormalizationText;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"col-md-4 form-group normalization-units\"])[1]")]
        public IWebElement NormalizationField;

        [FindsBy(How = How.Id, Using = "normunit")]
        public IWebElement NormalizationUnitsTextBox;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"col-md-4 form-group normalization-units\"])[2]")]
        public IWebElement ScaleFactorField;

        [FindsBy(How = How.Id, Using = "scalefactor")]
        public IWebElement ScaleFactorTextBox;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"col-md-10 normalization-head\"]/p")]
        public IWebElement NormalizationValuesHeading;

        [FindsBy(How = How.CssSelector, Using = ".normalization-table")]
        public IWebElement NormalizationValuesTable;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"col-md-2 form-group normalization-action-btns\"]/button[1]")]
        public IWebElement SelectAllBtn;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"col-md-2 form-group normalization-action-btns\"]/button[2]")]
        public IWebElement ClearAllDataBtn;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"col-md-2 form-group normalization-action-btns\"]/span")]
        public IWebElement CtrlVPasteLabel;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"form-group checkAllGroup\"]")]
        public IWebElement ApplyToAllWidgetsLabel;

        [FindsBy(How = How.XPath, Using = "//label[@for=\"chkselectallwidget\"]")]
        public IWebElement ApplyToAllWidgetsChkBox;

        [FindsBy(How = How.Id, Using = "chkselectallwidget")]
        public IWebElement ApplyWidgetsBtn;

        [FindsBy(How = How.CssSelector, Using = ".normalization-footer .btn-primary")]
        public IWebElement SaveBtn;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"modal-footer normalization-footer\"]/button[2]")]
        public IWebElement CancelBtn;

        [FindsBy(How = How.Id, Using = "divnorm1")]
        public IWebElement NormalizedWell;

        [FindsBy(How = How.XPath, Using = "//button[@class='btn btn-default btn-import']")]
        public IWebElement ReimportButton;

        [FindsBy(How = How.CssSelector, Using = ".normalization-property")]
        public IWebElement? NormalizationToggleBtn;

        #endregion

        public void NormalizationElementsVerification()
        {
            try
            {
                Thread.Sleep(5000);

                _commonFunc.HandleCurrentWindow();

                _findElements.ClickElement(NormalizationIcon, _currentPage, "Analysis Page - Normalization Icon");

                _findElements.VerifyElement(NormalizationPopup, _currentPage, $"Normalization Popup");

                _findElements.ElementTextVerify(EditNormalizationText, "Edit Normalization", _currentPage, $"Normalization Heading - {EditNormalizationText.Text}");

                _findElements.VerifyElement(NormalizationField, _currentPage, $"Units heading text - {NormalizationField.Text}");

                _findElements.VerifyElement(ScaleFactorField, _currentPage, $"Scale Factor heading text -{ScaleFactorField.Text}");

                _findElements.ElementTextVerify(NormalizationValuesHeading, "Normalization Values", _currentPage, $"Normalization table heading name - {NormalizationValuesHeading.Text}");

                _findElements.VerifyElement(NormalizationValuesTable, _currentPage, $"Normalization Values Table");

                _findElements.ElementTextVerify(SelectAllBtn, "Select All", _currentPage, $"Normalization button -{SelectAllBtn.Text}");

                _findElements.ElementTextVerify(ClearAllDataBtn, "Clear All Data", _currentPage, $"Normalization buttons -{ClearAllDataBtn.Text}");

                _findElements.ElementTextVerify(CtrlVPasteLabel, "Ctrl-V to Paste", _currentPage, $"Normalization label -{CtrlVPasteLabel.Text}");

                _findElements.ElementTextVerify(ApplyToAllWidgetsLabel, "Apply to all Widgets", _currentPage, $"Normalization label -{ApplyToAllWidgetsLabel.Text}");

                if (!ApplyToAllWidgetsChkBox.Selected)
                    _findElements.VerifyElement(ApplyToAllWidgetsChkBox, _currentPage, $"Apply to all widgets check box is unselected");
                else
                    ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, $"Apply to all widgets check box is selected");

                _findElements.VerifyElement(SaveBtn, _currentPage, $"Normalization buttons - {SaveBtn.Text}");

                _findElements.VerifyElement(CancelBtn, _currentPage, $"Normalization buttons - {CancelBtn.Text}");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, $"Error occured while verfiying the file with the Normalization Popup elements. The error is {e.Message}");
            }
        }

        public void NormalizationElements()
        {
            try
            {
                _findElements?.VerifyElement(NormalizedWell, _currentPage, "Normalized Well with Default Value Present in it");

                NormalizedWell.Clear();

                NormalizedWell.SendKeys("20");

                _findElements?.ClickElementByJavaScript(SaveBtn, _currentPage, "Normalization - Save button");

                Thread.Sleep(8000);

                _findElements.ClickElement(NormalizationIcon, _currentPage, "Analysis Page - Normalization Icon");

                _findElements?.ClickElementByJavaScript(ReimportButton, _currentPage, "Normalization - Re Import button");

                _findElements?.VerifyElement(NormalizedWell, _currentPage, "Normalized Well With the Old Value");

                _findElements?.ClickElementByJavaScript(SaveBtn, _currentPage, "Normalization - Save button");
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, "Error occured while verfiying the file with the Normalization Concept: Message" + ex.Message);
            }
        }

        public void ApplyNormalizationValues(bool ApplyToAllWidgets)
        {
            try
            {
                //_findElements.ClickElement(NormalizationIcon, _currentPage, "Analysis Page - Normalization Icon");

                _findElements.SendKeys(_normalizationData.Units, NormalizationUnitsTextBox, _currentPage, "Given normalization unit is - " + _normalizationData.Units);

                _findElements.SendKeys(_normalizationData.ScaleFactor, ScaleFactorTextBox, _currentPage, "Given scale factor is - " + _normalizationData.ScaleFactor);

                for (int i = 0; i < _normalizationData.Values.Count; i++)
                {
                    var NormWells = _driver.FindElement(By.Id("divnorm" + i));
                    NormWells.Clear();
                    NormWells.SendKeys(_normalizationData.Values[i].Trim());
                }

                if (ApplyToAllWidgets && !ApplyWidgetsBtn.Selected)
                    _findElements.ClickElementByJavaScript(ApplyWidgetsBtn, _currentPage, "Normalization - Apply all widgets button");

                _findElements.ClickElementByJavaScript(SaveBtn, _currentPage, "Normalization - Save button");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $" Error occurred while verifiying apply normalization values functionality. The error is { e.Message }");
            }
        }

        public void NormalizationToggle()
        {
            try
            {
                _findElements.VerifyElement(NormalizationToggleBtn, _currentPage, "Normalization toggle");

                bool Toogle = NormalizationToggleBtn.Selected;
                ExtentReport.ExtentTest("ExtendTestNode", Toogle ? Status.Pass : Status.Fail, Toogle ? "Normalization Toggle in Enabled" : "Normalization Toggle in Disabled");
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, "Error occured while verfiying the File with the Normalization Concept: Message" + ex.Message);
            }
        }
    }
}
