using AventStack.ExtentReports;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SHAProject.PageObject;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
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

        [FindsBy(How = How.CssSelector, Using = "[title='Normalize']")]
        public IWebElement NormalizationIcon;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"ModalNormalizeSetting\"]/div")]
        public IWebElement NormalizationPopup;

        [FindsBy(How = How.Id, Using = "normunit")]
        public IWebElement NormalizationUnits;

        [FindsBy(How = How.Id, Using = "scalefactor")]
        public IWebElement ScaleFactorField;

        [FindsBy(How = How.Id, Using = "chkselectallwidget")]
        public IWebElement ApplyWidgetsBtn;

        [FindsBy(How = How.CssSelector, Using = ".normalization-footer .btn-primary")]
        public IWebElement SaveBtn;

        [FindsBy(How = How.Id, Using = "divnorm1")]
        public IWebElement NormalizedWell;

        [FindsBy(How = How.XPath, Using = "//button[@class='btn btn-default btn-import']")]
        public IWebElement ReimportButton;

        [FindsBy(How = How.CssSelector, Using = ".normalization-property")]
        public IWebElement? NormalizationField;

        public void NormalizationElements()
        {

            try
            {
                Thread.Sleep(5000);

                _commonFunc.HandleCurrentWindow();

                _findElements.ClickElement(NormalizationIcon, _currentPage, "Analysis Page - Normalization Icon");

                _findElements?.VerifyElement(NormalizedWell, _currentPage, "Normalized Well with Default Value Present in it");

                NormalizedWell.Clear();

                NormalizedWell.SendKeys("20");

                _findElements?.ClickElementByJavaScript(SaveBtn, _currentPage, "Normalization - Save button");

                Thread.Sleep(8000);
                _findElements.ClickElement(NormalizationIcon, _currentPage, "Analysis Page - Normalization Icon");

                _findElements?.ClickElementByJavaScript(ReimportButton, _currentPage, "Normalization - Re Import button");

                _findElements?.VerifyElement(NormalizedWell, _currentPage, "Normalized Well With the Old Value");
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, "Error Occured while verfiying the File with the Normalization Concept: Message"+ex.Message);
            }

        }

        public void ApplyNormalizationValues(bool ApplyToAllWidgets)
        {
            try
            {
                _findElements.ClickElement(NormalizationIcon, _currentPage, "Analysis Page - Normalization Icon");

                _findElements.VerifyElement(NormalizationPopup, _currentPage, $"Normalization Popup");

                _findElements.SendKeys(_normalizationData.Units, NormalizationUnits, _currentPage, "Given normalization unit is - " + _normalizationData.Units);

                _findElements.SendKeys(_normalizationData.ScaleFactor, ScaleFactorField, _currentPage, "Given scale factor is - " + _normalizationData.ScaleFactor);

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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The error occured in Apply Normalization Icon functionality. The error is { e.Message }");
            }
        }

        public void NormalizationToggle()
        {
            try
            {
                _findElements.VerifyElement(NormalizationField, _currentPage, "Normalization toggle");

                bool Toogle = NormalizationField.Selected;
                ExtentReport.ExtentTest("ExtendTestNode", Toogle ? Status.Pass : Status.Fail, Toogle ? "Normalization Toddled in Enabled": "Normalization Toddled in Disabled");
            }
            catch (Exception ex)
            {
                ExtentReport.ExtentTest("ExtendTestNode", Status.Fail, "Error Occured while verfiying the File with the Normalization Concept: Message" + ex.Message);
            }
        }
    }
}
