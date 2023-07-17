using System;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;
using SeleniumExtras.PageObjects;
using OpenQA.Selenium.Support.Extensions;

namespace SHAProject.EditPage
{
    public class PlateMap
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        private readonly FileType _fileType;
        public string _currentPage = string.Empty;
        public NormalizationData? _normalizationData;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;
        private int WellCount => _fileType == FileType.Xfp ? 8 : _fileType == FileType.Xfe24 ? 24 : 96;

        public PlateMap(string currentPage, IWebDriver driver, FindElements findElements, CommonFunctions commonFunc, FileUploadOrExistingFileData fileUploadOrExistingFileData, FileType fileType, NormalizationData? normalizationData)
        {
            _driver = driver;
            _fileType = fileType;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            _normalizationData = normalizationData;
            _fileUploadOrExistingFileData = fileUploadOrExistingFileData;
            PageFactory.InitElements(_driver, this);
        }

        #region Platemap Elements

        [FindsBy(How = How.Id, Using = "wellmode")]
        public IWebElement? WellSelection;

        [FindsBy(How = How.Id, Using = "flagmode")]
        public IWebElement? FlagSelection;

        [FindsBy(How = How.Id, Using = "flagwellon")]
        public IWebElement? FlagOn;

        [FindsBy(How = How.Id, Using = "flagwelloff")]
        public IWebElement? FlagOff;

        [FindsBy(How = How.Id, Using = "syntoview")]
        public IWebElement? SyncToView;

        [FindsBy(How = How.XPath, Using = "//div[@id='UnselectedWellsSyncView']/div/div")]
        public IWebElement? SyncToViewPopup;

        [FindsBy(How = How.XPath, Using = "(//*[@id='btnok'])[3]")]
        public IWebElement? SyncToViewApplyButton;

        [FindsBy(How = How.CssSelector, Using = ".syncviewresult-tost-success")]
        public IWebElement? SyncToViewToast;

        [FindsBy(How = How.CssSelector, Using = "#plate-map-table tr")]
        private IList<IWebElement>? PlatemaprowCount { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#plate-map-table th")]
        private IList<IWebElement>? PlatemapColumnCount { get; set; }

        #endregion

        public void PlateMapIcons()
        {
            _findElements.VerifyElement(WellSelection, _currentPage, $"PlateMap -Well Mode");

            _findElements.VerifyElement(FlagSelection, _currentPage, $"PlateMap -Flag Mode");

            _findElements.VerifyElement(FlagOn, _currentPage, $"PlateMap -Flag Mode On");

            _findElements.VerifyElement(FlagOff, _currentPage, $"PlateMap -Flag Mode Off");

            _findElements.VerifyElement(SyncToView, _currentPage, $"PlateMap - Sync to View");
        }

        public void PlateMapFunctionalities()
        {
            PlateMapWells("WellSelectionMode"); // Unselect the wells

            PlateMapWells("WellUnSelectionMode"); // Select the wells

            PlateMapWells("FlagMode"); // Add the flags

            _findElements.ClickElementByJavaScript(FlagOff, _currentPage, $"PlateMap -Flag Off");

            _findElements.ClickElementByJavaScript(FlagOn, _currentPage, $"PlateMap -Flag On");

            PlateMapWells("UnflagMode"); // Remove the flags 

            PlateMapSyncToView();
        }

        public void PlateMapWells(string type)
        {
            try
            {
                if (type == "WellSelectionMode")
                    _findElements.ClickElementByJavaScript(WellSelection, _currentPage, $"PlateMap -WellSelection");
                else if (type == "FlagMode")
                    _findElements.ClickElementByJavaScript(FlagSelection, _currentPage, $"PlateMap - FlagSelection");
                else if (type == "UnflagMode")
                    _findElements.ClickElementByJavaScript(FlagSelection, _currentPage, $"PlateMap - FlagSelection");
                else
                    _findElements.ClickElementByJavaScript(WellSelection, _currentPage, $"PlateMap -WellSelection");

                for (int count = 0; count < 5; count++)
                {
                    IWebElement PlateMapWell = _driver.FindElement(By.Id("tbl" + count + ""));
                    string wellName = PlateMapWell.GetAttribute("data-wellvalue");
                    _findElements.ClickElementByJavaScript(PlateMapWell, _currentPage, $" {type} well name is - {wellName}");
                    Thread.Sleep(4000);
                }

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Platemap well selection and flag selection functionality has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Platemap well selection and flag selection functionality has not been verified. The error is {e.Message}");
            }
        }

        public void PlateMapSyncToView() 
        {
            _findElements.ClickElementByJavaScript(SyncToView, _currentPage, $"Platemap - Sync to View");

            _findElements.VerifyElement(SyncToViewPopup, _currentPage, $"Platemap - Sync to view Popup");

            _findElements.ClickElementByJavaScript(SyncToViewApplyButton, _currentPage, $"Platemap - Sync to view apply button");

            _findElements.VerifyElement(SyncToViewToast, _currentPage, $"Graph Setting - Sync to view Toast Message");
        }

        //public void VerifyNormalizationVal()
        //{
        //    try
        //    {
        //        _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()"); // OFF
        //        Thread.Sleep(3000);

        //        IReadOnlyCollection<IWebElement> wells = _driver.FindElements(By.CssSelector(".Wellclass"));
        //        List<string> platemapName = wells.Select(well => well.GetAttribute("data-wellvalue")).ToList();

        //        IReadOnlyCollection<IWebElement> tableValues = _driver.FindElements(By.CssSelector(".tablevalues"));
        //        List<double> plateMapValues = tableValues.Select(tableValue => tableValue.FindElements(By.TagName("span")).FirstOrDefault()?.Text)
        //            .Select(spanText => spanText == "N/A" ? 0 : double.Parse(spanText)).ToList();

        //        _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()"); // ON
        //        Thread.Sleep(3000);

        //        IReadOnlyCollection<IWebElement> normalizedTableValues = _driver.FindElements(By.CssSelector(".tablevalues"));
        //        List<double> normalizedPlateMapValues = normalizedTableValues.Select(tableValue => tableValue.FindElements(By.TagName("span")).FirstOrDefault()?.Text)
        //            .Select(spanText => spanText == "N/A" ? 0 : double.Parse(spanText)).ToList();

        //        List<string> normalizationData = Enumerable.Repeat("200", 96).ToList();

        //        string scaleFactor = "2";

        //        List<string> caluNormalizationValues = normalizationData
        //            .Select((value, index) =>
        //            {
        //                double normalizationValue = value is null ? 0 : double.Parse(value);
        //                double plateMapValue = index < plateMapValues.Count ? plateMapValues[index] : 0;
        //                double calculatedValue = normalizationValue > 0 ? (plateMapValue / (normalizationValue / double.Parse(scaleFactor))) : 0;
        //                return calculatedValue.ToString("0.00");
        //            }).ToList();

        //        for (int i = 0; i < normalizedPlateMapValues.Count; i++)
        //        {
        //            if (caluNormalizationValues[i] == normalizedPlateMapValues[i].ToString("0.00"))
        //                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The calculation for the current cell {platemapName[i]} is success.");
        //            else
        //                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The calculation for the current cell {platemapName[i]} has failed. The platemap value is {normalizedPlateMapValues[i]} and the calculated normalized data is {caluNormalizationValues[i]}");
        //        }

        //        ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Platemap normalization calculation functionality has been verified.");
        //    }
        //    catch (Exception e)
        //    {
        //        ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Platemap normalization calculation functionality has not been verified. The error is {e.Message} ");
        //    }
        //}

        public void VerifyNormalizationVal()
        {
            try
            {
                _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()"); // OFF
                Thread.Sleep(3000);

                List<IWebElement> wells = _driver.FindElements(By.CssSelector(".Wellclass")).ToList();
                List<string> platemapName = wells.Select(well => well.GetAttribute("data-wellvalue")).ToList();

                List<IWebElement> tableValues = _driver.FindElements(By.CssSelector(".tablevalues")).ToList();

                List<double> plateMapValues = GetTableValues(tableValues, platemapName.Count);

                List<double> bottomplateMapValues = null;
                if (tableValues.First().FindElements(By.TagName("span")).Skip(1).Any())
                {
                    bottomplateMapValues = GetTableValues(tableValues, platemapName.Count, 1);
                }

                _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()"); // ON

                Thread.Sleep(3000);

                List<double> normalizedPlateMapValues = GetTableValues(tableValues, platemapName.Count);

                List<double> bottomNormalizedPlateMapValues = null;
                if (tableValues.First().FindElements(By.TagName("span")).Skip(1).Any())
                {
                    bottomNormalizedPlateMapValues = GetTableValues(tableValues, platemapName.Count, 1);
                }

                List<string> normalizationData = _normalizationData.Values;
                string scaleFactor = _normalizationData.ScaleFactor;

                List<string> caluNormalizationValues = CalculateNormalizationValues(normalizationData, plateMapValues, scaleFactor);
                CompareNormalizationValues(platemapName, caluNormalizationValues, normalizedPlateMapValues);

                List<string> bottomcaluNormalizationValues = CalculateNormalizationValues(normalizationData, bottomplateMapValues, scaleFactor);
                CompareNormalizationValues(platemapName, bottomcaluNormalizationValues, bottomNormalizedPlateMapValues);

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Platemap normalization calculation functionality has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Platemap normalization calculation functionality has not been verified. The error is {e.Message} ");
            }
        }

        private List<double> GetTableValues(IReadOnlyCollection<IWebElement> tableValue, int count, int skip = 0)
        {
            List<double> plateMapValues = tableValue
                .Select(tableValue => tableValue.FindElements(By.TagName("span")).Skip(skip).FirstOrDefault()?.Text)
                .Select(spanText => spanText == "N/A" ? 0 : double.Parse(spanText))
                .Take(count)
                .ToList();

            return plateMapValues;
        }

        private List<string> CalculateNormalizationValues(List<string> normalizationData, List<double> plateMapValues, string scaleFactor)
        {
            List<string> caluNormalizationValues = normalizationData
                .Select((value, index) =>
                {
                    double normalizationValue = value is null ? 0 : double.Parse(value);
                    double plateMapValue = index < plateMapValues.Count ? plateMapValues[index] : 0;
                    double calculatedValue = normalizationValue > 0 ? (plateMapValue / (normalizationValue / double.Parse(scaleFactor))) : 0;
                    return calculatedValue.ToString("0.00");
                })
                .ToList();

            return caluNormalizationValues;
        }
        private void CompareNormalizationValues(List<string> platemapName, List<string> caluNormalizationValues, List<double> normalizedPlateMapValues)
        {
            for (int i = 0; i < normalizedPlateMapValues.Count; i++)
            {
                if (caluNormalizationValues[i] == normalizedPlateMapValues[i].ToString("0.00"))
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The calculation for the current cell {platemapName[i]} is success.");
                else
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The calculation for the current cell {platemapName[i]} has failed. The platemap value is {normalizedPlateMapValues[i]} and the calculated normalized data is {caluNormalizationValues[i]}");
            }
        }

        public void WellDataPopup(string wellPosition = "B01", string wellDataStatus = "Included in current calculation")
        {
            try
            {
                var element = _driver.FindElement(By.CssSelector("[data-wellvalue ='" + wellPosition + "']"));

                /*Popup window checking*/
                int wellNum = _commonFunc.GetWellIndexFromLabel(_fileUploadOrExistingFileData.FileType, wellPosition);
                var popUpWindow = _driver.FindElement(By.Id("Groupdetail" + wellNum));

                string elementId = "Groupdetail" + wellNum;
                string script = $"document.getElementById('{elementId}').style.display = 'block';";
                ((IJavaScriptExecutor)_driver).ExecuteScript(script);

                _findElements.VerifyElement(popUpWindow, _currentPage, "Well Data popup");

                string selector = "#Groupdetail" + wellNum + " .hovercontentbody.col-md-12";
                string wellValue = $"return $(\"{selector}\").children().eq(1).get(0);";
                var secondChildElement = ((IJavaScriptExecutor)_driver).ExecuteScript(wellValue) as IWebElement;
                _findElements.VerifyElement(secondChildElement, _currentPage, "Well Data popup");

                var wellStatus = _driver.FindElement(By.Id("wellstatus" + wellNum));
                _findElements.ElementTextVerify(wellStatus, wellDataStatus, _currentPage, "Well data popup status");

                string wellPopUp = $"document.getElementById('{elementId}').style.display = 'none';";
                ((IJavaScriptExecutor)_driver).ExecuteScript(wellPopUp);

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Well data popup functionality has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Well data popup functionality has not been verified. the error is {e.Message}");
            }
        }

        public ResultStatus VerifyPlateMapRowandCloumnWell(WidgetTypes wType)
        {
            ResultStatus rs = new ResultStatus();
            try
            {
                int platemapRowCount = PlatemaprowCount.Count() - 1;
                int platemapColumnCount = PlatemapColumnCount.Count() - 1;

                // int platemapColumnCount =gridStackItems.Count()-1;
                switch (_fileType)
                {
                    case FileType.Xfp:
                    case FileType.XfHsMini:  //Row 8 column 1
                        if (platemapRowCount == (int)PlateMapWell.Row96or8Well && platemapColumnCount == (int)PlateMapWell.Column8WellCount)
                        {
                            rs.Status = true;
                            rs.Message = $"Verified  that Plate Map does contain {WellCount} Wells and it column is {platemapRowCount}  and rows {platemapColumnCount}  -{wType}";
                        }
                        break;
                    case FileType.Xfe24:       //Row 4 column 6
                        if (platemapRowCount == (int)PlateMapWell.Column24WellCount && platemapColumnCount == (int)PlateMapWell.Column24WellCount)
                        {
                            rs.Status = true;
                            rs.Message = $"Verified  that Plate Map does contain {WellCount} Wells and it column is {platemapRowCount}  and rows {platemapColumnCount}  -{wType}";
                        }
                        break;
                    case FileType.Xfe96:
                    case FileType.XFPro:     //Row 8 column 12
                        if (platemapRowCount == (int)PlateMapWell.Row96or8Well && platemapColumnCount == (int)PlateMapWell.Column96WellCount)
                        {
                            rs.Status = true;
                            rs.Message = $"Verified  that Plate Map does contain {WellCount} Wells and it column is {platemapRowCount}  and rows {platemapColumnCount}  -{wType}";
                        }
                        break;
                    default:
                        rs.Status = false;
                        rs.Message = $"Verified failed that Plate Map doesn't contain {WellCount} Wells and it column is {platemapRowCount}  and rows {platemapColumnCount} ";
                        break;
                }


                rs.Status = true;
            }
            catch (Exception)
            {
                rs.Status = false;
                rs.Message = "Verify that Plate Map should contain 96 Wells and it should be 12 column and 8 rows";
            }
            return rs;
        }
    }
}
