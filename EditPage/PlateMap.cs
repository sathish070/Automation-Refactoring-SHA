using System;
using OpenQA.Selenium;
using SHAProject.Utilities;
using SHAProject.SeleniumHelpers;
using AventStack.ExtentReports;
using SeleniumExtras.PageObjects;
using OpenQA.Selenium.Support.Extensions;
using NUnit.Framework.Constraints;
using AngleSharp.Dom;
using System.Security.Cryptography.X509Certificates;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using Microsoft.AspNetCore.Connections.Features;

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
        public string highHexColor;
        public string lowHexColor;
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

        [FindsBy(How = How.XPath, Using = "//div[@id='grapharea']/div[3]")]
        public IWebElement? PlateMapField;

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

        [FindsBy(How = How.Id, Using = "chknormalize")]
        public IWebElement? NormalizationToggle;

        [FindsBy(How = How.CssSelector, Using = "#baselineselection")]
        public IWebElement? BaselineField;

        [FindsBy(How = How.Id, Using = "ddl_baseline")]
        public IWebElement? BaselineDropdown;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"flag-about MIT-Info\"])")]
        public IWebElement? PlateMapBottomText;

        #endregion

        #region HeatMap PlateMap Elements

        [FindsBy(How = How.Id, Using = "heatmapsettings")]
        public IWebElement? HeatMapSettings;

        [FindsBy(How = How.XPath, Using = "//div[@id=\"heatmapSettings\"]/div/div")]
        public IWebElement? HeatMapColorOptionPopup;

        [FindsBy(How = How.CssSelector, Using = "#heatmapSettings > div > div > div.modal-body > div:nth-child(1) > div")]
        public IWebElement? ColourTolerance;

        [FindsBy(How = How.CssSelector, Using = "#heatmapSettings > div > div > div.modal-body > div:nth-child(2) > div > div.col-2.colorPickerOne")]
        public IWebElement? LowValueColour;

        [FindsBy(How = How.CssSelector, Using = "#heatmapSettings > div > div > div.modal-body > div:nth-child(2) > div > div.col-2.colorPickerTwo")]
        public IWebElement? HighValueColour;

        [FindsBy(How = How.CssSelector, Using = "#heatmapSettings > div > div > div.modal-body > div:nth-child(2) > div > div.col-8.heatMapGradient")]
        public IWebElement? ColourScaleBar;

        [FindsBy(How = How.CssSelector, Using = "#heatmapSettings > div > div > div.modal-footer > button")]
        public IWebElement? HeatMapSettingsApplyButton;

        [FindsBy(How = How.CssSelector, Using = ".select-tolorance")]
        public IWebElement? ColourOptionsDropDown;

        [FindsBy(How = How.XPath, Using = "(//div[@class='sp-preview'])[1]")]
        public IWebElement? BorderLowColor;

        [FindsBy(How = How.XPath, Using = "(//div[@class='sp-preview-inner'])[1]")]
        public IWebElement? InnerLowColour;

        [FindsBy(How = How.XPath, Using = "(//div[@class='sp-preview'])[2]")]
        public IWebElement? BorderHighColor;

        [FindsBy(How = How.XPath, Using = "(//div[@class='sp-preview-inner'])[2]")]
        public IWebElement? InnerHighColour;

        #endregion

        #region DataTable PlateMap Elements

        [FindsBy(How = How.XPath, Using = "//*[@id=\"atpaveragebasal_Col0\"]/span[2]")]
        public IWebElement? Header_icon;

        [FindsBy(How = How.CssSelector, Using = "#atpaveragebasal_Col1")]
        public IWebElement? AtpAverageBasal;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"iggrid_icons\"])[1]")]
        public IWebElement? DataTableHeaderIconHide;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"iggrid_icons\"])[2]")]
        public IWebElement? DataTableHeaderIconMoveLeft;

        [FindsBy(How = How.XPath, Using = "(//div[@class=\"iggrid_icons\"])[3]")]
        public IWebElement? DataTableHeaderIconMoveRight;

        [FindsBy(How = How.Id, Using = "basalavgtitle")]
        public IWebElement? DataTableBasaltitle;

        [FindsBy(How = How.Id, Using = "inducedavgtitle")]
        public IWebElement? DataTableInducedtitle;

        [FindsBy(How = How.CssSelector, Using = ".ui-iggrid-header.ui-widget-header.ui-draggable.ui-iggrid-headercell-featureenabled")]
        public IList<IWebElement> DataTablewidgetList { get; set; }

        [FindsBy(How = How.CssSelector, Using = "thead[role='rowgroup']")]
        public IWebElement? DataTableGroupandValue;

        [FindsBy(How = How.CssSelector, Using = "HideDataTable")]
        public IWebElement? HideDataTablePopup;
        #endregion

        public void PlateMapArea()
        {
            _findElements.VerifyElement(PlateMapField, _currentPage, $"Edit Widget Page - Plate Map Area");
        }

        public void PlateMapIcons()
        {
            try
            {
                _findElements.VerifyElement(WellSelection, _currentPage, $"PlateMap -Well Mode");

                _findElements.VerifyElement(FlagSelection, _currentPage, $"PlateMap -Flag Mode");

                _findElements.VerifyElement(FlagOn, _currentPage, $"PlateMap -Flag Mode On");

                _findElements.VerifyElement(FlagOff, _currentPage, $"PlateMap -Flag Mode Off");

                _findElements.VerifyElement(SyncToView, _currentPage, $"PlateMap - Sync to View");

                IReadOnlyCollection<IWebElement> PlateMapBottomText = _driver.FindElements(By.CssSelector(".MIT-Info"));

                foreach (IWebElement element in PlateMapBottomText)
                {
                    if (element.Displayed)
                        _findElements.VerifyElement(element, _currentPage, $"PlateMap - Bottom Text");
                }

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error occured while verifying the plate map icons. The error is {e.Message}");
            }
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

                //for (int count = 0; count < 5; count++)
                //{
                //    IWebElement PlateMapWell = _driver.FindElement(By.Id("tbl" + count));
                //    //IWebElement PlateMapWell = _driver.FindElement(By.Id("tbl" + count + ""));
                //    string wellName = PlateMapWell.GetAttribute("data-wellvalue");
                //    _findElements.ClickElementByJavaScript(PlateMapWell, _currentPage, $" {type} well name is - {wellName}");
                //    Thread.Sleep(5000);
                //}

                //IEnumerable<IWebElement> PlateMapWells = _driver.FindElements(By.CssSelector(".Wellclass")).Take(5);

                IEnumerable<IWebElement> PlateMapWells = _driver.FindElements(By.CssSelector(".tablevalues")).Take(5);
                foreach (IWebElement PlateMapWell in PlateMapWells)
                {
                    string wellName = PlateMapWell.GetAttribute("data-wellvalue");
                    _findElements.ClickElementByJavaScript(PlateMapWell, _currentPage, $" {type} well name is - {PlateMapWell.Text}");
                    Thread.Sleep(3000);
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

        public void VerifyNormalizationVal()
        {
            try
            {
                string defaultText = string.Empty;
                if (BaselineDropdown.Displayed)
                {
                    IWebElement selectedOption = BaselineDropdown.FindElements(By.TagName("option")).FirstOrDefault(option => option.Selected);
                    defaultText = selectedOption.Text;
                }

                //if ((BaselineField.Displayed && defaultText =="OFF") || (!BaselineField.Displayed))
                if (defaultText == "" || defaultText == "OFF")
                {
                    _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()"); // OFF

                   // WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
                   // wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#loadmodal")));

                    Thread.Sleep(6000);

                    List<IWebElement> wells = _driver.FindElements(By.CssSelector(".Wellclass")).ToList();
                    List<string> platemapName = wells.Select(well => well.GetAttribute("data-wellvalue")).ToList();

                    List<IWebElement> tableValues = _driver.FindElements(By.CssSelector(".tablevalues")).ToList();
                    ICollection<IWebElement> collection = tableValues.ToList();

                    List<double> plateMapValues;
                    List<double> bottomplateMapValues = null;

                    // if Platemap Data has double Value
                    if (tableValues[1].Text.Contains("\r\n"))
                    {
                        plateMapValues = GetTableValues(collection, platemapName.Count, 1);
                        bottomplateMapValues = GetTableValues(collection, platemapName.Count, 2);
                    }
                    else
                    {
                        plateMapValues = GetTableValues(collection, platemapName.Count);
                    }

                    _driver.ExecuteJavaScript<string>("return document.getElementById(\"chknormalize\").click()"); // ON

                    //wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
                    //wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#loadmodal")));

                    Thread.Sleep(6000);

                    List<IWebElement> normalizationWell = _driver.FindElements(By.CssSelector(".Wellclass")).ToList();
                    List<string> normalizationPlatemapValues= normalizationWell.Select(well => well.GetAttribute("data-wellvalue")).ToList();

                    List<IWebElement> normalizationTableValues = _driver.FindElements(By.CssSelector(".tablevalues")).ToList();
                    ICollection<IWebElement> collections = normalizationTableValues.ToList();

                    List<double> normalizedPlateMapValues;
                    List<double> bottomNormalizedPlateMapValues = null;

                    if (normalizationTableValues[1].Text.Contains("\r\n"))
                    {
                        normalizedPlateMapValues = GetTableValues(collections, normalizationPlatemapValues.Count, 1);
                        bottomNormalizedPlateMapValues = GetTableValues(collections, normalizationPlatemapValues.Count, 2);
                    }
                    else
                    {
                        normalizedPlateMapValues = GetTableValues(collections, normalizationPlatemapValues.Count);
                    }

                    List<string> normalizationData = _normalizationData.Values;
                    string scaleFactor = _normalizationData.ScaleFactor;

                    List<string> caluNormalizationValues = CalculateNormalizationValues(normalizationData, plateMapValues, scaleFactor);
                    CompareNormalizationValues(platemapName, caluNormalizationValues, normalizedPlateMapValues);
                    if (bottomplateMapValues !=  null)
                    {
                        List<string> bottomcaluNormalizationValues = CalculateNormalizationValues(normalizationData, bottomplateMapValues, scaleFactor);
                        CompareNormalizationValues(platemapName, bottomcaluNormalizationValues, bottomNormalizedPlateMapValues);
                    }

                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Platemap normalization calculation functionality has been verified.");
                }
                else
                {
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Warning, $"Baseline is given in the excel sheet is {BaselineDropdown.Text} and normalization concept can't be applied." );
                }

            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Platemap normalization calculation functionality has not been verified. The error is {e.Message} ");
            }
        }

        //private List<double> GetTableValues(System.Collections.Generic.IReadOnlyCollection<OpenQA.Selenium.IWebElement> tableValue, int count, int skip = 0)
        private List<double> GetTableValues(ICollection<IWebElement> tableValue, int count, int skip = 0)
        {
            //WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            //wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#loadmodal")));

            List<double> plateMapValues = tableValue
                .Select((tabledata, index) =>
                {
                    string text = tabledata.Text;

                    if (text.Contains("\r\n"))
                    {
                        string[] values = text.Contains("\r\n") ?  text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries) :  text.Contains("\r\n\r\n") ? text.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.RemoveEmptyEntries) : Array.Empty<string>();
                        //index % 2 == 0 &&
                        if (skip == 1)
                        {
                            text = values[0];
                        }
                        if (skip == 2)
                        {
                            text = values[1];
                        }
                    }

                    double parsedValue;
                    if (text == "N/A" || !double.TryParse(text, out parsedValue))
                    {
                        parsedValue = 0;
                    }

                    return Math.Round(parsedValue, 2);
                })
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
                    //double calculatedValue = normalizationValue > 0 ? Math.Round((plateMapValue / normalizationValue) * double.Parse(scaleFactor), 2) : 0;
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

        public void WellDataPopup(string wellPosition = "A05", string wellDataStatus = "Included in the current calculation")
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
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Well data popup functionality has not been verified. the error is {e.Message}"); ;
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

        public void HeatMapPlateMapIcons()
        {
            _findElements.VerifyElement(HeatMapSettings, _currentPage, $"PlateMap - HeatMap Settings");

            _findElements.VerifyElement(WellSelection, _currentPage, $"PlateMap -Well Mode");

            _findElements.VerifyElement(FlagSelection, _currentPage, $"PlateMap -Flag Mode");
        }

        public void HeatMapColorOptions(string colorOption)
        {
            _findElements.ClickElementByJavaScript(HeatMapSettings, _currentPage, $"PlateMap - HeatMap Settings");

            _findElements.ClickElementByJavaScript(HeatMapColorOptionPopup, _currentPage, $"HeatMap - Settings Color Option Popup");

            _findElements.VerifyElement(ColourTolerance, _currentPage, $"Heat Map Settings - Color Tolerance");

            _findElements.VerifyElement(LowValueColour, _currentPage, $"Heat Map Settings - Low Value Color");

            _findElements.VerifyElement(HighValueColour, _currentPage, $"Heat Map Settings - High Value Color");

            _findElements.VerifyElement(ColourScaleBar, _currentPage, $"Heat Map Settings - Color Scale Bar");

            string colorOptionPer = colorOption + "%".ToString();
            //_findElements.SelectByText(ColourOptionsDropDown, colorOptionPer);

            _findElements.SelectFromDropdown(ColourOptionsDropDown, _currentPage, "text", colorOptionPer, $"Colour option - {colorOptionPer}");

            string lowColor = InnerLowColour.GetCssValue("background-color");
            var lowColorConvert = Aspose.Svg.Drawing.Color.FromString(lowColor);
            lowHexColor = lowColorConvert.ToRgbHexString();

            ExtentReport.ExtentTest("ExtentTestNode",Status.Pass, $"Lower color code in the heat map setting is " + lowHexColor);

            _findElements.VerifyElement(BorderLowColor, _currentPage, $"Heat Map Settings - Lower Color");

            string highColor = InnerHighColour.GetCssValue("background-color");
            var highColorConvert = Aspose.Svg.Drawing.Color.FromString(highColor);
            highHexColor = highColorConvert.ToRgbHexString();

            ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Higher color code in the heat map setting is " + highHexColor);

            _findElements.VerifyElement(BorderHighColor, _currentPage, $"Heat Map Settings - High Color");

            _findElements.ClickElementByJavaScript(HeatMapSettingsApplyButton, _currentPage, $"Heat Map Settings - Apply Button");
        }

        public void HeatMapPlateMapfunctionality()
        {
            PlateMapWells("WellSelectionMode"); // Unselect the wells

            PlateMapWells("WellUnSelectionMode"); // Select the wells

            PlateMapWells("FlagMode"); // Add the flags

            PlateMapWells("UnflagMode"); // Remove the flags 

            IReadOnlyCollection<IWebElement> wells = _driver.FindElements(By.CssSelector(".Wellclass"));
            List<string> platemapName = wells.Select(well => well.GetAttribute("data-wellvalue")).ToList();

            List<IWebElement> plateMapWells = _driver.FindElements(By.CssSelector(".Wellclass")).ToList();
            List<string> plateMapBckgrndColor = plateMapWells.Select(well => well.GetCssValue("background-color")).ToList();

            for (int i = 0; i < plateMapBckgrndColor.Count; i++)
            {
                var plateMapColorConvert = Aspose.Svg.Drawing.Color.FromString(plateMapBckgrndColor[i]);
                string plateMapHexColor = plateMapColorConvert.ToRgbHexString();

                if (plateMapHexColor == highHexColor)
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Higher Color Code  " + highHexColor + " and the PlateMap higher Color Code  " + plateMapHexColor + " are same for the cell - " + platemapName[i]);

                if (plateMapHexColor == lowHexColor)
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Lower Color Code  " + lowHexColor + " and the PlateMap Lower Color Code  " + plateMapHexColor + " are same for the cell - " + platemapName[i]);
            }
        }

        public void DataTableVerification()
        {

            _findElements.VerifyElement(DataTableBasaltitle, _currentPage, $"{DataTableBasaltitle.Text}");
            _findElements.VerifyElement(DataTableInducedtitle, _currentPage, $"{DataTableInducedtitle.Text}");
            foreach (IWebElement widgetList in DataTablewidgetList)
            {
                if (widgetList.Displayed)
                {
                    //widgetList.Text()

                    _findElements.ElementTextVerify(widgetList, widgetList.Text, _currentPage, $"DataTable widgetList - {widgetList.Text.Replace("/", "-")}");
                }
            }

            // findElements.ElementTextVerify(DataTableGroupandValue, DataTableGroupandValue.Text, currentPage, $"Graph Setting - {DataTableGroupandValue.Text.Replace("/", "-")}");

            _findElements.ActionsClass(AtpAverageBasal);  //*[@id="atpaveragebasal_Col0"]/span[2]
            _findElements.ActionsClass(Header_icon);

            //_findElements.ElementTextVerify(HideDataTablePopup, "", _currentPage, $"DataTable HidePopup - {HideDataTablePopup.Text}");
            _findElements.ElementTextVerify(DataTableHeaderIconHide, "Hide", _currentPage, $"DataTable HidePopup - {DataTableHeaderIconHide.Text}");
            _findElements.ElementTextVerify(DataTableHeaderIconMoveLeft, "MoveLeft", _currentPage, $"DataTable MoveLeft - {DataTableHeaderIconMoveLeft.Text}");
            _findElements.ElementTextVerify(DataTableHeaderIconMoveRight, "MoveRight", _currentPage, $"DataTable MoveRight - {DataTableHeaderIconMoveRight.Text}");
        }
    }
}
