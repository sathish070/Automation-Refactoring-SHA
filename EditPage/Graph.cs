using AngleSharp.Io;
using Aspose.Svg.Drawing;
using Aspose.Svg.Net;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Gherkin.Model;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using SeleniumExtras.PageObjects;
using SHAProject.SeleniumHelpers;
using SHAProject.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.EditPage 
{
    public class Graph
    {
        public IWebDriver? _driver;
        public FindElements? _findElements;
        public CommonFunctions? _commonFunc;
        public string _currentPage = string.Empty;
        public FileUploadOrExistingFileData _fileUploadOrExistingFileData;
        private double maxValue;
        private double minValue;
        public IWebElement? AmchartChart;
        private readonly List<double> doubles;

        public Graph(string currentPage, IWebDriver driver, FindElements findElements, CommonFunctions commonFunc)
        {
            _driver = driver;
            _commonFunc = commonFunc;
            _currentPage = currentPage;
            _findElements = findElements;
            PageFactory.InitElements(_driver, this);
        }

        #region PanZoom
        [FindsBy(How = How.XPath, Using = "(//canvas[@class='canvasjs-chart-canvas'])[2]")]
        public IWebElement? CanvasChart;

        //[FindsBy(How = How.XPath, Using = "//button[@title='Switch to Pan']")]
        //public IWebElement? PanIcon;

        [FindsBy(How = How.XPath, Using = "//button[@class=\"zoom-btn\"]")]
        public IWebElement? PanIcon;

        [FindsBy(How = How.XPath, Using = "//button[@class=\"reset-btn\"]")]
        public IWebElement? ZoomIcon;

        //[FindsBy(How = How.XPath, Using = "//button[@title='Switch to Zoom']")]
        //public IWebElement? ZoomIcon;

        [FindsBy(How = How.XPath, Using = "//button[@title='Reset']")]
        public IWebElement? ResetIcon;

        //[FindsBy(How = How.Id, Using = "divwidget1")]
        //public IWebElement? AMchartChart;

        [FindsBy(How = How.XPath, Using = "//div[@class=\"barchart-area ui-resizable barwidget\"]")]
        public IWebElement? AMchartChart;

        #endregion


        [FindsBy(How = How.XPath, Using = "//div[@id='grapharea']/div[1]")]
        public IWebElement? GraphAreaField;

        public void GraphArea()
        {
            _findElements.VerifyElement(GraphAreaField, _currentPage, $"Edit Widget Page -Graph Area");
        }

        public void PanZoom(ChartType Chart)
        {
            IWebElement? element = ChartType.CanvasJS == Chart ? CanvasChart : AMchartChart;

            _findElements.VerifyElement(element, _currentPage, $"Canvas Chart");

            Actions actions = new Actions(_driver);
            actions.MoveToElement(element)
                  .ClickAndHold()
                  .Build()
                  .Perform();

            actions.MoveByOffset(300, 150)
                  .Release()
                  .Build()
                  .Perform();

            _findElements.ClickElementByJavaScript(PanIcon, _currentPage, $"Pan Icon");

            actions.MoveToElement(element)
                  .ClickAndHold()
                  .Build()
                  .Perform();

            actions.MoveByOffset(300, 150)
                  .Release()
                  .Build()
                  .Perform();

            _findElements.ClickElementByJavaScript(ResetIcon, _currentPage, $"Reset Icon");

            _findElements.VerifyElement(element, _currentPage, $"Canvas Chart");

        }

        public void VerifyExpectedGraphUnits(string ExpectedGraphUnits, WidgetTypes widgetType)
        {
            try
            {
                Thread.Sleep(2000);

                ChartType chartType = _commonFunc.GetChartType(widgetType);
                string graphUnits = _commonFunc.GetGraphUnits(chartType);

                if (graphUnits == ExpectedGraphUnits)
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"The graph units -{graphUnits} and exact graph units -{ExpectedGraphUnits} are equal.");
                else
                    ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"The graph units -{graphUnits} and exact graph units -{ExpectedGraphUnits} are not equal.");

                ExtentReport.ExtentTest("ExtentTestNode", Status.Pass, $"Verify Normalization units has been verified.");
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Verify Normalization units has not been verified. The error is {e.Message}");
            }
        }

        public void AmChartToolTip()
        {
            try
            {
                string path = "(//*[@r='2' or @r='3' or @r='5' or @r='6'])";
                IList<IWebElement> toolTips = _driver.FindElements(By.XPath(path));

                foreach (IWebElement toolTip in toolTips.Take(5))
                {
                    _findElements.ActionsClass(toolTip);
                    Thread.Sleep(1000);
                    string tooltipId = toolTip.GetAttribute("id");
                    _findElements.VerifyElement(toolTip, _currentPage, $"Tooltip graph {tooltipId}");
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error occurred while verifiying the graph tooltip.The error is {e.Message}");
            }
        }

        public void BarGraphVerification()
        {
            try
            {
                string path = "(//*[local-name()='svg'])[1]//*[name()='g' and @role='list']//*[contains(@role, 'listitem')]";

                IList<IWebElement> elements = _driver.FindElements(By.XPath(path));

                foreach (IWebElement element in elements.Take(5))
                {
                    _findElements.ActionsClass(element);
                    Thread.Sleep(1000);

                    _findElements.VerifyElement(element, _currentPage, "Tootip graph");
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error occurred while verifiying the graph tooltip.The error is {e.Message}");
            }
        }

        public void CanvasChartTooltip()
        {
            Thread.Sleep(5000);

            Actions actions = new Actions(_driver);

            actions.MoveToElement(CanvasChart).MoveByOffset(0, 0).Build().Perform();

            IWebElement tooltip = _driver.FindElement(By.XPath("//div[@class=\"canvasjs-chart-tooltip\"]/div"));

            string tooltipValue = tooltip.Text;

            ScreenShot.ScreenshotNow(_driver, _currentPage, $"Canvas Chart tooltip value - {tooltipValue}", ScreenshotType.Error, tooltip);
        }

        public (double maxValue, double minValue, List<double> doubles) GraphYmaxYminVerification()
        {
            double maxValue = 0.0;
            double minValue = 0.0;

            try
            {
                IList<IWebElement> nextSiblings = GetNextSiblings("[transform='translate(16,0)'] g");
                List<double> doubles = GetTspanTextValues(nextSiblings);

                maxValue = doubles.Max();
                minValue = doubles.Min();

                foreach (var item in nextSiblings)
                {
                    Thread.Sleep(1000);
                    if (item.Displayed)
                    {
                        if (!string.IsNullOrEmpty(item.Text))
                        {
                            if (double.TryParse(item.Text, out double itemValue))
                            {
                                if (itemValue == maxValue)
                                {
                                    _findElements.ScrollIntoView(item);
                                    _findElements.VerifyElement(item, _currentPage, $"Graph Y-Axis Maximum value: {maxValue}");
                                }

                                if (itemValue == minValue)
                                {
                                    _findElements.ScrollIntoView(item);
                                    _findElements.VerifyElement(item, _currentPage, "Graph Y-Axis Minimum value: {minValue}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTestNode", Status.Fail, $"Error Occurred while verifying the graph maximum and minimum value. The error is {e.Message}");
            }
            return (maxValue, minValue, doubles);
        }

        private IList<IWebElement> GetNextSiblings(string selector)
        {
            return _driver.FindElements(By.CssSelector($"{selector} + *"));
        }

        private List<double> GetTspanTextValues(IList<IWebElement> elements)
        {
            List<double> values = new List<double>();

            foreach (var element in elements)
            {
                double value;
                if (double.TryParse(element.Text, out value))
                {
                    values.Add(value);
                }
            }
            return values;
        }
    }
}
