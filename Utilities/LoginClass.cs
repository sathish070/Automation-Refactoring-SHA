using AngleSharp.Dom;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Model;
using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using NUnit.Framework.Internal;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;
using SHAProject.PageObject;
using SHAProject.SeleniumHelpers;
using SHAProject.Workflows;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WebDriverManager;
using static System.Net.WebRequestMethods;

namespace SHAProject.Utilities
{
    public class LoginClass : Tests
    {
        public readonly IWebDriver _driver;
        public LoginData? loginData;
        public UploadFile? uploadFile;
        public FindElements? findElements;
        private static IWebDriver webdriver;

        public LoginClass(IWebDriver driver,LoginData loginData,CommonFunctions commonFunctions) 
        {
            _driver = driver;
            this.loginData = loginData;
            this.commonFunc = commonFunctions;
            PageFactory.InitElements(_driver, this);
        }

        [FindsBy(How = How.Id, Using = "username")]
        public IWebElement emailtext;

        [FindsBy(How = How.Id, Using = "password")]
        public IWebElement passwordtext ;

        [FindsBy(How = How.CssSelector, Using = ".agt-btn")]
        private IWebElement signIn;

        public bool LoginAsExcelUser()
        {
            try
            {
                string folderPath = Tests.loginFolderPath;
                ExtentReport.CreateExtentTest("LoginPage");
                commonFunc.CreateDirectory(folderPath, "Login");
                string loginFoldersPath = folderPath + "\\" + "Login";
                commonFunc.CreateDirectory(loginFoldersPath, "Success");
                commonFunc.CreateDirectory(loginFoldersPath, "Error");

                findElements = new FindElements(_driver);

                findElements.VerifyElement(emailtext, "Login",$"Login Page -Email Field");

                findElements.SendKeys(loginData.UserName, emailtext, "Login", $"Given mail id is { loginData.UserName}");

                findElements.VerifyElement(passwordtext, "Login", "Login Page- Password Field");

                findElements.SendKeys(loginData.Password, passwordtext, "Login", $"Given Password is { loginData.Password}");

                findElements.ClickElement(signIn,"Login","Login Page - Sign in Button");

                ExtentReport.ExtentTest("ExtentTest", Status.Pass, $"Login successfully with the given credentials");
                return true;
            }
            catch (Exception e)
            {
                ExtentReport.ExtentTest("ExtentTest", Status.Fail, $"Not login successfully with the given credentials {e.Message}");
                return false;   
            }
        }
    }
}
