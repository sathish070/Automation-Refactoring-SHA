using Microsoft.Extensions.Options;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V112.WebAuthn;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Safari;
using SeleniumExtras.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHAProject.Utilities
{
    public class DriverSetup
    {
        public IWebDriver? driver;
        public string loginFolderPath;

        public IWebDriver browser(string currentBroswer, string website, string downloadPath)
        {
            getDriver( currentBroswer, downloadPath);
            driver.Manage().Cookies.DeleteAllCookies();
            driver.Navigate().GoToUrl(website);
            driver.Manage().Window.Maximize();
            return driver;
        }

        public IWebDriver getDriver(string currentBrowser, string downloadPath)
        {
            switch (currentBrowser)
            {
                case "Chrome":
                    ChromeOptions chromeOptions = new();
                    chromeOptions.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                    chromeOptions.AddUserProfilePreference("download.default_directory", downloadPath);
                    chromeOptions.PageLoadStrategy = PageLoadStrategy.Eager;
                    driver = new ChromeDriver(chromeOptions);
                    break;
                case "Edge":
                    EdgeOptions edgeOptions = new();
                    edgeOptions.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                    edgeOptions.AddUserProfilePreference("download.default_directory", downloadPath);
                    edgeOptions.PageLoadStrategy = PageLoadStrategy.Eager;
                    //edgeOptions.AddArguments("--headless");
                    edgeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                    edgeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                    driver = new EdgeDriver(edgeOptions);
                    break;
                case "Firefox":
                    FirefoxOptions firefoxOptions = new();
                    firefoxOptions.SetPreference("browser.download.dir", downloadPath);
                    firefoxOptions.SetPreference("browser.download.folderList", 2);
                    firefoxOptions.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream");
                    firefoxOptions.SetPreference("browser.download.manager.showWhenStarting", false);
                    firefoxOptions.PageLoadStrategy = PageLoadStrategy.Eager;
                    //firefoxOptions.AddArguments("--headless");
                    driver = new FirefoxDriver(firefoxOptions);
                    break;
                case "Safari":
                    SafariOptions safariOptions = new SafariOptions();
                    Dictionary<string, object> addSafariOptions = new Dictionary<string, object>();
                    addSafariOptions.Add("safari:automaticDownloads", false);
                    addSafariOptions.Add("safari:automaticDownloadPolicy", 2);
                    addSafariOptions.Add("webdriver.safari.nosleep", true);
                    addSafariOptions.Add("browser.download.folderList", 2);
                    //addSafariOptions.Add("safari:downloadDir", DownloadPath);
                    safariOptions.AddAdditionalOption("safari.options", addSafariOptions);
                    //safariOptions.AddAdditionalOption("download.default_directory", DownloadPath);
                    driver = new SafariDriver(safariOptions);
                    //Thread.sleep(1000);
                    break;
                default:
                    /*This block of code will be executed if the browserName argument does not match any of the cases above*/
                    ChromeOptions options = new();
                    options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                    options.AddUserProfilePreference("download.default_directory", downloadPath);
                    options.PageLoadStrategy = PageLoadStrategy.Eager;
                    driver = new ChromeDriver(options);
                    break;
            }
            return driver;
        }
    }
}
