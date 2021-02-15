using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Safari;
using SeleniumDotNetCoreFramework.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static SeleniumDotNetCoreFramework.Base.Browser;

namespace SeleniumDotNetCoreFramework.Base
{
    public class DriverSetUp
    {
        private static DriverSetUp driverSetUpInstance = null;

        private static IWebDriver driver;


        public DriverSetUp()
        {

        }
        public static DriverSetUp getInstance()
        {

            if (driverSetUpInstance == null)
            {

                driverSetUpInstance = new DriverSetUp();


            }
            return driverSetUpInstance;

        }



        public IWebDriver getWebDriver()
        {

            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--disable-notifications");
            //Headless ChromeBrowser
            //options.AddArgument("--headless");
            InternetExplorerOptions caps = new InternetExplorerOptions();
            caps.IgnoreZoomLevel = true;
            caps.EnableNativeEvents = false;
            caps.InitialBrowserUrl = "http://localhost";
            caps.UnhandledPromptBehavior = UnhandledPromptBehavior.Accept;
            caps.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            caps.EnablePersistentHover = true;


            string browserType = ExcelHelpers.getParameter("Browser");

            switch (browserType)
            {
                case "Chrome":
                    //chromeDriverDirectory, options,TimeSpan.FromMinutes(5)
                    driver = new ChromeDriver(options);


                    break;
                case "IE":
                    //set capability 

                    driver = new InternetExplorerDriver(caps);
                    break;

                case "Safari":
                    driver = new SafariDriver();
                    break;

                default:
                    //no browser found
                    break;

            }

            return driver;


        }
        public IWebDriver OpenBrowser(BrowserType browserType)
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--disable-notifications");
            //Headless ChromeBrowser
            //options.AddArgument("--headless");
            InternetExplorerOptions caps = new InternetExplorerOptions();
            caps.IgnoreZoomLevel = true;
            caps.EnableNativeEvents = false;
            caps.InitialBrowserUrl = "http://localhost";
            caps.UnhandledPromptBehavior = UnhandledPromptBehavior.Accept;
            caps.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            caps.EnablePersistentHover = true;
            switch (browserType)
            {
                case BrowserType.InternetExplorer:
                    driver = new InternetExplorerDriver(caps);
                    break;
                case BrowserType.Chrome:
                    driver = new ChromeDriver(options);
                    options.AddArguments("--disable-backgrounding-occluded-windows");
                    break;
                default:
                    driver = new ChromeDriver(options);
                    options.AddArguments("--disable-backgrounding-occluded-windows");
                    break;
            }
            return driver;

        }
    }
}
