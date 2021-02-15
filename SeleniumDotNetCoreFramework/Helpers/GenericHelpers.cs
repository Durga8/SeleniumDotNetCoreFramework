using OpenQA.Selenium;
using SeleniumDotNetCoreFramework.Base;
using System;
using System.Collections.Generic;
using System.Threading;
using NUnit.Framework;
using AventStack.ExtentReports;
using System.Globalization;
using OpenQA.Selenium.Interactions;
using System.Data;

namespace SeleniumDotNetCoreFramework.Helpers
{
    public class GenericHelpers:HookBase
    {
        static IJavaScriptExecutor jse;
        public static void Verification(string expected, string actual, string message)
        {
            try
            {
                if (expected == actual)
                {
                    Console.WriteLine("verification completed Sucessfully");
                    reportLog(Status.Pass, message + " verification completed Sucessfully . <br />  Expected value -  " + expected + " <br /> Actual Result -  " + actual);
                }
                else
                {
                    Console.WriteLine("verification Failed");
                    reportLog(Status.Fail, message + " verification failed . <br /> Expected value -  " + expected + " <br /> Actual Result -  " + actual);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
            }

        }

        public static void AssertValues(string expected, string actual, string message)
        {
            Thread.Sleep(8000);
            if (expected.Equals(actual))
            {
                reportLog(Status.Pass, message + " verification completed Sucessfully . <br /> Expected value -  " + expected + " <br /> Actual Result -  " + actual);
            }
            else
            {
                reportLog(Status.Fail, message + " verification failed . <br /> Expected value -  " + expected + " <br /> Actual Result -  " + actual);
                Assert.AreEqual(actual, expected);

            }
        }

        public static void AssertValues(Boolean expected, Boolean actual, string message)
        {
            if (expected == actual)
            {
                reportLog(Status.Pass, message + " verification completed Sucessfully . <br /> Expected value -  " + expected + " <br /> Actual Result -  " + actual);
            }
            else
            {
                reportLog(Status.Fail, message + " verification failed . <br /> Expected value -  " + expected + " <br /> Actual Result -  " + actual);
                Assert.AreEqual(actual, expected);
            }
        }

        public static void AutofillValue(By field, By elementList, string autoFillValue, string type)
        {
            Thread.Sleep(5000);
            IList<IWebElement> matchingList;
            Boolean found = false;
            if (type == "text")
            {
                Thread.Sleep(5000);
                driver.FindElement(field).SendKeys(autoFillValue);

                Thread.Sleep(5000);

                matchingList = driver.FindElement(elementList).FindElements(By.TagName("li"));

                foreach (IWebElement item in matchingList)
                {
                    if (item.Text.Contains(autoFillValue))
                    {
                        Thread.Sleep(1000);
                        item.Click();
                        found = true;
                        reportLog(Status.Info, item.Text + "item found");
                        if (found == true)
                        {
                            reportLog(Status.Info, autoFillValue + "is found in one of list item texts" + item.Text);
                        }
                        else
                        {
                            reportLog(Status.Fail, "On " + autoFillValue + " Could not find the " + item.Text);
                        }
                        break;
                    }
                }
            }
            else if (type == "dropdown")
            {
                if (!driver.FindElement(field).Enabled)
                {
                    reportLog(Status.Info, "The dropdown is not enabled...");
                }
                else
                {
                    //driver.FindElement(field).Click();
                    IWebElement dropdownButton = driver.FindElement(field); // locate the button, can be done with any other selector
                    Actions action = new Actions(driver);
                    action.MoveToElement(dropdownButton).Perform(); // move to the button
                    dropdownButton.Click();
                    reportLog(Status.Info, "Clicked on the dropdown..");
                }
                Thread.Sleep(1000);

                matchingList = driver.FindElement(elementList).FindElements(By.TagName("li"));

                foreach (IWebElement item in matchingList)
                {
                    if (item.Text.Equals(autoFillValue))
                    {
                        item.Click();
                        found = true;
                        break;
                    }
                }
            }

        }



        public static void JavaScriptScrollAndClick(By webElement)
        {
            jse = (IJavaScriptExecutor)driver;
            jse.ExecuteScript("arguments[0].scrollIntoView()", driver.FindElement(webElement));
            jse.ExecuteScript("arguments[0].click()", driver.FindElement(webElement));
        }

        public static string getUSTime()
        {
            return TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("US Mountain Standard Time")).ToString("MMM dd yyyy");
        }

        public static string DecimalFormatter(string value)
        {
            var decimalvalue = decimal.TryParse(value, out decimal parser) ? parser : parser;
            string convertedvalue = decimalvalue.ToString("N2", CultureInfo.InvariantCulture);
            return convertedvalue;
        }

        // added by megha - create common xpath
        public static IWebElement prepareWebElementtWithDynamicXpath(string xpathValue, string substitutionValue)
        {
            string xpathAfterReplace = xpathValue.Replace("firstDynamicValue", substitutionValue);
            reportLog(Status.Info, "dynamic xpath created is = " + xpathAfterReplace);
            return driver.FindElement(By.XPath(xpathAfterReplace));
        }


        //Added by megha - common xpath for list of elemments
        public static IList<IWebElement> prepareWebElementsXpathWithDynamicData(string xpathValue, string substitutionValue)
        {
            string xpathAfterReplace = xpathValue.Replace("firstDynamicValue", substitutionValue);
            reportLog(Status.Info, "dynamic xpath created is = " + xpathAfterReplace);
            return driver.FindElements(By.XPath(xpathAfterReplace));
        }
    }
}

