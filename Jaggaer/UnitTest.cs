using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Web;
using Jaggaer.TSM;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumHelpers;
using OpenQA.Selenium.Support.Extensions;
using IronXL;
using IronXL.Xml.Spreadsheet;
using System.Linq;
using System.Data;
using System.Runtime.DesignerServices;
using System.Threading;
using System.Net.NetworkInformation;
using System.Web.UI.WebControls;

namespace SeleniumHelpers
{
    public static class Extensions
    {
        public static string GetStringFromWebElement(this IWebDriver d, string xPath)
        {
            string value = "";
            var els = d.FindElements(By.XPath(xPath));
            if (els.Count > 0)
                value = els[0].Text;
            return value;
        }
        public static int GetNumberFromWebElement(this IWebDriver d, string xPath)
        {
            int value = 0;
            IWebElement el1 = d.FindElement(By.XPath(xPath));

            string ele = (el1.TagName.ToUpper() != "SELECT") ? d.GetStringFromWebElement(xPath) : el1.GetAttribute("value");
            if (ele != "")
                value = Convert.ToInt32(Regex.Replace(ele, "[^0-9]+", string.Empty));
            return value;
        }
        public static IWebElement ClickWebElement(this IWebDriver d, string xPath, int delay = 5)
        {
            IWebElement el = null;
            Actions actions = new Actions(d);
            IJavaScriptExecutor JS = (IJavaScriptExecutor)d;
            try
            {
                el = new WebDriverWait(d, TimeSpan.FromSeconds(delay)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(xPath)));
                el = d.FindElement(By.XPath(xPath));
            }
            catch (Exception e)
            {
                if (e.Message != "")
                    el = null;
            }
            if (el != null)
            {
                JS.ExecuteScript("arguments[0].scrollIntoView(true);", el);
                JS.ExecuteScript("arguments[0].click();", el);
            }
            return el;
        }
        public static IWebElement FillWebElement(this IWebDriver d, string xPath, string text, int delay = 5)
        {
            IWebElement el = null;
            try
            {
                el = new WebDriverWait(d, TimeSpan.FromSeconds(delay)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath(xPath)));
                el = d.FindElement(By.XPath(xPath));
            }
            catch (Exception e)
            {
                if (e.Message != "")
                    el = null;
            }
            if (el != null)
                el.SendKeys(text);
            return el;
        }
        // Uses JavaScript to click a checkbox - avoiding interaction errors
        public static void ToggleCheckBox(this IWebDriver d, string xPath)
        {
                var ele =d.FindElement(By.XPath(xPath));
                d.ExecuteJavaScript("arguments[0].click();", ele);
        }
        public static void ClickButton(this IWebDriver d, string xPath)
        {
            ToggleCheckBox(d, xPath);
        }
        public static void FillField(this IWebDriver d, string xPath, string value)
        {
            var ele = d.FindElement(By.XPath(xPath));
            d.ExecuteJavaScript("arguments[0].value=arguments[1];", ele, value);
        }
        public static void SelectOption(this IWebDriver d, string xPath)
        {
            var ele = d.FindElement(By.XPath(xPath));
            d.ExecuteJavaScript("arguments[0].selected='selected';", ele);
        }
        public static void ClickLink(this IWebDriver d, string XPath)
        {
            ToggleCheckBox(d, XPath);
        }
        public static string GetFormValue(this IWebDriver d, string xPath)
        {
            return d.FindElement(By.XPath(xPath)).GetAttribute("value");          
        }
        public static IWebElement GetWebElement(this IWebDriver d, string xPath, int delay = 5)
        {
            IWebElement el = null;
            try
            {
                el = new WebDriverWait(d, TimeSpan.FromSeconds(delay)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath(xPath)));
                el = d.FindElement(By.XPath(xPath));
            }
            catch (Exception e)
            {
                if (e.Message != "")
                    el = null;
            }
            return el;
        }
    }
}
namespace Jaggaer
{
    public class HomePage
    {
        public enum ActionItemType
        {
            CartsAssignedToMe,AssignedPurchaseOrders,AssignedInvoices,AssignedSupplierRegistrations,AssignedSupplierRequestApprovals,
            UnassignedPurchaseOrders,UnassignedInvoices,UnassignedSupplierRegistrations,UnassignedSupplierRequests,
            PriceFileToReview,ImportExportCompleted,SupplierDataReview
        }
        public IWebDriver driver { get; set; }
        public HomePage(IWebDriver d)
        {
            driver = d;
        }
        public void SelectMenu(params string[] names)
        {
            if (names[0]=="Home"){
                driver.ClickWebElement("//a[@aria-label='Home']");
            }
            else
            {
                driver.ClickWebElement("//a[@aria-label='" + names[0] + "']");

                driver.ClickWebElement("//a[@aria-label='" + names[1] + "']");

                driver.ClickWebElement("//a[@title='" + names[2] + "']");
            }
        }
        public void GoAction(ActionItemType t)
        {
            ////h5[.='My Assigned Approvals']/../ul/li/span/a[.='Supplier Registrations']
        }
        public ActionItem GetAction(ActionItemType type)
        {
            ActionItem result = new ActionItem("", driver);
            switch (type)
            {
                case ActionItemType.CartsAssignedToMe:
                    result = new ActionItem("Carts Assigned To Me",driver);
                    result.Parent = "My Assigned Approvals";
                    break;
                case ActionItemType.AssignedPurchaseOrders:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "My Assigned Approvals";
                    break;
                case ActionItemType.AssignedInvoices:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "My Assigned Approvals";
                    break;
                case ActionItemType.AssignedSupplierRegistrations:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "My Assigned Approvals";
                    break;
                case ActionItemType.AssignedSupplierRequestApprovals:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "My Assigned Approvals";
                    break;
                case ActionItemType.UnassignedPurchaseOrders:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Unassigned Approvals";
                    break;
                case ActionItemType.UnassignedInvoices:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Unassigned Approvals";
                    break;
                case ActionItemType.UnassignedSupplierRegistrations:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Unassigned Approvals";
                    break;
                case ActionItemType.UnassignedSupplierRequests:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Unassigned Approvals";
                    break;
                case ActionItemType.PriceFileToReview:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Administrative Items";
                    break;
                case ActionItemType.ImportExportCompleted:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Administrative Items";
                    break;
                case ActionItemType.SupplierDataReview:
                    result = new ActionItem("Carts Assigned To Me", driver);
                    result.Parent = "Administrative Items";
                    break;
                default:                
                    return result;
            }
            return result;
        }
        public void SelectActionItem(ActionItemType t)
        {
            switch (t)
            {
                case ActionItemType.CartsAssignedToMe:
                    break;
                case ActionItemType.AssignedPurchaseOrders:
                    break;
                case ActionItemType.AssignedInvoices:
                    break;
                case ActionItemType.AssignedSupplierRegistrations:
                    break;
                case ActionItemType.AssignedSupplierRequestApprovals:
                    break;
                case ActionItemType.UnassignedPurchaseOrders:
                    break;
                case ActionItemType.UnassignedInvoices:
                    break;
                case ActionItemType.UnassignedSupplierRegistrations:
                    break;
                case ActionItemType.UnassignedSupplierRequests:
                    break;
                case ActionItemType.PriceFileToReview:
                    break;
                case ActionItemType.ImportExportCompleted:
                    break;
                case ActionItemType.SupplierDataReview:
                    break;
                default:
                    break;
            }
        }
        public void QuickSearch(string search, string searchOption)
        {
            try
            {
                driver.SelectOption("//select[@id='Phoenix_Notifications_QuickSearchType']/option[.='" + searchOption + "']");
                driver.FillField("//input[@id='Phoenix_Notifications_QuickSearchTerms' and @type='text']", search);
                driver.ClickButton("//button[@id='Phoenix_Notifications_QuickSearchSubmit']");
                //driver.FillWebElement("//select[@id='Phoenix_Notifications_QuickSearchType']", text: searchOption);
                //driver.FillWebElement("//input[@id='Phoenix_Notifications_QuickSearchTerms' and @type='text']",text: search);
                //driver.ClickWebElement("//button[@id='Phoenix_Notifications_QuickSearchSubmit']");
            }
            catch (Exception e)
            {
                return;
            }
        }
        public HomePage Home() { driver.ClickWebElement("//a[@aria-label='Home']"); return new HomePage(driver); }
        //public void Orders() { driver.ClickWebElement("//a[@aria-label='Orders']"); }
        //public void Contracts() { driver.ClickWebElement("//a[@aria-label='Contracts']"); }
        //public void AccountsPayable() { driver.ClickWebElement("//a[@aria-label='Accounts Payable']"); }
        //public void Shop() { driver.ClickWebElement("//a[@aria-label='Shop']"); }
        //public void Supplier() { driver.ClickWebElement("//a[@aria-label='Suppliers']"); }
        //public void Administer() { driver.ClickWebElement("//a[@aria-label='Administer']"); }
        //public void SetUp() { driver.ClickWebElement("//a[@aria-label='SetUp']"); }
    }
    public class ActionItems
    {
        public IWebDriver Driver { get; set; }
        public List<ActionItem> Items { get; set; }
        public ActionItems(IWebDriver driver)
        {
            Driver = driver;
        }
    }
    public class ActionItem
    {
        public string Parent;
        public string Name { get; private set; }
        public string Section { get; set; }
        public IWebDriver Driver { get; set; }
        public int Count { get; set; }
        public void Click()
        {
            // //h5[text()='Administrative Items']/../ul/li/span/a[@title='Price Files To Review']
           // Driver.ClickWebElement()
        }
        public ActionItem(string name, IWebDriver driver)
        {
            Driver = driver;
            Name = name;
            Section = Driver.GetStringFromWebElement("//a[@title='"+name+"']/../../../../h5");
            Count = Driver.GetNumberFromWebElement("//a[@title='" + name + "']/../../div");
        }
    }

    namespace TSM
    {
        public class PayeeProfile
        {
            public int BENNumber
            {
                get
                {
                    return Driver.GetNumberFromWebElement("(//div[text()='BEN Number']/../../../div)[2]/div");
                }
            }
            public string JaggaerID { get
                {
                    // string id = HttpUtility.ParseQueryString(href).Get("CMMSP_SupplierID");
                    return Driver.FindElement(By.XPath("(//input[@name='ResultsSelectedId' and @type='hidden'])[1]")).GetAttribute("value");

                }
            }
            public IWebDriver Driver { get; }
            public string RegistrationStatus
            {
                get
                {
                    return Driver.GetStringFromWebElement("(//div[text()='Registration Status']/../../../div)[2]/div");
                }
            }
            public string RegistrationType
            {
                get
                {
                    string test = Driver.GetStringFromWebElement("(//div[text()='Registration Type']/../../../div)[2]/div/div");
                    return test;
                }
            }

            public PayeeProfile(IWebDriver driver)
            {
                Driver = driver;
                string href = driver.FindElement(By.XPath("//a[contains(@href,'CMMSP_SupplierID=')]")).GetAttribute("href");

                //JaggaerID = HttpUtility.ParseQueryString(href).Get(0);
            }
  /*          public void Select(string menu, string menuitem)
            {
                if (Driver.FindElements(By.XPath("//a[@aria-expanded='false']/span[text()='" + menu + "']")).Count == 1)
                {
                    Driver.ClickWebElement("//a[@aria-expanded='false']/span[text()='" + menu + "']");
                }
                Driver.ClickWebElement("//ul/li/a[text()='" + menuitem + "']");
            }
    */
            public MenuResult Select(string menu, string menuitem)
            {
                if (Driver.FindElements(By.XPath("//a[@aria-expanded='false']/span[text()='" + menu + "']")).Count == 1)
                {
                    Driver.ClickWebElement("//a[@aria-expanded='false']/span[text()='" + menu + "']");
                }
                Driver.ClickWebElement("//ul/li/a[text()='" + menuitem + "']");
                return new MenuResult(Driver);
            }
 /*           public MenuResult Select(string menu, string menuitem)
            {
                if (Driver.FindElements(By.XPath("//a[@aria-expanded='false']/span[text()='" + menu + "']")).Count == 1)
                {
                    Driver.ClickWebElement("//a[@aria-expanded='false']/span[text()='" + menu + "']");
                }
                Driver.ClickWebElement("//ul/li/a[text()='" + menuitem + "']");
                return new MenuResult(Driver);
            }
 */
              
            public void Save()
            { }            
        }
        public class SupplierSearchResults
        {
            public SupplierSearchResults(IWebDriver driver)
            {
                Driver = driver;
            }
            public IWebDriver Driver { get; set; }
            public int Count
            {
                get
                {
                    return Driver.FindElements(By.XPath("//button[text()='Manage']")).Count;
                }
            }
            public PayeeProfile Next()
            {
                // TODO: 

                // click Slider Next
                // return null if at the first result.
                return new PayeeProfile(Driver);
            }
            public PayeeProfile Prev()
            {
                // click Slider Previous 
                // Return Null if already at the first result.
                return new PayeeProfile(Driver);
            }

            public PayeeProfile Select(int index = 1)
            {
                var el = Driver.FindElement(By.XPath("//li/a[starts-with(@onclick,'editOrViewProfile(')][" + index.ToString() + "]"));

                Driver.ExecuteJavaScript(el.GetAttribute("onclick"));
                PayeeProfile result = new PayeeProfile(Driver);
                return result;
            }
            public PayeeProfile Select(String name)
            {
                PayeeProfile result = null;
                var selected = Driver.ClickWebElement("//a[text()='" + name + "']");
                if (selected != null)
                    result = new PayeeProfile(Driver);
                return result;
            }
        }
    }

    namespace TestGrounds
    {
        [TestClass]
        public class UnitTest
        {
            IWebDriver driver;
            [TestInitialize]
            public void SetUp()
            {
                driver = new ChromeDriver("C:\\Selenium");
                driver.Url = "https://solutions.sciquest.com/apps/Router/Login?OrgName=UPenn";
                driver.Manage().Window.Maximize();
                // driver.FindElement(By.Id("Username")).SendKeys("caputo");
                //driver.FindElement(By.Id("Password")).SendKeys("blue72");
                //      driver.FindElement(By.Id("Username")).SendKeys("Brian.Caputo");
                driver.FindElement(By.Id("Username")).SendKeys("Otto.Mate");
                driver.FindElement(By.Id("Password")).SendKeys("Aut0m@te88");
                driver.FindElement(By.TagName("button")).Click();
            }

            [TestCleanup]
            public void CleanUp()
            {
                driver.Close();
            }

            [TestMethod]
            public void setPMExceptions()
            {
                string FileName = "C:\\Selenium\\Payment Terms change.xlsx";
                HomePage hp = new HomePage(driver);
                WorkBook wb = WorkBook.Load(FileName);
                WorkSheet ws = wb.DefaultWorkSheet;
                for (int i = 1; i < ws.RowCount; i++)
                {
                    string VID = ws.GetCellAt(i, 1).StringValue;
                    string VendorName = ws.GetCellAt(i, 0).StringValue;
                    hp.QuickSearch(VID, "Supplier Profile");
                    SupplierSearchResults results = new SupplierSearchResults(driver);
                    PayeeProfile pp = null;
                    if (results.Count > 0)
                        pp = (results.Count == 1) ? results.Select() : results.Select(VendorName);
                    else
                    {
                        //error  Supplier not found.
                        ws.SetCellValue(i, 4, "No Matching Suppliers found?.");
                        ws.SetCellValue(i, 5, DateTime.Now.ToString("dddd, dd MMMM yyyy"));
                        wb.Save();
                    }
                    try
                    {
                        pp.Select("Accounts Payable", "Payment Custom Fields");
                        PaymentCustomField pcf = new PaymentCustomField(driver);
                        pcf.Terms = ws.GetCellAt(i, 2).ToString();
                        pcf.Basis = ws.GetCellAt(i, 3).ToString();
                        pcf.Save();

                        ws.SetCellValue(i, 4, "Success");
                        ws.SetCellValue(i, 5, DateTime.Now.ToString("dddd, dd MMMM yyyy"));
                    }
                    catch (Exception e)
                    {
                        ws.SetCellValue(i, 4, e.Message);
                        ws.SetCellValue(i, 5, DateTime.Now.ToString("dddd, dd MMMM yyyy"));
                    }
                    wb.Save();
                }
            }
            [TestMethod]
            public void TestMethod1()
            {
                /*
                            HomePage hp = new HomePage(driver);
                            //         hp.SelectMenu("Suppliers", "Manage Suppliers","Search for a Supplier");
                            hp.QuickSearch("51793", "Supplier Profile");
                            var results = new SupplierSearchResults(driver);
                            PayeeProfile pp = null;

                            switch (results.Count)
                            {
                                case 0:
                                    throw new Exception("Supplier not found!");
                                case 1:
                                    pp = results.Select();
                                    break;
                                default:
                                    pp = results.Select("Fisher Scientific");//??
                                    break;
                            }
                            /*            pp.Select("About", "Supplier Classes");
                                        SupplierClasses sc = new SupplierClasses(driver);
                                        sc.POSupplier = true;
                                        sc.POSupplier = false;
                             */
                /*            pp.Select("Contacts and Addresses", "Addresses");
                            Addresses adds = new Addresses(driver);
                            Address add = adds.Select(2);
                            string country = add.Country;

                            string street = add.Address1;
                            Address add2 = adds.Select("BOSTON-15", false);
                            int conut2 = adds.Count;
                            int poCount = adds.getPOSiteCount(true);
                            poCount = adds.getPOSiteCount(); 

                */
                HomePage hp = new HomePage(driver);
                string DeActivatePath = "C:\\Selenium\\SupplierSite_Purge_FY2020_16Nov_TEST.xlsx";
                WorkBook wb = WorkBook.Load(DeActivatePath);
                WorkSheet ws = wb.GetWorkSheet("Site Activity");
                string prevVID = "";
                PayeeProfile pp = null;

                for (int i = 1; i < ws.RowCount; i++)
                {
                    string VID = ws.GetCellAt(i, 1).StringValue;
                    string VendorName = ws.GetCellAt(i, 0).StringValue;
                    if (VID != prevVID)// get the new supplier profile..
                    {
                        hp.QuickSearch(VID, "Supplier Profile");
                        SupplierSearchResults results = new SupplierSearchResults(driver);

                        if (results.Count > 0)
                        {
                            pp = (results.Count == 1) ? results.Select() : results.Select(VendorName);
                            prevVID = VID;
                        }
                        else// error on the supplier quick search 
                        {
                            ws.SetCellValue(i, 6, "No Matching Suppliers found?.");
                            ws.SetCellValue(i, 7, DateTime.Now.ToString());
                            wb.Save();
                        }
                    }
                    //make sure addresses and the payee profile are available
                    if (pp != null)
                    {
                        pp.Select("Contacts and Addresses", "Addresses");

                        Addresses adds = new Addresses(driver);
                        string siteName = ws.GetCellAt(i, 2).StringValue;
                        string siteID = ws.GetCellAt(i, 3).StringValue;
                        string PaySite = ws.GetCellAt(i, 4).StringValue;
                        string POSite = ws.GetCellAt(i, 5).StringValue;
                        try
                        {
                            if (PaySite == "Y")
                            {
                                Address add = adds.Select(siteName, false);
                                if (add.Active)
                                {
                                    add.Active = false;
                                    add.Save();
                                }
                            }

                            if (POSite == "Y")
                            {
                                Address add = adds.Select(siteName, true);

                                if (add.Active)
                                {
                                    add.Active = false;
                                    add.Save();

                                    //var el2 = driver.FindElement(By.XPath("//input[@id='Phoenix_Notifications_QuickSearchTerms']"));
                                    //driver.ExecuteJavaScript("arguments[0].value=arguments[1];", el2, "TEST ARGUMENT");
                                    //TODO adds.GetPOSiteCount() fails and returns old values
                                    if (adds.getPOSiteCount() == 0)
                                    {
                                        pp.Select("About", "Supplier Classes");
                                        SupplierClasses sc = new SupplierClasses(driver);
                                        sc.POSupplier = false;
                                        sc.Save();
                                    }
                                }
                            }
                            ws.SetCellValue(i, 8, "Corrected");
                            ws.SetCellValue(i, 9, DateTime.Now.ToString());
                            wb.Save();
                        }
                        catch (Exception x)
                        {
                            ws.SetCellValue(i, 6, "Error!! Moving on..");
                            ws.SetCellValue(i, 7, x.Message);
                            wb.Save();
                        }
                    }

                    else
                    {
                        ws.SetCellValue(i, 6, "Error!--moving on....");
                        ws.SetCellValue(i, 7, "Supplier not found!");
                        wb.Save();
                    }

                }


            }
            public string DeactivateSupplierSite()
            {
                return "";
            }
            public string DeactivateSites()
            {
                return "";
            }
        }
    }
}