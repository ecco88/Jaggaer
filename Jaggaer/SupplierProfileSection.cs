using SeleniumExtras;
using OpenQA.Selenium;
using SeleniumHelpers;
using OpenQA.Selenium.Support.Extensions;
using System.Runtime.InteropServices;
using System.Runtime.Hosting;

namespace Jaggaer
{
    namespace TSM
    {
        public class MenuResult
        {
            public IWebDriver Driver { get; set; }
            public MenuResult(IWebDriver driver)
            {
                Driver = driver;
            }
        }
        public class SupplierProfileSection
        {
            public IWebDriver Driver { get; set; }
            public void Save()
            {

                //var ele = Driver.FindElement(By.XPath("//form[@name='ActiveForm']"));
                //Driver.ExecuteJavaScript("arguments[0].submit();", ele);
                var ele = Driver.FindElement(By.XPath("//input[@type='submit' and @value='Save']"));
                Driver.ExecuteJavaScript("arguments[0].click();", ele);
            }
            public SupplierProfileSection(IWebDriver driver)
        {
                Driver = driver;
            }
        }
        public class SupplierClasses : SupplierProfileSection
        {
            public SupplierClasses(IWebDriver driver) : base(driver)
            {
                Driver = driver;
            }
            private bool GetChecked(string name)
            {
                return (Driver.FindElement(By.XPath("(//span[text()='" + name + "']/../..//input)[1]")).GetAttribute("checked").ToLower() == "checked");
            }
            private void SetChecked(string name, bool val)
            {
                var eles = Driver.FindElements(By.XPath("(//span[text()='" + name + "']/../..//input)"));
                if (!eles[0].Selected)
                    Driver.ExecuteJavaScript("arguments[0].click();", eles[0]);
                if (eles[1].Selected != val)
                    Driver.ExecuteJavaScript("arguments[0].click();", eles[1]);
            }
            public void Save()
            {
                var ele = Driver.FindElement(By.XPath("//input[@type='button' and @value='Save']"));
                Driver.ExecuteJavaScript("arguments[0].click();", ele);
            }
            public bool POSupplier
            {
                get
                {
                    return GetChecked("PO Supplier");
                }
                set
                {
                    SetChecked("PO Supplier", value);
                }
            }
        }
        public class Addresses : MenuResult
        {
            public Addresses(IWebDriver driver) :base(driver)
            {
                driver.ClickWebElement("//a[.='Show Inactive Addresses']");
            }
            public bool ShowInactiveSites 
            {
                get 
                {
                    return Driver.FindElements(By.XPath("//a[.='Show Inactive Addresses']")).Count != 0;
                }
                set
                {
                    int count = Driver.FindElements(By.XPath("//a[.='Show Inactive Addresses']")).Count;
                    if (count == 1) Driver.ClickWebElement("//a[.='Show Inactive Addresses']");
                }
            }
            public int Count 
            {
                get
                {
                    return Driver.FindElements(By.XPath("//table[@class='SearchResults']/tbody/tr")).Count;
                }
            }
            public int getSiteCount(bool activeOnly = true)
            {
                return Driver.FindElements(By.XPath("//table[@class='SearchResults']/tbody/tr/td/a[@class='"+((activeOnly)?"ListActive":"ListInactive")+"']")).Count;
            }
            public int getPOSiteCount(bool activeOnly=true)
            {
                string xp = "//table[@class='SearchResults']/tbody/tr/td/a[(@class='ListSelected' or @class='" + ((activeOnly) ? "ListActive" : "ListInactive") + "') and 'Fulfillment)'=substring(.,string-length(.)-string-length('Fulfillment)')+1)]";
                return (Driver.FindElements(By.XPath(xp))).Count;
            }
            public int getPaySiteCount(bool activeOnly=true)
            {
                string xp = "//table[@class='SearchResults']/tbody/tr/td/a[(@class='ListSelected' or @class='" + ((activeOnly) ? "ListActive" : "ListInactive") + "') and 'Remittance)'=substring(.,string-length(.)-string-length('Remittance)')+1)]";
                return (Driver.FindElements(By.XPath(xp))).Count;
            }
            public Address Select(int index=1)
            {
                var link = Driver.FindElement(By.XPath("(//table[@class='SearchResults']/tbody/tr/td/a)[" + index + "]"));
                Driver.ExecuteJavaScript("arguments[0].click();", link);
                return new Address(Driver);
            }
            public Address Select(string name, bool POsite)
            {
                string x2 = POsite ? "Fulfillment" : "Remittance";
                string xp = "//table[@class='SearchResults']/tbody/tr/td/a[starts-with(.,'" + name + "') and '" + x2 + ")'=substring(., string-length(.) - string-length('" + x2 + ")') +1)]";
                Driver.ClickLink(xp);
               // string xpath = "//table[@class='SearchResults']/tbody/tr/td/a[starts-with('" + name + " (',text()') and ends-with('" + ((POsite) ? "Fulfillment" : "Remittance") + ")',text())]'";
                //var link = Driver.FindElement(By.XPath(xp));              
                //Driver.ExecuteJavaScript("arguments[0].click();", link);
                return new Address(Driver);
            }
            public Address Select(string name)
            {
                Driver.ExecuteJavaScript("arguments[0].click();", Driver.FindElement(By.XPath("//table[@class='SearchResults']/tbody/tr/td/a[starts-with(text(),'" + name + "')]")));
                return new Address(Driver);
            }
            private bool GetChecked(string name)
            {
                return (Driver.FindElement(By.XPath("(//span[text()='" + name + "']/../..//input)[1]")).GetAttribute("checked").ToLower() == "checked");
            }
            private void SetChecked(string name, bool val)
            {
                var eles = Driver.FindElements(By.XPath("(//span[text()='" + name + "']/../..//input)"));
                if (!eles[0].Selected)
                    Driver.ExecuteJavaScript("arguments[0].click();", eles[0]);
                if (eles[1].Selected != val)
                    Driver.ExecuteJavaScript("arguments[0].click();", eles[1]);
            }
        }

        public class PaymentCustomField : MenuResult
        {
            public PaymentCustomField(IWebDriver driver) : base(driver)
            {

            }
            public string Terms
            {
                get
                {
                    return Driver.GetFormValue("//select[@name='CustElement_8177']");
                }
                set
                {
                    Driver.FillWebElement("//select[@name='CustElement_8177']", value);
                }
            }
            public string Basis
            {
                get
                {
                    return Driver.GetFormValue("//select[@name='CustElement_8171']");
                }
                set
                {
                    Driver.FillWebElement("//select[@name='CustElement_8171']", value);
                }
            }
            public void Save()
            {
                Driver.ClickButton("//input[@value='Save' and @type='submit']");
            }
        }
        public class Address : SupplierProfileSection
        {
            public Address(IWebDriver driver) : base(driver) { Driver = driver; }
            public string Name {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='Name']/../../td)[2]/input[@type='text']");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='Name']/../../td)[2]/input[@type='text']", value);
                }
            }
            public string AddressType
            {
                get
                {
                    return Driver.GetStringFromWebElement("(//tr/td/a[text()='Address Type']/../../td)[2]");
                }
            }
            public string AddressID {
                get
                {
                    return Driver.GetStringFromWebElement("(//tr/td/a[starts-with(.,'Address ID')]/../../td)/../td[2]/span");
                }
                set
                {
                    //TODO
                }
            }
            public string ThirdPartyID
            {
                get
                {
                    return Driver.GetStringFromWebElement("(//tr/td/a[text()='3rd Party ID']/../../td)[2]/span");
                }
                set
                {
                    //TODO
                }
            }
            public bool Active
            {
                get
                {
                    string test = "";
                    test= Driver.FindElement(By.XPath("(//tr/td/a[text()='Active']/../../td)[2]/input")).GetAttribute("checked");
                    //is it 'checked' or 'true'?
                    return test == "true";
                }
                set
                {
                    if (Active!=value)
                    {
                        Driver.ToggleCheckBox("(//tr/td/a[text()='Active']/../../td)[2]/input");
                    }
                }
            }
            public bool Primary
            {
                get
                {
                    return Driver.FindElement(By.XPath("(//tr/td/a[text()='Primary']/../../td)[2]/input")).GetAttribute("checked") == "checked";
                }
                set
                {
                    if(Primary!=value)
                    {
                        Driver.ToggleCheckBox("(//tr/td/a[text()='Primary']/../../td)[2]/input");
                    }
                    //TODO
                }
            }
            public string Country
            {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='Country']/../../td)[2]/select");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='Country']/../../td)[1]/select", value);
                }
            }
            public string Address1 {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='Street Line 1']/../../td)[2]/input[@type='hidden']");
                   // return Driver.FindElement(By.XPath("(//tr/td/a[text()='Street Line 1']/../../td)[2]/input[@type='hidden']")).GetAttribute("value");
                   // return Driver.GetStringFromWebElement("(//tr/td/a[text()='Street Line 1']/../../td)[2]/input[@type='hidden'");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='Street Line 1']/../../td)[2]/input[1]", value);
                }
            }
            public string Address2
            {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='Street Line 2']/../../td)[2]/input[@type='hidden']");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='Street Line 2']/../../td)[2]/input[1]", value);
                }
            }
            public string Address3
            {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='Street Line 3']/../../td)[2]/input[@type='hidden']");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='Street Line 3']/../../td)[2]/input[1]", value);
                }
            }
            public string City
            {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='City/Town']/../../td)[2]/input[@type='hidden']");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='City/Town']/../../td)[2]/input[1]", value);
                }
            }
            public string State
            {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='State/Province']/../../td)[2]/input[@type='hidden']");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='State/Province']/../../td)[2]/input[1]", value);
                }
            }
            public string PostalCode
            {
                get
                {
                    return Driver.GetFormValue("(//tr/td/a[text()='Postal Code']/../../td)[2]/input[@type='text']");
                }
                set
                {
                    Driver.FillWebElement("(//tr/td/a[text()='Postal Code']/../../td)[2]/input[1]", value);
                }
            }
            public string Notes { 
                get
                {
                    return Driver.GetStringFromWebElement("//textarea");
                }
                set
                {
                    Driver.FillWebElement("//textarea", value);
                }
            }
        }
        public class FulfillmentCenters :MenuResult
        {
            public FulfillmentCenters(IWebDriver driver) : base(driver)
            {

            }
            public int Count {
                get
                {
                    throw new System.Exception("Not Implemented......");
                    int result = 0;
                    return result;
                } 
            }
            public FulfillmentCenter Select(string name)
            {
                Driver.ClickLink("//table[@class='SearchResults']/tbody/tr/td/a[.='" + name + "')]");
                return new FulfillmentCenter(Driver);
            }

        }
        public class FulfillmentCenter :SupplierProfileSection
        {
            public FulfillmentCenter(IWebDriver driver):base (driver)
            {

            }
        }
    }
}