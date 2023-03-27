using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI.DTO;
using Microsoft.Dynamics365.UIAutomation.Api.UCI.Constant;
using System;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using System.Text;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Data.OleDb;
using System.Data;
using HtmlAgilityPack;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI
{
    public class HelperMethods : Element
    {
        public const string PhoneNumberFormat = @"^\(?([0-9]{3})\)?[-. ]?([0-9]{3})[-. ]?([0-9]{4})$";
        public const string EmailFormat = @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z";
        public const string WebsiteFormat = @"((http|https)://)(www.)?[a-zA-Z0-9@:%._\\+~#?&//=]{2,256}\\.[a-z]{2,6}\\b([-a-zA-Z0-9@:%._\\+~#?&//=]*)";
        private const string GRIDXPATH = "//div[contains(@id, 'entity_control-pcf_grid_control_container')] //*[@data-id='grid-container']";
        
        // private const string GRIDXPATH = "( //div[contains(@id, 'entity_control-pcf_grid_control_container')] //*[@data-id='grid-container']  | //*[@data-id='data-set-body-container' and //*[@class='wj-cells'] ] )";
        //private const string GRID_SCROLL_LEFT_COMMAND = "document.querySelectorAll(\"[id='entity_control-pcf_grid_control_container'] div[class='ag-body-horizontal-scroll'] div[ref='eViewport']\").scrollBy(-400, 0);";
        //private const string GRID_SCROLL_ALL_RIGHT_COMMAND = "document.querySelectorAll(\"[id='entity_control-pcf_grid_control_container'] div[class='ag-body-horizontal-scroll'] div[ref='eViewport']\").scrollBy(12000, 0);";
        //private const string GRID_QUERY_HOW_MUCH_UNTIL_RESET = "return document.querySelectorAll(\"[id='entity_control-pcf_grid_control_container'] div[class='ag-body-horizontal-scroll'] div[ref='eViewport']\").scrollLeft;";
        //use HDR=YES if first excel row contains headers, HDR=NO means your excel's first row is not headers and it's data.
        public static void excelDataReadandcreateNewAccounts(XrmApp xrm, string connString)
        {
            // Create the connection object
            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();                 
                // Create OleDbCommand object and select data from worksheet Sample-spreadsheet-file
                // Here sheet name is Sample-spreadsheet-file, usually it is Sheet1, Sheet2 etc..
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Accounts$]", oledbConn);                 
                // Create new OleDbDataAdapter
                OleDbDataAdapter oleda = new OleDbDataAdapter(); 
                oleda.SelectCommand = cmd;                 
                // Create a DataSet which will hold the data extracted from the worksheet.
                DataSet ds = new DataSet();                 
                // Fill the DataSet from the data extracted from the worksheet.
                oleda.Fill(ds, "Employees");
                foreach (var m in ds.Tables[0].DefaultView)
                {
                    xrm.CommandBar.ClickCommand("New");
                    var AccountName = ((System.Data.DataRowView)m).Row.ItemArray[0];
                    var Phone = ((System.Data.DataRowView)m).Row.ItemArray[1];
                    var Website = ((System.Data.DataRowView)m).Row.ItemArray[2];
                    var RelationshipType = ((System.Data.DataRowView)m).Row.ItemArray[3];
                    var ParentAccount = ((System.Data.DataRowView)m).Row.ItemArray[4];
                    var ProductPriceList = ((System.Data.DataRowView)m).Row.ItemArray[5];

                    xrm.Entity.SetValue("name", AccountName.ToString());
                    xrm.Entity.SetValue("telephone1", Phone.ToString());
                    xrm.Entity.SetValue("websiteurl", Website.ToString());
                    xrm.Entity.SetValue(new OptionSet { Name = "customertypecode", Value = RelationshipType.ToString() });
                    xrm.Entity.SetValue(new LookupItem { Name = "parentaccountid", Value = ParentAccount.ToString(), Index = 0 });
                    //xrm.Entity.SetValue(new LookupItem { Name = "defaultpricelevelid", Value = "France Bill Rates", Index = 0 });
                    xrm.Entity.Save();
                    xrm.ThinkTime(5000);

                }


            }
            catch (Exception e)
            {
                Console.WriteLine("Error :" + e.Message);
            }
        }
        public static void GetAllRequiredFields(XrmApp xrmApp, WebClient client)
        {
            var TotalReqFields = client.Browser.Driver.FindElements(By.XPath("//input[@aria-required='true']"));

            List<string> RequriedLabel = new List<string> { };

            if (TotalReqFields != null)
            {
                for (int i = 1; i <= TotalReqFields.Count; i++)
                {
                    var formLabels = client.Browser.Driver.FindElements(By.XPath("//div[@data-fieldrequirement=" + i + "]//label"));

                    foreach (var fieldLabel in formLabels)
                    {
                        RequriedLabel.Add(fieldLabel.Text);
                        xrmApp.Entity.SetValue(fieldLabel.Text, fieldLabel.Text + "23123");
                    }
                }
                xrmApp.ThinkTime(5000);
            }
        }

        public static Dictionary<string, string> EntityDisplayNames = new Dictionary<string, string>()
        {
            { "Accounts", "Account"},
            { "Bookings", "Booking"},
            { "Bookable Resources", "Bookable Resource"},
            { "Booking Statuses", "Booking Status"}
        };

        public static Dictionary<string, string> EntityLogicalNames = new Dictionary<string, string>()
        {
            { "Accounts","account" },
            { "Bookings","bookableresourcebooking" },
            { "Bookable Resources","bookableresource" }
        };        

        private const int DefaultThinkTime = 2000;

        public static void CreateTestAccount(XrmApp xrm, string BusinessName)
        {
            Random random = new Random();
            string AccountNumber = "Account" + random.Next(10, 9999).ToString();
            xrm.Entity.SetValue("name", BusinessName + random.Next(10, 9999).ToString());
            xrm.Entity.SetValue("telephone1",string.Format("9{0}", GeneratRandomNumber(9)));
            xrm.Entity.SetValue("websiteurl","www.test"+ random.Next(1, 99).ToString() + ".com");
            xrm.Entity.SetValue(new OptionSet { Name = "customertypecode", Value = "Other" });
            xrm.Entity.SetValue(new LookupItem { Name = "parentaccountid", Value = "A Datum Corporation", Index = 0 });
            //xrm.Entity.SetValue(new LookupItem { Name = "defaultpricelevelid", Value = "France Bill Rates", Index = 0 });
            xrm.ThinkTime(5000);
            
        }

        public static void CreateContact(XrmApp xrm, string BusinessName)
        {
            Random random = new Random();
            string AccountNumber = "Contact" + random.Next(100, 9999).ToString();
            xrm.Entity.SetValue("fullname_compositionLinkControl_firstname", BusinessName + random.Next(100, 9999).ToString());
            xrm.Entity.SetValue("fullname_compositionLinkControl_lastname", random.Next(100, 9999).ToString());
            xrm.Entity.SetValue("telephone1", string.Format("9{0}", GeneratRandomNumber(9)));

            xrm.Entity.SetValue("emailaddress1", "Test" + random.Next(1, 99).ToString() + "@gmail.com");
            
            xrm.ThinkTime(5000);
           
        }

        public static string GeneratRandomNumber(int length)
        {
            if (length > 0)
            {
                var sb = new StringBuilder();

                var rnd = SeedRandom();
                for (int i = 0; i < length; i++)
                {
                    sb.Append(rnd.Next(0, 9).ToString());
                }

                return sb.ToString();
            }

            return string.Empty;
        }
        private static Random SeedRandom()
        {
            return new Random(Guid.NewGuid().GetHashCode());
        }

        public static List<string> GetCoreField(List<string> NonCoreFieldLabels, List<string> LabelList)
        {
            foreach (string label in NonCoreFieldLabels)
            {
                LabelList.Remove(label);
            }
            return LabelList;
        }

        public static ReadOnlyCollection<IWebElement> Form_GetLabels(XrmApp xrm, WebClient client, bool debug = false)
        {
            //Method returns a collection of all the forms labels that can then be converted into a list
            var xpathFormLabels = "//div[contains(@data-lp-id,'MscrmControls.Containers.FieldSectionItem|')]//label";
            var formLabels = client.Browser.Driver.FindElements(By.XPath(xpathFormLabels));
            return formLabels;
        }

        public static List<string> Labels_CreateList(XrmApp xrm, ReadOnlyCollection<IWebElement> labelList, bool debug = false)
        {
            List<string> returnLabelList = new List<string> { };
            if (labelList == null) return returnLabelList;

            foreach (var fieldLabel in labelList)
            {
                if (fieldLabel.Text != "")
                {
                    //Build our reference List
                    returnLabelList.Add(fieldLabel.Text);
                }
            }

            //check to see if we should use a value instead of Text
            if (returnLabelList.Count == 0)
            {
                foreach (var fieldLabel in labelList)
                {
                    if (fieldLabel.GetAttribute("value") != "")
                    {
                        //Build our reference List
                        returnLabelList.Add(fieldLabel.GetAttribute("value"));
                    }
                }
            }
            return returnLabelList;
        }

        public static bool CompareStringLists(List<string> referenceList, List<string> listForComparison, bool inOrderCheck = true, bool debug = false)
        {
            var counter = 0;
            List<string> additionalListItems = new List<string> { };
            if (debug)
            {
                Debug.WriteLine("***** Start of HM.CompareLists() Debug Output *****");
                if (referenceList.Count != listForComparison.Count)
                {
                    Debug.WriteLine("HM.CompareLists(): ReferenceList count of " + referenceList.Count() + " items does NOT match listForComparison count of " + listForComparison.Count() + " items");
                }
                else
                {
                    Debug.WriteLine("HM.CompareLists(): ReferenceList count matches listForComparison count of " + referenceList.Count() + " items");
                }
                Debug.WriteLine("HM.CompareLists() - referenceList");
                foreach (string element in referenceList)
                {
                    Debug.WriteLine("referenceList" + "[" + counter + "] = " + element.ToString());
                    counter++;
                }
                counter = 0;
                Debug.WriteLine("HM.CompareLists() - listForComparison");
                foreach (string element in listForComparison)
                {
                    Debug.WriteLine("listForComparison" + "[" + counter + "] = " + element.ToString());
                    counter++;
                }
                counter = 0;
                Debug.WriteLine("HM.CompareLists() - Comparison Results:");
                Debug.WriteLine("In Order Check is " + inOrderCheck.ToString());
                string result = "";
                if (inOrderCheck)
                {
                    foreach (string element in listForComparison)
                    {
                        //result = (referenceList[counter] == element) ? " matched." : " NO ORDERED MATCH FOUND!";
                        result = (element.Contains(referenceList[counter])) ? " matched." : " NO ORDERED MATCH FOUND!";

                        Debug.WriteLine("referenceList" + "[" + counter + "] = " + element.ToString() + result);

                        counter++;
                        if (counter > (referenceList.Count - 1)) break;
                    }
                }
                else
                {
                    foreach (string element in listForComparison)
                    {
                        result = (referenceList.Contains(element)) ? " found." : " NOT FOUND!";
                        if (result == " NOT FOUND!")
                        {
                            additionalListItems.Add(element);
                        }
                        Debug.WriteLine("referenceList" + "[" + counter + "] = " + element.ToString() + result);
                        counter++;
                        if (counter > (referenceList.Count - 1)) break;
                    }
                }
                counter = 0;
                if (additionalListItems.Count > 0)
                {
                    Debug.WriteLine(">>> ADDITIONAL " + additionalListItems.Count + " item(s) found in the Comparison List");
                    foreach (string element in additionalListItems)
                    {
                        Debug.WriteLine(element);
                    }
                }
                Debug.WriteLine("***** End of HM.CompareLists() Debug Output *****");
            }
            if (inOrderCheck)
            {
                foreach (string element in referenceList)
                {
                    if (listForComparison[counter].Contains(element) != true)
                    {
                        return false;
                    }
                    counter++;
                    if (counter > (referenceList.Count - 1)) break;
                }
            }
            else
            {
                foreach (string element in referenceList)
                {
                    if (listForComparison.Contains(element) != true)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public static bool ValidateWebsite(Action website)
        {
            throw new NotImplementedException();
        }

        public static BrowserCommandResult<List<string>> GetTabs(XrmApp xrm, WebClient client, bool debug = false)
        {

            List<string> tabList = new List<string>();
            IWebElement tabListElement = client.Browser.Driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TabList]));
            var tabs = tabListElement.FindElements(By.XPath(".//li"));
            foreach (var tab in tabs)
            {
                tabList.Add(tab.Text);
                if (tab.GetAttribute("aria-label") == "More Tabs")
                {
                    tab.Click();
                    IWebElement flyoutTabElement = client.Browser.Driver.WaitUntilAvailable(By.Id("__flyoutRootNode"));
                    var flyoutTabs = flyoutTabElement.FindElements(By.XPath(".//li[contains(@data-id,'tablist')]"));
                    foreach (var flyoutTab in flyoutTabs)
                    {
                        tabList.Add(flyoutTab.Text);
                    }
                }
            }
            tabList.Remove("Related");

            if (debug) HelperMethods.DebugOutputForReferenceListCreation(tabList);
            return tabList;
        }

        public static void DebugOutputForReferenceListCreation(List<string> list)
        {
            // Method writes to the output window each string ina  List, in order to be used in a List<string> referenceList definition.
            // This saves time over creating a manual reference lists
            // Simply, we can copy the whole list into a reference list definition and remove the very last comma

            Debug.WriteLine("List<string> referenceList = new List<string> {");
            if (list.Count > 0)
            {
                foreach (string label in list)
                {
                    Debug.WriteLine($"\t\"{label}\",");
                }
            }
            Debug.WriteLine("};");
        }

        public static List<string> GetTabRequiredFields(XrmApp xrm, WebClient client, string tabname, bool debug = false)
        {
            xrm.Entity.SelectTab(tabname);
            HelperMethods.WaitForInvisibilityOfProgressIndicator(xrm, client);
            List<string> RequriedLabel = new List<string> { };
            //Get required elements     
            var requiredLabelElements = client.Browser.Driver.FindElements(By.XPath("//div[contains(@data-id, '-required-icon')]"));

            for (int i = 1; i <= requiredLabelElements.Count; i++)
            {
                var formLabels = client.Browser.Driver.FindElements(By.XPath("//div[@data-fieldrequirement=" + i + "]//label"));


                foreach (var fieldLabel in formLabels)
                {
                    if (fieldLabel.Text != "")
                    {
                        //Build our reference List
                        RequriedLabel.Add(fieldLabel.Text);
                    }
                }

            }

            return RequriedLabel;
        }

        public static void WaitForInvisibilityOfProgressIndicator(XrmApp xrm, WebClient client)
        {
            string elementID = "appProgressIndicatorContainer";
            TimeSpan timeout = TimeSpan.FromSeconds(120);
            WebDriverWait wait = new WebDriverWait(client.Browser.Driver, timeout);

            // wait for element to disappear

            Func<IWebDriver, bool> elementIsInvisible =
            d =>
            {
                IWebElement e = d.FindElements(By.Id(elementID)).FirstOrDefault();
                return e == null;
            };

            wait.Until(elementIsInvisible);

            // TODO: Refactor to use WaitUntilClickable/etc. at a later date
            client.Browser.Driver.WaitForTransaction();
            xrm.ThinkTime(500);
        }

        public static void Grid_SwitchView(XrmApp xrm, WebClient client, string viewName, bool debug = false)
        {
            // Xpath
            string viewSelectorButtonXpath = "//button[contains(@id,'ViewSelector')]//i[@data-icon-name='ChevronDown']";
            string systemViewListXpath = "//div[contains(@data-id,'ViewSelector')]//li[contains(@class,'ContextualMenu-item')]";
            string selectedViewXpath = $"//div[contains(@data-id,'ViewSelector')]//li[contains(@class,'ContextualMenu-item')]//label[text()='{viewName}']";

            client.Browser.Driver.ClickWhenAvailable(By.XPath(viewSelectorButtonXpath), new TimeSpan(0, 0, 10), "View Selector Dropdown button was not found.");
            client.Browser.Driver.WaitForTransaction();

            string switchView = "//div[contains(@data-id,'ViewSelector')]//li[contains(@class,'ContextualMenu-item')]//label[contains(@title,'" + viewName + "')]";
            string switchViewControl = "//div[@aria-label='Views']//li[contains(@class,'ContextualMenu-item')]//button[@role='menuitemradio']//span[.//label[@title='" + viewName + ". System view']]";

            client.Browser.Driver.WaitUntilVisible(By.XPath(systemViewListXpath));
            if (client.Browser.Driver.HasElement(By.XPath(selectedViewXpath)))
            {
                var selectedView = client.Browser.Driver.FindElements(By.XPath(selectedViewXpath)).Where(x => x.Text == viewName).FirstOrDefault();
                if (selectedView != null)
                    selectedView.Click();
                else
                    throw new Exception("Could not find the specified view name in the view selector.");
                client.Browser.Driver.WaitForTransaction();
            }
            else if (client.Browser.Driver.HasElement(By.XPath(switchViewControl)))
            {
                client.Browser.Driver.FindElement(By.XPath(switchViewControl)).Click();
            }
            if (debug) HelperMethods.DebugOutputForReferenceListCreation(viewName);
        }

        public static void DebugOutputForReferenceListCreation(string list)
        {
            // Method writes to the output window a string passed in as an argument
            // This saves time over creating a manual reference lists
            // Simply, we can copy the whole list into a reference list definition and remove the very last comma
            Debug.WriteLine($"< string > \t\"{list}\",");
        }

        public static void OpenFirstRecord(XrmApp xrm, WebClient client, int thinkTime = DefaultThinkTime)
        {
            //Call the method with first element index == 0
            OpenRecord(xrm, client, 0);
        }

        public static void OpenRecord(XrmApp xrm, WebClient client, int index, bool doubleClick = false, int thinkTime = DefaultThinkTime)
        {
            string subGridEditButtonXpath = "//button[@aria-label='Edit']";
            string selectedSiteMapEntityXpath = "//li[contains(@id,'sitemap-entity') and @aria-selected='true']";
            string selectedEntityFormXpath = "//span[@data-id='entity_name_span' and text()='[NAME]']";
            string requestedEntityHyperlinkXpath = "//div[@class='ag-center-cols-container']//a[contains(@href,'[ENTITY]')]";

            //WaitForInvisibilityOfProgressIndicator(client);
            client.Browser.Driver.WaitUntilAvailable(By.XPath("//*[@data-id='GridRoot']"));

            // Select the record
            if (doubleClick)
                SelectRecord(xrm, client, index, doubleClick);
            else
            {
                SelectRecord(xrm, client, index, doubleClick);
                // Click the edit button to open the record
                var subGridEditButton = client.Browser.Driver.ClickIfVisible(By.XPath(subGridEditButtonXpath), new TimeSpan(0, 0, 5));
                if (subGridEditButton == null)
                {
                    // If edit button isn't visible, it's because of security role context or a system slow response time
                    client.ThinkTime(DefaultThinkTime);
                    string entityFormKey = EntityDisplayNames[client.Browser.Driver.FindElements(By.XPath(selectedSiteMapEntityXpath)).FirstOrDefault().Text];

                    var selectedEntityForm = client.Browser.Driver.WaitUntilVisible(By.XPath(selectedEntityFormXpath.Replace("[NAME]", entityFormKey)), new TimeSpan(0, 0, 3));
                    if (selectedEntityForm == null)
                    {
                        string entityKey = client.Browser.Driver.FindElements(By.XPath(selectedSiteMapEntityXpath)).FirstOrDefault().Text;
                        string selectedEntityLogicalName = EntityLogicalNames[entityKey];
                        var requestedEntityHyperlink = client.Browser.Driver.FindElements(By.XPath(requestedEntityHyperlinkXpath
                            .Replace("[ENTITY]", selectedEntityLogicalName))).ElementAtOrDefault(index);
                        if (requestedEntityHyperlink != null)
                        {
                            requestedEntityHyperlink.Click();
                            client.Browser.Driver.WaitForTransaction();
                            client.Browser.Driver.WaitUntilVisible(By.XPath(selectedEntityFormXpath.Replace("[NAME]", entityFormKey)), new TimeSpan(0, 0, 3));
                        }
                        else
                            throw new Exception($". Failed to locate 'Edit' button in Homepage Grid OR failed to find hyperlink to navigate to {entityKey} record form.");
                    }
                }
            }
            client.Browser.Driver.WaitForTransaction();
        }

        public static void SelectRecord(XrmApp xrm, WebClient client, int index, bool doubleClick = false, int thinkTime = DefaultThinkTime)
        {
            //NOTE: Review in future to see if we can change static wait times to waitfor(something)
            string xpathRecord = $"//*[@id='entity_control-pcf_grid_control_container']//div[@class='ag-center-cols-clipper']//div[@row-index='0']//div[contains(@class,'RowSelectionCheckMarkSpan')]";
            string editableRecordXpath = $"//div[@data-id='entity_control_container']//div[@role='checkbox']";
            bool hasPCFGridRecord = client.Browser.Driver.HasElement(By.XPath(xpathRecord));
            bool hasEditableGridRecord = client.Browser.Driver.HasElement(By.XPath(editableRecordXpath));

            IWebElement record = null;

            if (hasPCFGridRecord)
                record = client.Browser.Driver.FindElements(By.XPath(xpathRecord)).FirstOrDefault();
            else if (hasEditableGridRecord)
                record = client.Browser.Driver.FindElements(By.XPath(editableRecordXpath)).ElementAtOrDefault(index);

            try
            {
                if (record != null)
                {
                    if (doubleClick)
                        new Actions(client.Browser.Driver).MoveToElement(record).MoveByOffset(-10, 0).DoubleClick().Perform();
                    else
                        new Actions(client.Browser.Driver).MoveToElement(record).MoveByOffset(-10, 0).Click().Perform();
                    client.Browser.Driver.WaitForTransaction();
                }
                else
                    throw new Exception(". Failed to locate a record at the specified index.");
            }
            catch (Exception ex)
            {
                throw new Exception(ex + ". Failed to click or double click the record select element to open the record form.");
            }
        }

        #region Get Values
        ////public static string GetValue(XrmApp xrm, WebClient client, DateFieldItem control)
        ////{
        ////    var dateControlXpath = "//div[contains(@data-id,'[NAME].fieldControl._datecontrol-date-container')]//input";
        ////    var timeControlXpath = "//div[contains(@data-id,'[NAME].fieldControl._timecontrol-datetime-container')]//input";

        ////    var text = string.Empty;
        ////    var fieldContainer = client.Browser.Driver.WaitUntilAvailable(By.XPath(dateControlXpath.Replace("[NAME]", control.Name)));
        ////    var hasTimeControl = client.Browser.Driver.HasElement(By.XPath(timeControlXpath.Replace("[NAME]", control.Name)));

        ////    if (fieldContainer != null)
        ////    {
        ////        text = fieldContainer.GetAttribute<string>("value");
        ////    }
        ////    if (hasTimeControl)
        ////    {
        ////        var timeControlContainer = client.Browser.Driver.FindElement(By.XPath(timeControlXpath.Replace("[NAME]", control.Name)));
        ////        if (timeControlContainer != null)
        ////        {
        ////            text += " ";
        ////            text += timeControlContainer.GetAttribute<string>("value");
        ////        }
        ////    }

        ////    return text;
        ////}

        //public static string GetValue(XrmApp xrm, WebClient client, OptionSet option)
        //{
        //    var text = string.Empty;
        //    var fieldContainer = client.Browser.Driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));
        //    var xpathCheckboxSelect = $"//select[@data-id='{option.Name}.fieldControl-checkbox-select']";
        //    string xpathDropOptionSet = $"//select[contains(@data-id,'{option.Name}.fieldControl-option-set-select')]";

        //    var hasCheckboxSelect = client.Browser.Driver.HasElement(By.XPath(xpathCheckboxSelect));
        //    var hasOptionSet = client.Browser.Driver.HasElement(By.XPath(xpathDropOptionSet));
        //    if (hasCheckboxSelect)
        //    {
        //        var select = fieldContainer.FindElement(By.TagName("select"));
        //        var options = select.FindElements(By.TagName("option"));
        //        foreach (var op in options)
        //        {
        //            if (!op.Selected) continue;
        //            text = op.Text;
        //            break;
        //        }
        //    }
        //    else if (hasOptionSet)
        //    {
        //        var select = fieldContainer.FindElement(By.XPath(xpathDropOptionSet.Replace("[NAME]", option.Name)));
        //        var options = select.FindElements(By.TagName("option"));
        //        foreach (var op in options)
        //        {
        //            if (!op.Selected) continue;
        //            text = op.Text;
        //            break;
        //        }
        //    }
        //    else if (fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusCombo].Replace("[NAME]", option.Name))).Count > 0)
        //    {
        //        // This is for statuscode (type = status) that should act like an optionset doesn't doesn't follow the same pattern when rendered
        //        var valueSpan = client.Browser.Driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusTextValue].Replace("[NAME]", option.Name)));

        //        text = valueSpan.Text;
        //    }
        //    else
        //    {
        //        throw new InvalidOperationException($"Field: {option.Name} Does not exist");
        //    }
        //    return text;
        //}

        //public static string GetValue(XrmApp xrm, WebClient client, string field, bool debug = false)
        //{
        //    //Sprint 61 - Changed argument so that new simplified hm.GetValue(TextFieldItem), see next method below, is now called by default.
        //    string text = string.Empty;
        //    var fieldContainer = client.Browser.Driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", field.Name)), "content panel not rendering...");

        //    string xpathDrop = "//input[contains(@data-id,'[NAME].fieldControl-text-box-text')]";
        //    string decimalXpathDrop = "//input[contains(@data-id,'[NAME].fieldControl-decimal-number-text-input')]";
        //    string currencyXpathDrop = "//input[@data-id='[NAME].fieldControl-currency-text-input']";
        //    string floatingPointXpathDrop = "//input[@data-id='[NAME].fieldControl-floating-point-text-input']";
        //    string pcfControlXpathDrop = "//div[@data-id='[NAME].fieldControl_container']//input";
        //    string phoneXpathDrop = "//input[contains(@data-id,'[NAME].fieldControl-phone-text-input')]";
        //    string durationXpathDrop = "//input[contains(@data-id,'[NAME].fieldControl-duration-combobox-text')]";
        //    string wholeXpathDrop = "//input[contains(@data-id,'[NAME].fieldControl-whole-number-text-input')]";

        //    // Determine what type of text box the user is dealing with
        //    var isDefaultTextBox = client.Browser.Driver.HasElement(By.XPath(xpathDrop.Replace("[NAME]", field.Name)));
        //    var isDecimalTextBox = client.Browser.Driver.HasElement(By.XPath(decimalXpathDrop.Replace("[NAME]", field.Name)));
        //    var isCurrencyTextBox = client.Browser.Driver.HasElement(By.XPath(currencyXpathDrop.Replace("[NAME]", field.Name)));
        //    var isFloatingPointTextbox = client.Browser.Driver.HasElement(By.XPath(floatingPointXpathDrop.Replace("[NAME]", field.Name)));
        //    var isPCFControlTextBox = client.Browser.Driver.HasElement(By.XPath(pcfControlXpathDrop.Replace("[NAME]", field.Name)));
        //    var isPhoneTextBox = client.Browser.Driver.HasElement(By.XPath(phoneXpathDrop.Replace("[NAME]", field.Name)));
        //    var isDurationTextBox = client.Browser.Driver.HasElement(By.XPath(durationXpathDrop.Replace("[NAME]", field.Name)));
        //    var isWholeTextBox = client.Browser.Driver.HasElement(By.XPath(wholeXpathDrop.Replace("[NAME]", field.Name)));


        //    if (isDefaultTextBox)
        //    {
        //        var input = fieldContainer.FindElement(By.TagName("input"));
        //        if (input != null)
        //        {
        //            IWebElement fieldValue = input.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldValue].Replace("[NAME]", field.Name)));
        //            text = fieldValue.GetAttribute("value").ToString();

        //            // Needed if getting a date field which also displays time as there isn't a date specifc GetValue method
        //            var timefields = client.Browser.Driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.FieldControlDateTimeTimeInputUCI].Replace("[FIELD]", field.Name)));
        //            if (timefields.Any())
        //            {
        //                text = $" {timefields.FirstOrDefault().GetAttribute("value")}";
        //            }
        //        }
        //    }
        //    else if (fieldContainer.FindElements(By.TagName("textarea")).Count > 0)
        //    {
        //        text = fieldContainer.FindElement(By.TagName("textarea")).GetAttribute("value");
        //    }
        //    else if (isDecimalTextBox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(decimalXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else if (isCurrencyTextBox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(currencyXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else if (isFloatingPointTextbox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(floatingPointXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else if (isPCFControlTextBox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(pcfControlXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else if (isPhoneTextBox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(phoneXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else if (isDurationTextBox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(durationXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else if (isWholeTextBox)
        //    {
        //        var fieldElement = client.Browser.Driver.FindElements(By.XPath(wholeXpathDrop.Replace("[NAME]", field.Name)));
        //        text = fieldElement.FirstOrDefault().GetAttribute("value");
        //    }
        //    else
        //    {
        //        throw new Exception($"Field with name {field.Name} does not exist.");
        //    }
        //    return text;
        //}


        ////public static string GetValue(XrmApp xrm, WebClient client, TextFieldItem control)
        ////{
        ////    //Sprint 61 - New simplified TextFieldItem hm.GetValue()
        ////    //  Replaces the one above as, as this implementation improves maintenance and application.

        ////    IWebElement valueContainer = null;  //Used to hold the object we wil be setting the value on

        ////    //NOTE:  In SO, the use of a pcf control requires us to look for that, in addition to the standard text-input,
        ////    //          Refer to SO test case #35988 for ability to see hm.SetValue() and add hm.GetValue() in action on a pcf control.
        ////    string textAreaXpathDrop = $"//textarea[contains(@data-id, '{control.Name}.fieldControl')]";
        ////    string pcfControlXpathDrop = $"//div[@data-id='{control.Name}.fieldControl_container']//input";
        ////    string decimalXpathDrop = $"//input[contains(@data-id,'{control.Name}.fieldControl-decimal-number-text-input')]";
        ////    string currencyXpathDrop = $"//input[@data-id='{control.Name}.fieldControl-currency-text-input']";
        ////    string floatingPointXpathDrop = $"//input[@data-id='{control.Name}.fieldControl-floating-point-text-input']";
        ////    string phoneXpathDrop = $"//input[contains(@data-id,'{control.Name}.fieldControl-phone-text-input')]";
        ////    string durationXpathDrop = $"//input[contains(@data-id,'{control.Name}.fieldControl-duration-combobox-text')]";
        ////    string wholeXpathDrop = $"//input[contains(@data-id,'{control.Name}.fieldControl-whole-number-text-input')]";
        ////    string textInputXpathDrop = $"//input[contains(@data-id,'{control.Name}.fieldControl-text-box-text')]";
        ////    string emailInputXpathDrop = $"//input[contains(@data-id,'{control.Name}.fieldControl-mail-text-input')]";

        ////    bool isPCFControlTextBox = client.Browser.Driver.HasElement(By.XPath(pcfControlXpathDrop));
        ////    bool isTextAreaTextBox = client.Browser.Driver.HasElement(By.XPath(textAreaXpathDrop));
        ////    bool isDecimalTextBox = client.Browser.Driver.HasElement(By.XPath(decimalXpathDrop));
        ////    bool isCurrencyTextBox = client.Browser.Driver.HasElement(By.XPath(currencyXpathDrop));
        ////    bool isFloatingPointTextbox = client.Browser.Driver.HasElement(By.XPath(floatingPointXpathDrop));
        ////    bool isPhoneTextBox = client.Browser.Driver.HasElement(By.XPath(phoneXpathDrop));
        ////    bool isDurationTextBox = client.Browser.Driver.HasElement(By.XPath(durationXpathDrop));
        ////    bool isWholeTextBox = client.Browser.Driver.HasElement(By.XPath(wholeXpathDrop));
        ////    bool isTextInputTextBox = client.Browser.Driver.HasElement(By.XPath(textInputXpathDrop));
        ////    bool isMailInputTextBox = client.Browser.Driver.HasElement(By.XPath(emailInputXpathDrop));

        ////    if (isPCFControlTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(pcfControlXpathDrop));
        ////    else if (isTextAreaTextBox)
        ////    {
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(textAreaXpathDrop));
        ////        return valueContainer.Text;
        ////    }
        ////    else if (isDecimalTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(decimalXpathDrop));
        ////    else if (isCurrencyTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(currencyXpathDrop));
        ////    else if (isFloatingPointTextbox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(floatingPointXpathDrop));
        ////    else if (isPhoneTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(phoneXpathDrop));
        ////    else if (isDurationTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(durationXpathDrop));
        ////    else if (isWholeTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(wholeXpathDrop));
        ////    else if (isTextInputTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(textInputXpathDrop));
        ////    else if (isMailInputTextBox)
        ////        valueContainer = client.Browser.Driver.FindVisible(By.XPath(emailInputXpathDrop));
        ////    else
        ////        throw new InvalidOperationException($"TextFieldItem: {control.Name} Does not exist");

        ////    return valueContainer.GetAttribute<string>("value");
        ////}

        //public static string GetValue(XrmApp xrm, WebClient client, LookupItem field)
        //{

        //    string text = string.Empty;
        //    var emptyFieldContainerXpath = "//input[contains(@data-id,'[NAME].fieldControl-LookupResultsDropdown_[NAME]_textInputBox_with_filter_new') and not(ancestor::div[(contains(@data-id,'QuickFormContainer'))])]";
        //    var populatedFieldContainerXpath = "//div[contains(@data-id,'[NAME].fieldControl-LookupResultsDropdown_[NAME]_selected_tag_text')]";

        //    var hasEmptyFieldContainer = client.Browser.Driver.HasElement(By.XPath(emptyFieldContainerXpath.Replace("[NAME]", field.Name)));
        //    var hasPopulatedFieldContainer = client.Browser.Driver.HasElement(By.XPath(populatedFieldContainerXpath.Replace("[NAME]", field.Name)));

        //    if (hasEmptyFieldContainer)
        //    {
        //        var emptyFieldContainer = client.Browser.Driver.FindElement(By.XPath(emptyFieldContainerXpath.Replace("[NAME]", field.Name)));
        //        text = emptyFieldContainer.GetAttribute("value");
        //    }
        //    else if (hasPopulatedFieldContainer)
        //    {
        //        var populatedFieldContainer = client.Browser.Driver.FindAvailable(By.XPath(populatedFieldContainerXpath.Replace("[NAME]", field.Name)));
        //        text = populatedFieldContainer.Text;
        //    }
        //    else
        //    {
        //        throw new Exception($"Field with name {field.Name} does not exist.");
        //    }
        //    return text;
        //}

        //public static bool GetValue(XrmApp xrm, WebClient client, BooleanItem option)
        //{

        //    var check = false;

        //    var fieldContainer = client.Browser.Driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));
        //    var flipswitchXpathDrop = "//div[@data-id='[NAME].fieldControl_container']//div[contains(@class,'flipswitch')]//a";
        //    string toggleXpath = "//div[@data-id='[NAME].fieldControl-toggle-container']//button[@role='switch']";

        //    var hasRadio = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioContainer].Replace("[NAME]", option.Name)));
        //    var hasCheckbox = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckbox].Replace("[NAME]", option.Name)));
        //    var hasList = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldList].Replace("[NAME]", option.Name)));
        //    var emptyList = fieldContainer.HasElement(By.XPath("//input[contains(@data-id,'[NAME].fieldControl-checkbox-empty')]".Replace("[NAME]", option.Name)));
        //    var newList = fieldContainer.HasElement(By.XPath("//div[contains(@data-id,'[NAME].fieldControl-checkbox-containercheckbox')]".Replace("[NAME]", option.Name)));
        //    var hasFlipswitch = client.Browser.Driver.HasElement(By.XPath(flipswitchXpathDrop.Replace("[NAME]", option.Name)));
        //    bool hasToggle = client.Browser.Driver.HasElement(By.XPath(toggleXpath.Replace("[NAME]", option.Name)));

        //    if (hasRadio)
        //    {
        //        var trueRadio = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioTrue].Replace("[NAME]", option.Name)));

        //        check = bool.Parse(trueRadio.GetAttribute("aria-checked"));
        //    }
        //    else if (hasCheckbox)
        //    {
        //        var checkbox = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckbox].Replace("[NAME]", option.Name)));

        //        check = bool.Parse(checkbox.GetAttribute("aria-checked"));
        //    }
        //    else if (hasList)
        //    {
        //        var list = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldList].Replace("[NAME]", option.Name)));
        //        var options = list.FindElements(By.TagName("option"));
        //        var selectedOption = options.FirstOrDefault(a => a.HasAttribute("data-selected") && bool.Parse(a.GetAttribute("data-selected")));

        //        if (selectedOption != null)
        //        {
        //            check = int.Parse(selectedOption.GetAttribute("value")) == 1;
        //        }
        //    }
        //    else if (emptyList)
        //    {
        //        check = false;
        //    }
        //    else if (newList)
        //    {
        //        var selectedOption = client.Browser.Driver.FindElements(By.XPath("//div[contains(@data-id,'[NAME].fieldControl-checkbox-containercheckbox')]".Replace("[NAME]", option.Name)));
        //        if (selectedOption.Count > 0)
        //        {
        //            check = bool.Parse(selectedOption[0].GetAttribute("aria-checked"));
        //        }
        //    }
        //    else if (hasFlipswitch)
        //    {
        //        var selectedOption = client.Browser.Driver.FindElements(By.XPath(flipswitchXpathDrop.Replace("[NAME]", option.Name)));
        //        if (selectedOption.Count > 0) check = bool.Parse(selectedOption[0].GetAttribute("aria-checked"));
        //    }
        //    else if (hasToggle)
        //    {
        //        check = bool.Parse(client.Browser.Driver.FindElements(By.XPath(toggleXpath.Replace("[NAME]", option.Name))).FirstOrDefault().GetAttribute<string>("aria-checked"));
        //    }
        //    else
        //        throw new InvalidOperationException($"Field: {option.Name} Does not exist");

        //    return check;
        //}
        #endregion

        #region Views Filter
        public class GridHeader
        {
            public int Index { get; set; }
            public string LogicalName { get; set; }
            public string DisplayName { get; set; }
        }

        public static List<GridHeader> GetGridHeaders(XrmApp xrm, WebClient client)
        {

           // IJavaScriptExecutor js = (IJavaScriptExecutor)client.Browser.Driver;

            string gridHeaderXPath = $"{GRIDXPATH}  //*[@role='columnheader']/..";
            string gridCellXPath = $"//*[@role='columnheader']";

            client.Browser.Driver.WaitUntilAvailable(By.XPath(GRIDXPATH));
            var headerCount = int.Parse(client.Browser.Driver.FindElement(By.XPath(GRIDXPATH + "//*[@aria-colcount]")).GetAttribute("aria-colcount"));

            List<GridHeader> result = new List<GridHeader>();

            //js.ExecuteScript(GRID_SCROLL_ALL_RIGHT_COMMAND);
            //while (Convert.ToDouble(js.ExecuteScript(GRID_QUERY_HOW_MUCH_UNTIL_RESET)) != 0)
            //{

                var headerRow = client.Browser.Driver.FindElement(By.XPath(gridHeaderXPath));
                var headerRowHtml = headerRow.GetAttribute("innerHTML");

                var headerDoc = new HtmlDocument();
                headerDoc.LoadHtml(headerRowHtml);

                var headers = headerDoc.DocumentNode.SelectNodes(gridCellXPath);

                foreach (var header in headers)
                {
                    var index = int.Parse(header.GetAttributeValue("aria-colindex", null));
                    var logicalName = header.GetAttributeValue("col-id", null);
                    var displayText = header.GetAttributeValue("title", null);

                    //PCF grids have title in a child node
                    if (displayText == null)
                    {
                        var rowDoc = new HtmlDocument();
                        rowDoc.LoadHtml(header.InnerHtml);

                        displayText = rowDoc.DocumentNode.SelectSingleNode("//*[@title]")?.InnerText;
                    }


                    if (!result.Exists(it => it.Index == index))
                    {
                        result.Add(new GridHeader { Index = index, LogicalName = logicalName, DisplayName = displayText });
                    }
                }

            //    js.ExecuteScript(GRID_SCROLL_LEFT_COMMAND);
            //}

            return result;

        }

        public static List<string> getGridHeaderList(XrmApp xrm, WebClient client, string GridName, int thinkTime = DefaultThinkTime, bool debug = false)
        {
            var columnNames = GetGridHeaders(xrm, client).Select(it => it.DisplayName).ToList();

            if (debug) HelperMethods.DebugOutputForReferenceListCreation(columnNames);

            columnNames.RemoveAll(string.IsNullOrEmpty);

            return columnNames;
        }
        #endregion

        #region Validate Reg Exprestions
        public static bool ValidatePhoneNumber(string PhNumber)
        {
            if (PhNumber != null)
            {
                return Regex.IsMatch(PhNumber, PhoneNumberFormat);
            }
            else
            {
                return false;
            }
        }

        public static bool ValidateEmail(string Email)
        {
            if (Email != null)
            {
                return Regex.IsMatch(Email, EmailFormat);
            }
            else
            {
                return false;
            }
        }

        public static bool ValidateZIP(string zip)
        {
            if (zip != null)
            {
                return Regex.IsMatch(zip, zip);
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Logger ScreenShots, Word Exports.
        //Take Screen Short and pass the name what you want...!
        public static void ScreenShot(XrmApp xrm, WebClient client, string ScreenShortName, string FolderPath)
        {
            string ScreenName = ScreenShortName.ToString();           
            Screenshot ss = ((ITakesScreenshot)client.Browser.Driver).GetScreenshot();
            ss.SaveAsFile(FolderPath +"\\"+ ScreenName + ".jpg");
        }

        public static void CopyScreenShotsIntoWord(XrmApp xrm, WebClient client, string FolderPath, string MethodName)
        {
            
            // Get a list of all image files in the source folder
            string[] images = Directory.GetFiles(FolderPath, "*.jpg", SearchOption.AllDirectories).OrderByDescending(x => x).ToArray();

            // Start a new instance of Word
            Word.Application wordApp = new Word.Application();             
            
           // Create a new document
            Word.Document wordDoc = wordApp.Documents.Add();    
            
            // Loop through the list of images
            foreach (string image in images)
            {

                // Get the file name of the image 
                string imageName = Path.GetFileNameWithoutExtension(image); 
                // Insert the image into the Word document 
                InlineShape shape = wordDoc.InlineShapes.AddPicture(image);
                // Add a description for the image below the image 

                shape.Range.InsertAfter("\n" + imageName + "\n");
                //wordDoc.Content.InsertAfter("\n" + imageName + "\n");
            }

            // Save the Word document
            wordDoc.SaveAs2(FolderPath + MethodName +".docx");             
            // Close the Word application
            wordApp.Quit();
        }

        #endregion

        #region Edit Filter on Views

        public struct advancedFilterTest
        {
            // TestType -  valid values are "fields", "values", and "dropDowns"
            // We can add more as we see fit.
            //Test for examples, see usage references for details.

            public int Index;
            public string TestType, TestFor;

            public advancedFilterTest(int index, string fieldOrValue, string testFor)
            {
                this.Index = index;
                this.TestType = fieldOrValue;
                this.TestFor = testFor;
            }
        }//afTeststruct

        public static void AdvancedFilter_Open(XrmApp xrm, WebClient client)
        {
            // This method navigates to the Advanced Filter page
            try
            {
                //lets make sure we have the filter available.
                HelperMethods.WaitForElementToBeDisplayed(xrm, client, "//button[@id='open-advanced-filter']");

                var button = client.Browser.Driver.FindElement(By.Id("open-advanced-filter"));
                client.ThinkTime(500); // take a pause to avoid race conditions.
                button.Click();

                HelperMethods.WaitForInvisibilityOfProgressIndicator(xrm, client);
                HelperMethods.WaitForElementToBeDisplayed(xrm, client, "//div[@aria-labelledby='advanced-filters-header']");
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + " HelperMethods.AdvancedFilter_Open() failed. ";
                throw new Exception(msg);
            }
        }

        public static void WaitForElementToBeDisplayed(XrmApp xrm, WebClient client, string locatorString, bool byID, bool byXpath)
        {
            TimeSpan timeout = TimeSpan.FromSeconds(90);
            WebDriverWait wait = new WebDriverWait(client.Browser.Driver, timeout);

            // wait for element to appear
            Func<IWebDriver, bool> elementIsDisplayed =
            d =>
            {
                if (byID)
                {
                    ReadOnlyCollection<IWebElement> e = d.FindElements(By.Id(locatorString));
                    return e.Count > 0;
                }
                else if (byXpath)
                {
                    ReadOnlyCollection<IWebElement> e = d.FindElements(By.XPath(locatorString));
                    return e.Count > 0;
                }
                else
                    throw new Exception("Could not locate element to wait for.");
            };

            wait.Until(elementIsDisplayed);

            // TODO: Refactor to use WaitUntilClickable/etc. at a later date
            xrm.ThinkTime(500);
        }

        public static void WaitForElementToBeDisplayed(XrmApp xrm, WebClient client, string elementXpath)
        {
            TimeSpan timeout = TimeSpan.FromSeconds(60);
            WebDriverWait wait = new WebDriverWait(client.Browser.Driver, timeout);

            // wait for element to appear
            Func<IWebDriver, bool> elementIsDisplayed =
            d =>
            {
                ReadOnlyCollection<IWebElement> e = d.FindElements(By.XPath(elementXpath));
                return e.Count > 0;
            };
            wait.Until(elementIsDisplayed);

            // TODO: Refactor to use WaitUntilClickable/etc. at a later date
            xrm.ThinkTime(500);
        }

        public static List<IWebElement> AdvancedFilter_GetFields(XrmApp xrm, WebClient client)
        {
            // This method gets the left two column fields from the AF dialog, e.g. Status, Equals
            try
            {
                string xpathFields = "//input[contains(@class, 'ms-ComboBox-Input')]";
                var fields = client.Browser.Driver.FindElements(By.XPath(xpathFields)).ToList();
                return fields;
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + " HelperMethods.AdvancedFilter_GetFields() failed. ";
                throw new Exception(msg);
            }
        }

        public static ReadOnlyCollection<IWebElement> AdvancedFilter_GetTagItems(XrmApp xrm, WebClient client)
        {
            // This method gets the right most values from the AF dialog, e.g. Active
            try
            {
                string xpathValues = "//div[contains(@class, 'ms-BasePicker')]//span[contains(@class, 'ms-TagItem-text')]";
                var values = client.Browser.Driver.FindElements(By.XPath(xpathValues));
                return values;
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + " HelperMethods.AdvancedFilter_GetValues() failed. ";
                throw new Exception(msg);
            }
        }

        public static ReadOnlyCollection<IWebElement> AdvancedFilter_GetDropDowns(XrmApp xrm, WebClient client)
        {
            // This method gets the Drop Down from teh AF dialog, e.g. Contains data
            try
            {
                string xpath = "//span[contains(@class, 'ms-Dropdown-title')]";
                var values = client.Browser.Driver.FindElements(By.XPath(xpath));
                return values;
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + " HelperMethods.AdvancedFilter_GetValues() failed. ";
                throw new Exception(msg);
            }
        }

        public static List<IWebElement> AdvancedFilter_GetFreeTextFields(XrmApp xrm, WebClient client)
        {
            // This method gets the right most values from the AF dialog, e.g. Active
            try
            {
                string xpathValues = "//input[contains(@class,'ms-TextField-field')]";
                var values = client.Browser.Driver.FindElements(By.XPath(xpathValues)).ToList();
                return values;
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + " HelperMethods.AdvancedFilter_GetFreeTextFields() failed. ";
                throw new Exception(msg);
            }
        }

        public static bool AdvancedFilter_TestFor(XrmApp xrm, advancedFilterTest test, List<IWebElement> fields, ReadOnlyCollection<IWebElement> values, ReadOnlyCollection<IWebElement> dropDowns, List<IWebElement> freeTextField = null)
        {
            // This method takes in an advanced filterstruct and performas the test
            try
            {
                if (test.TestType == "fields")
                {
                    return (fields[test.Index].GetAttribute("value") == test.TestFor);
                }
                else if (test.TestType == "values")
                {
                    return (values[test.Index].Text == test.TestFor);
                }
                else if (test.TestType == "dropDowns")
                {
                    return (dropDowns[test.Index].Text == test.TestFor);
                }
                else if (test.TestType == "FreeTextValue")
                {
                    return (freeTextField[test.Index].GetAttribute("value") == test.TestFor);
                }
                else
                {
                    //Incorrect TestType passed in... fail TC
                    return false;
                }
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + $" HelperMethods.AdvancedFilter_TestFor() failed for test: >> {test} << ";
                throw new Exception(msg);
            }
        }

        public static void AdvancedFilter_Cancel(XrmApp xrm, WebClient client)
        {
            try
            {
                //lets make sure we have the cancel available.
                HelperMethods.WaitForElementToBeDisplayed(xrm, client, "//button[@title='Cancel']");

                var button = client.Browser.Driver.FindElement(By.XPath("//button[@title='Cancel']"));
                client.ThinkTime(500); // take a pause to avoid race conditions.
                button.Click();

                HelperMethods.WaitForInvisibilityOfProgressIndicator(xrm, client);
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + " HelperMethods.AdvancedFilter_Cancel() failed. ";
                throw new Exception(msg);
            }
        }
        #endregion

        public static List<string> GetGridViewList(XrmApp xrm, WebClient client, bool debug = false)
        {
            string xpathViewSelector = "//button[contains(@id,'ViewSelector')]//i[@data-icon-name='ChevronDown']";
            //string homepageGridXpath = "//div[@data-id='GridRoot']";
            //string xpathViewList = "//div[@aria-label='Views']//li[contains(@class,'ContextualMenu-item')]//button[@role='menuitemradio']//span[contains(@class,'itemText')]";
            string xpathViewList = "//div[@aria-label='Views']//li[contains(@class,'ContextualMenu-item')]//button[@role='menuitemradio']";
            List<string> viewList = new List<string>();

            //client.Browser.Driver.WaitUntilVisible(By.XPath(homepageGridXpath));  //Unnecessary code.

            var viewSelector = client.Browser.Driver.ClickIfVisible(By.XPath(xpathViewSelector), new TimeSpan(0, 0, 2));
            client.Browser.Driver.WaitForTransaction();

            client.Browser.Driver.WaitUntilVisible(By.XPath(xpathViewList));

            var dropList = client.Browser.Driver.FindElements(By.XPath(xpathViewList));

            //If viewSelector is null then previous xpath no good, change it here.  drop list must be empty as we didn't find a view selector
            if (viewSelector == null)
            {
                xpathViewList = "//ul[contains(@id, 'ViewSelector')]//li";
                xpathViewSelector = "//span[contains(@id,'ViewSelector')]";
            }

            // Let's try to fill our drop list
            if (dropList.Count == 0)
            {
                // Click the new viewSelector so we get the drop list visible.
                viewSelector = client.Browser.Driver.ClickIfVisible(By.XPath(xpathViewSelector), new TimeSpan(0, 0, 2));
                //viewSelector.Click();
                client.Browser.Driver.WaitUntilVisible(By.XPath(xpathViewList));
                dropList = client.Browser.Driver.FindElements(By.XPath(xpathViewList));
            }

            //Populate our return list with drops
            foreach (var drop in dropList)
                viewList.Add(drop.Text);

            //Click the drop list so it doesn't remain empty
            viewSelector.Click();
            client.Browser.Driver.WaitForTransaction();

            if (debug) HelperMethods.DebugOutputForReferenceListCreation(viewList);
            return viewList;
        }

        public static ReadOnlyCollection<IWebElement> FindElementXpath_notificationWrapper (XrmApp xrm, WebClient client)
        {
            string findElementXpath = "//div[@role='presentation']//div[@data-id='notificationWrapper']";

            var isSaved = client.Browser.Driver.FindElements(By.XPath(findElementXpath));

            return isSaved;
        }




    }
}
