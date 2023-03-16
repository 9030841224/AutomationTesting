using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using System.Collections.Generic;
using OpenQA.Selenium;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using System.IO;
using System.Data.OleDb;
using System.Data;
using static Microsoft.Dynamics365.UIAutomation.Api.UCI.HelperMethods;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class Sprint1
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        string parentFolder = "D:\\New folder\\EasyRepro-develop\\TestResults";
        public const string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\\Projects\\Accounts.xlsx; Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'";

        private static ExtentReports extent;

        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            extent = new ExtentReports();
            Random r = new Random();
            var htmlReporter = new ExtentHtmlReporter(@"D:\New folder\EasyRepro-develop\Microsoft.Dynamics365.UIAutomation.Sample\ExtentReports\"+r.Next(10,100)+"index.HTML");
            extent.AttachReporter(htmlReporter);
        }

        [TestMethod]
        public void CreateNewAccount()
        {
            var test = extent.CreateTest("CreateNewAccount");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {                    
                    string childFolder = "CreateNewAccount";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    bool Flag = false;
                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);                    
                    test.Log(Status.Info, Info.login);                    
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);                    
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    test.Log(Status.Info, "Navigating to Account");
                    xrmApp.CommandBar.ClickCommand("New");
                    test.Log(Status.Info, "Create new account");
                    String BusinessName = "Account";
                    HelperMethods.CreateTestAccount(xrmApp, BusinessName);
                    xrmApp.Entity.Save();

                    Flag = true;
                    if (Flag)
                    {
                        test.Log(Status.Info, "Created new account");
                        test.Pass("Test Passed");
                        HelperMethods.ScreenShot(xrmApp, client, Info.createAccount, childFolderPath);
                        HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");

                        //Delete Account start here..xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                        //HelperMethods.Grid_SwitchView(xrmApp, client, "My Active Accounts");
                        //test.Log(Status.Info, "My Active accounts View");
                        //HelperMethods.SelectRecord(xrmApp, client, 0, false);
                        //xrmApp.CommandBar.ClickCommand("Delete");
                        //xrmApp.ThinkTime(9000);
                        //xrmApp.Dialogs.ConfirmationDialog(true);

                        Assert.IsTrue(Flag);
                    }
                    else
                    {
                        test.Fail("CreateNewAccount is Failed");
                        Assert.IsTrue(Flag);
                    }
                }
            }
            catch
            {
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }

        [TestMethod]
        public void AcountTabsVerify()
        {
            var test = extent.CreateTest("AcountTabsVerify");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "AcountTabsVerify";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    test.Log(Status.Info, "Navigating to Account");

                    xrmApp.CommandBar.ClickCommand("New");
                    test.Log(Status.Info, "AccountTab Verify");

                    // Verify Tab List
                    List<string> referenceList = new List<string> {
                                                                        "Summary",
                                                                        "Project Price Lists",
                                                                        "Details",
                                                                        "Servicing",
                                                                        "Files"
                                                                    };
                    List<string> tabList = HelperMethods.GetTabs(xrmApp, client, true);
                    bool result = HelperMethods.CompareStringLists(referenceList, tabList, false, true);
                    test.Log(Status.Info, "Tabs Verified");
                    test.Pass("Test Passed");
                    HelperMethods.ScreenShot(xrmApp, client, Info.tabsVerify, childFolderPath);
                    HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                    Assert.IsTrue(result && referenceList.Count == tabList.Count);
                }
            }
            catch(Exception ex)
            {
                test.Fail(ex);
                Assert.IsTrue(false);
            }
        }

        [TestMethod]
        public void AcountCoreFieldsVerify()
        {
            var test = extent.CreateTest("AcountCoreFieldsVerify");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "AcountCoreFieldsVerify";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    test.Log(Status.Info, "Navigating to Account");
                    xrmApp.CommandBar.ClickCommand("New");
                    test.Log(Status.Info, "AccountCoreFieldsVerify");
                    xrmApp.Entity.SelectTab("Summary");
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    xrmApp.ThinkTime(3000);

                    test.Log(Status.Info, "Core Field Verify List");
                    List<string> referenceListcorefield = new List<string> {
                    "Account Name",
                    "Relationship Type"
                };

                    // NonCore field verified list
                    test.Log(Status.Info, "Non-Core Field Verify List");
                    List<string> referenceListnoncorefield = new List<string> {
                    //"Account Name",
                    "Phone",
                    "Fax",
                    "Website",
                    "Parent Account",
                    "Ticker Symbol",
                    //"Relationship Type",
                    "Product Price List",
                    "Service Address" ,
                    "Primary Contact"
                };
                    var formLabelslocation = HelperMethods.Form_GetLabels(xrmApp, client, true);
                    var allFormLabelsTextlocation = HelperMethods.Labels_CreateList(xrmApp, formLabelslocation);
                    var GetCorefieldLabelsOnForm = HelperMethods.GetCoreField(referenceListnoncorefield, allFormLabelsTextlocation);
                    //Compare against our required Labels Reference list
                    bool FieldsMatch = HelperMethods.CompareStringLists(referenceListcorefield, GetCorefieldLabelsOnForm, true, true);
                    Assert.IsTrue(FieldsMatch);

                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    xrmApp.ThinkTime(2000);
                    test.Pass("AcountCoreFieldsVerified");
                    test.Pass("Test Passed");
                    HelperMethods.ScreenShot(xrmApp, client, Info.acountCoreFieldsVerify, childFolderPath);
                    HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                    Assert.IsTrue(true);
                }
            }
            catch
            {
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }

        [TestMethod]
        public void AcountPhoneNumberValidation()
        {
            var test = extent.CreateTest("AcountPhoneNumberValidation");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "AcountPhoneNumberValidation";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    test.Log(Status.Info, "Navigating to Account");
                    HelperMethods.Grid_SwitchView(xrmApp, client, "My Active Accounts");
                    test.Log(Status.Info, "My Active accounts View");
                    HelperMethods.OpenFirstRecord(xrmApp, client);
                    test.Log(Status.Info, "Open First Record");
                    var phone = xrmApp.Entity.GetValue("telephone1");
                    bool validatePhoneNumber = HelperMethods.ValidatePhoneNumber(phone);
                    if (validatePhoneNumber)
                    {
                        test.Log(Status.Info, "Phone Number Validation done");
                        test.Pass("Test Passed");
                        HelperMethods.ScreenShot(xrmApp, client, Info.phoneNumberValidation, childFolderPath);
                    }
                    else
                    {
                        test.Log(Status.Info, "Phone Number Validation Failed");
                        test.Pass("Test Failed");
                        HelperMethods.ScreenShot(xrmApp, client, "Phone Number Validation Failed", childFolderPath);
                    }                    
                    HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                    Assert.IsTrue(validatePhoneNumber);
                }
            }
            catch(Exception ex)
            {
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }

        [TestMethod]
        public void GlobalSearch()
        {
            var test = extent.CreateTest("GlobalSearch");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "GlobalSearch";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenGlobalSearch();
                    test.Log(Status.Info, "Open Global Search");
                    bool v = xrmApp.GlobalSearch.ChangeSearchType("Categorized Search");
                    xrmApp.GlobalSearch.Search("Test");
                    test.Log(Status.Info, "Search with Test");
                    xrmApp.GlobalSearch.FilterWith("Account");
                    test.Log(Status.Info, "Fliter With Account");
                    xrmApp.GlobalSearch.OpenRecord("account", 0);
                    test.Log(Status.Info, "Open first Record");
                    test.Log(Status.Info, "GlobalSearch Done");
                    test.Pass("Test Passed");
                    HelperMethods.ScreenShot(xrmApp, client, Info.GlobalSearch, childFolderPath);
                    HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                    Assert.IsTrue(true);
                }
            }
            catch
            {
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }

        [TestMethod]
        public void CreateContact()
        {
            var test = extent.CreateTest("CreateNewContact");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "CreateContact";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Contacts");
                    test.Log(Status.Info, "Navigating to Contact");
                    xrmApp.CommandBar.ClickCommand("New");
                    test.Log(Status.Info, "Create new Contact");
                    String BusinessName = "Contact";
                    HelperMethods.CreateContact(xrmApp, BusinessName);
                    xrmApp.Entity.Save();
                    test.Log(Status.Info, "Created new Contact");
                    test.Pass("Test Passed");
                    HelperMethods.ScreenShot(xrmApp, client, Info.CreateContact, childFolderPath);
                    HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                    Assert.IsTrue(true);
                }
            }
            catch
            {
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }
      
        public void DeleteAccount()
        {
            var test = extent.CreateTest("DeleteAccount");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "DeleteAccount";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    test.Log(Status.Info, "Navigating to Account");
                    HelperMethods.Grid_SwitchView(xrmApp, client, "My Active Accounts");
                    test.Log(Status.Info, "My Active accounts View");
                    HelperMethods.OpenFirstRecord(xrmApp, client);
                    test.Log(Status.Info, "Open First Record");
                    xrmApp.Entity.Delete();

                }
            }
            catch
            {

            }
        }

        [TestMethod]
        public void ImportExcelDataForCreatingNewAccounts()
        {
            var test = extent.CreateTest("ImportExcelDataForCreatingNewAccounts");
            InfoConstants Info = new InfoConstants();
            var client = new WebClient(TestSettings.Options);
            bool Flag = false;
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "ImportExcelDataForCreatingNewAccounts";
                    // Check if the parent folder exists
                    if (!Directory.Exists(parentFolder))
                    {
                        // If the parent folder does not exist, create it
                        Directory.CreateDirectory(parentFolder);
                    }
                    // Create the full path to the child folder
                    string childFolderPath = Path.Combine(parentFolder, childFolder);
                    // Create the child folder
                    string path = Directory.CreateDirectory(childFolderPath).ToString();

                    test.Log(Status.Info, "Loading Chrome Browser");
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    test.Log(Status.Info, Info.login);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    test.Log(Status.Info, Info.OpenSales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    test.Log(Status.Info, "Navigating to Account");
                    // Create the connection object
                    using(OleDbConnection oledbConn = new OleDbConnection(connString))
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
                            xrmApp.CommandBar.ClickCommand("New");
                            var AccountName = ((System.Data.DataRowView)m).Row.ItemArray[0];
                            var Phone = ((System.Data.DataRowView)m).Row.ItemArray[1];
                            var Website = ((System.Data.DataRowView)m).Row.ItemArray[2];
                            var RelationshipType = ((System.Data.DataRowView)m).Row.ItemArray[3];
                            var ParentAccount = ((System.Data.DataRowView)m).Row.ItemArray[4];
                            var ProductPriceList = ((System.Data.DataRowView)m).Row.ItemArray[5];

                            xrmApp.Entity.SetValue("name", AccountName.ToString());
                            xrmApp.Entity.SetValue("telephone1", Phone.ToString());
                            xrmApp.Entity.SetValue("websiteurl", Website.ToString());
                            xrmApp.Entity.SetValue(new OptionSet { Name = "customertypecode", Value = RelationshipType.ToString() });
                            xrmApp.Entity.SetValue(new LookupItem { Name = "parentaccountid", Value = ParentAccount.ToString(), Index = 0 });

                            xrmApp.Entity.Save();
                            HelperMethods.ScreenShot(xrmApp, client, Info.createAccount, childFolderPath);
                            xrmApp.ThinkTime(5000);
                        }

                        Flag = true;
                        if (Flag)
                        {
                            test.Log(Status.Info, "Created new accounts");
                            test.Pass("Test Passed");
                            HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");  
                            Assert.IsTrue(Flag);
                        }
                        else
                        {
                            test.Fail("CreateNewAccount is Failed");
                            Assert.IsTrue(Flag);
                        }
                    }                                       
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }        

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            extent.Flush();
        }        
    }
    public class InfoConstants
    {
        public string login = "1.Login into Dynamics365 Apps by using valid credentials.";
        public string OpenSales = "2.Navigated to Sale App in Dynamics365 CRM.";
        public string createAccount = "3.Creating new customers account in Sales app.";
        public string tabsVerify = "3.Tabs verified in account form.";
        public string acountCoreFieldsVerify = "3.Mandatory field are verified in accounts form.";
        public string phoneNumberValidation = "3.Phone number validated in account form";
        public string GlobalSearch = "3.Search input in global search";
        public string CreateContact = "3.Creating new contact in sales app.";

       



    }
}