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
    public class Sprint1 : BaseClass
    {       
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
                    #region Screenshot Folder Creation 

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

                    #endregion

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

                    var isSaved = HelperMethods.FindElementXpath_notificationWrapper(xrmApp, client);

                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);

                    if (isSaved == null || isSaved.Count == 0)
                    {
                        test.Log(Status.Info, "Created new account");
                        test.Pass("Test Passed");
                        HelperMethods.ScreenShot(xrmApp, client, Info.createAccount, childFolderPath);
                        HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                        Assert.IsTrue(true);
                    }
                    else
                    {
                        test.Fail("CreateNewAccount is Failed");
                        HelperMethods.ScreenShot(xrmApp, client, "new customers account in Sales app Failed", childFolderPath);
                        HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                        Assert.IsTrue(false);
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
                    //xrmApp.ThinkTime(3000);

                    test.Log(Status.Info, "Core Field Verify List");
                    List<string> referenceListcorefield = new List<string> {
                    "Account Name"
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
                    "Relationship Type",
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

        [TestMethod]
        public void Validate_Account_Views()
        {
            var test = extent.CreateTest("Validate_Account_Views");
            var client = new WebClient(TestSettings.Options);
            bool Flag = false;
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "Validate_Account_Views";
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

                    #region Accounts Being Followed

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Accounts Being Followed");
                    test.Log(Status.Info, "Checking Accounts Being Followed View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableAccountsBeingFollowed = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Primary Contact",
                     "Email (Primary Contact)",
                     "Status"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsAccountsBeingFollowed = HelperMethods.getGridHeaderList(xrmApp, client, "Accounts Being Followed");
                    HelperMethods.ScreenShot(xrmApp, client, "Accounts Being Followed View", childFolderPath);

                    var resultAccountsBeingFollowed = HelperMethods.CompareStringLists(viewColumnsAccountsBeingFollowed, referenceListViewColumnLableAccountsBeingFollowed, false, true);
                    Assert.IsTrue(resultAccountsBeingFollowed && viewColumnsAccountsBeingFollowed.Count == referenceListViewColumnLableAccountsBeingFollowed.Count);
                    #endregion

                    #region Accounts I Follow

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Accounts I Follow");
                    test.Log(Status.Info, "Checking Accounts I Follow View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableAccountsIFollow = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Primary Contact",
                     "Email (Primary Contact)",
                     "Status"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsAccountsIFollow = HelperMethods.getGridHeaderList(xrmApp, client, "Accounts I Follow");
                    HelperMethods.ScreenShot(xrmApp, client, "Accounts I Follow View", childFolderPath);

                    var resultAccountsIFollow = HelperMethods.CompareStringLists(viewColumnsAccountsIFollow, referenceListViewColumnLableAccountsIFollow, false, true);
                    Assert.IsTrue(resultAccountsIFollow && viewColumnsAccountsIFollow.Count == referenceListViewColumnLableAccountsIFollow.Count);
                    #endregion

                    #region Accounts: Influenced Deals That We Won*

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Accounts: Influenced Deals That We Won");
                    test.Log(Status.Info, "Checking Accounts: Influenced Deals That We Won View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableAccountsInfluencedDeals = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Address 1: City",
                     "Primary Contact",
                     "Email (Primary Contact)"
                     //"Status"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsAccountsInfluencedDeals = HelperMethods.getGridHeaderList(xrmApp, client, "Accounts: Influenced Deals That We Won");
                    HelperMethods.ScreenShot(xrmApp, client, "Accounts Influenced Deals That We Won View", childFolderPath);

                    var resultAccountsInfluencedDeals = HelperMethods.CompareStringLists(viewColumnsAccountsInfluencedDeals, referenceListViewColumnLableAccountsInfluencedDeals, false, true);
                    Assert.IsTrue(resultAccountsInfluencedDeals && viewColumnsAccountsInfluencedDeals.Count == referenceListViewColumnLableAccountsInfluencedDeals.Count);
                    #endregion

                    #region Accounts: No Campaign Activities in Last 3 Months

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Accounts: No Campaign Activities in Last 3 Months");
                    test.Log(Status.Info, "Checking Accounts: No Campaign Activities in Last 3 Months View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableAccountsNoCampaignActivitiesinLast3Months = new List<string> {
                    "Account Name",
                    "Last Date Included in Campaign",
                     "Main Phone",
                     "Address 1: City"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsAccountsNoCampaignActivitiesinLast3Months = HelperMethods.getGridHeaderList(xrmApp, client, "Accounts: No Campaign Activities in Last 3 Months");
                    HelperMethods.ScreenShot(xrmApp, client, "Accounts No Campaign Activities in Last 3 Months View", childFolderPath);

                    var resultAccountsNoCampaignActivitiesinLast3Months = HelperMethods.CompareStringLists(viewColumnsAccountsNoCampaignActivitiesinLast3Months, referenceListViewColumnLableAccountsNoCampaignActivitiesinLast3Months, false, true);
                    Assert.IsTrue(resultAccountsNoCampaignActivitiesinLast3Months && viewColumnsAccountsNoCampaignActivitiesinLast3Months.Count == referenceListViewColumnLableAccountsNoCampaignActivitiesinLast3Months.Count);
                    #endregion

                    #region Accounts: Responded to Campaigns in Last 6 Months

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Accounts: Responded to Campaigns in Last 6 Months");
                    test.Log(Status.Info, "Checking Accounts: Responded to Campaigns in Last 6 Months View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableAccountsRespondedtoCampaignsinLast6Months = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Address 1: City"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsAccountsRespondedtoCampaignsinLast6Months = HelperMethods.getGridHeaderList(xrmApp, client, "Accounts: Responded to Campaigns in Last 6 Months");
                    HelperMethods.ScreenShot(xrmApp, client, "Accounts Responded to Campaigns in Last 6 Months View", childFolderPath);

                    var resultAccountsRespondedtoCampaignsinLast6Months = HelperMethods.CompareStringLists(viewColumnsAccountsRespondedtoCampaignsinLast6Months, referenceListViewColumnLableAccountsRespondedtoCampaignsinLast6Months, false, true);
                    Assert.IsTrue(resultAccountsRespondedtoCampaignsinLast6Months && viewColumnsAccountsRespondedtoCampaignsinLast6Months.Count == referenceListViewColumnLableAccountsRespondedtoCampaignsinLast6Months.Count);
                    #endregion

                    #region Active Accounts

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Active Accounts");
                    test.Log(Status.Info, "Checking Active Accounts View");
                    xrmApp.ThinkTime(2000);

                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableActiveAccounts = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Address 1: City",
                     "Primary Contact",
                     "Email (Primary Contact)"
                    };

                    //Get View Columns Lable List
                    List<string> viewColumnsActiveAccounts = HelperMethods.getGridHeaderList(xrmApp, client, "Active Accounts");
                    HelperMethods.ScreenShot(xrmApp, client, "Active Accounts View", childFolderPath);

                    var resultActiveAccounts = HelperMethods.CompareStringLists(viewColumnsActiveAccounts, referenceListViewColumnLableActiveAccounts, false, true);
                    Assert.IsTrue(resultActiveAccounts && viewColumnsActiveAccounts.Count == referenceListViewColumnLableActiveAccounts.Count);

                    #endregion

                    #region All Accounts

                    HelperMethods.Grid_SwitchView(xrmApp, client, "All Accounts");
                    test.Log(Status.Info, "Checking All Accounts View");
                    xrmApp.ThinkTime(2000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableAllAccounts = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Primary Contact",
                     "Email (Primary Contact)",
                     "Address 1: City",
                     "Status"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsAllAccounts = HelperMethods.getGridHeaderList(xrmApp, client, "All Accounts");
                    HelperMethods.ScreenShot(xrmApp, client, "All Accounts View", childFolderPath);
                    var resultActiveAllAccounts = HelperMethods.CompareStringLists(viewColumnsAllAccounts, referenceListViewColumnLableAllAccounts, false, true);
                    Assert.IsTrue(resultActiveAllAccounts && viewColumnsAllAccounts.Count == referenceListViewColumnLableAllAccounts.Count);
                    #endregion                    

                    #region Customers
                    HelperMethods.Grid_SwitchView(xrmApp, client, "Customers");
                    test.Log(Status.Info, "Checking Customers View");
                    xrmApp.ThinkTime(2000);

                    List<string> referenceListViewColumnLableCustomers = new List<string>
                    {
                     "Account Name",
                     "Product Price List",
                     "Main Phone",
                     "Address 1: City",
                     "Primary Contact",
                     "Email (Primary Contact)"
                    };

                    List<string> viewColumnsCustomers = HelperMethods.getGridHeaderList(xrmApp, client, "Customers");
                    HelperMethods.ScreenShot(xrmApp, client, "Customers View", childFolderPath);
                    var resultActiveCustomers = HelperMethods.CompareStringLists(viewColumnsCustomers, referenceListViewColumnLableCustomers, false, true);

                    Assert.IsTrue(resultActiveCustomers && viewColumnsCustomers.Count == referenceListViewColumnLableCustomers.Count);
                    #endregion

                    #region Excluded Accounts Campaigns

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Excluded Accounts Campaigns");
                    test.Log(Status.Info, "Checking Excluded Accounts Campaigns View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableExcludedAccountsCampaigns = new List<string> {
                    "Account Name",
                     "Main Phone"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsExcludedAccountsCampaigns = HelperMethods.getGridHeaderList(xrmApp, client, "Excluded Accounts Campaigns");
                    HelperMethods.ScreenShot(xrmApp, client, "Excluded Accounts Campaigns View", childFolderPath);

                    var resultExcludedAccountsCampaigns = HelperMethods.CompareStringLists(viewColumnsExcludedAccountsCampaigns, referenceListViewColumnLableExcludedAccountsCampaigns, false, true);
                    Assert.IsTrue(resultExcludedAccountsCampaigns && viewColumnsExcludedAccountsCampaigns.Count == referenceListViewColumnLableExcludedAccountsCampaigns.Count);
                    #endregion

                    #region Inactive Accounts

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Inactive Accounts");
                    test.Log(Status.Info, "Checking Inactive accounts View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableInactiveAccounts = new List<string> {
                    "Account Name",
                    "Primary Contact",
                     "Main Phone",
                     "Address 1: City",
                     "Email (Primary Contact)"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsInactiveAccounts = HelperMethods.getGridHeaderList(xrmApp, client, "Inactive Accounts");
                    HelperMethods.ScreenShot(xrmApp, client, "Inactive accounts View", childFolderPath);

                    var resultInactiveAccounts = HelperMethods.CompareStringLists(viewColumnsInactiveAccounts, referenceListViewColumnLableInactiveAccounts, false, true);
                    Assert.IsTrue(resultInactiveAccounts && viewColumnsInactiveAccounts.Count == referenceListViewColumnLableInactiveAccounts.Count);
                    #endregion

                    #region My Active Accounts Default

                    HelperMethods.Grid_SwitchView(xrmApp, client, "My Active Accounts");
                    test.Log(Status.Info, "Checking My Active accounts View");
                    xrmApp.ThinkTime(2000);

                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableMyActiveAccountsDefault = new List<string>
                    {
                    "Account Name",
                    "Primary Contact",
                     "Main Phone",
                     "Address 1: City",
                     "Email (Primary Contact)"
                    };

                    //Get View Columns Lable List
                    List<string> viewColumnsMyActiveAccountDefault = HelperMethods.getGridHeaderList(xrmApp, client, "My Active Accounts");
                    HelperMethods.ScreenShot(xrmApp, client, "My Active Accounts View", childFolderPath);

                    var resultActiveMyActiveAccountsDefault = HelperMethods.CompareStringLists(viewColumnsMyActiveAccountDefault, referenceListViewColumnLableMyActiveAccountsDefault, false, true);
                    Assert.IsTrue(resultActiveMyActiveAccountsDefault && viewColumnsMyActiveAccountDefault.Count == referenceListViewColumnLableMyActiveAccountsDefault.Count);

                    #endregion

                    #region My Connections

                    HelperMethods.Grid_SwitchView(xrmApp, client, "My Connections");
                    test.Log(Status.Info, "Checking My Connections View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableMyConnections = new List<string> {
                    "Account Name",
                     "Main Phone",
                     "Address 1: City",
                     "Primary Contact",
                     "Email (Primary Contact)"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsMyConnections = HelperMethods.getGridHeaderList(xrmApp, client, "My Connections");
                    HelperMethods.ScreenShot(xrmApp, client, "My Connections View", childFolderPath);

                    var resultMyConnections = HelperMethods.CompareStringLists(viewColumnsMyConnections, referenceListViewColumnLableMyConnections, false, true);
                    Assert.IsTrue(resultMyConnections && viewColumnsMyConnections.Count == referenceListViewColumnLableMyConnections.Count);
                    #endregion

                    #region Selected Accounts Campaigns

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Selected Accounts Campaigns");
                    test.Log(Status.Info, "Checking Selected Accounts Campaigns View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableSelectedAccountsCampaigns = new List<string> {
                    "Account Name",
                     "Main Phone"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsSelectedAccountsCampaigns = HelperMethods.getGridHeaderList(xrmApp, client, "Selected Accounts Campaigns");
                    HelperMethods.ScreenShot(xrmApp, client, "Selected Accounts Campaigns View", childFolderPath);

                    var resultSelectedAccountsCampaigns = HelperMethods.CompareStringLists(viewColumnsSelectedAccountsCampaigns, referenceListViewColumnLableSelectedAccountsCampaigns, false, true);
                    Assert.IsTrue(resultSelectedAccountsCampaigns && viewColumnsSelectedAccountsCampaigns.Count == referenceListViewColumnLableSelectedAccountsCampaigns.Count);
                    #endregion

                    #region Service Account

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Service Account");
                    test.Log(Status.Info, "Checking Service Account View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableServiceAccount = new List<string> {
                    "Account Name",
                    "Email",
                     "Main Phone",
                     "Account Number",
                     "Primary Contact",
                     "Address 1: City"
                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsServiceAccount = HelperMethods.getGridHeaderList(xrmApp, client, "Service Account");
                    HelperMethods.ScreenShot(xrmApp, client, "Service Account view", childFolderPath);

                    var resultServiceAccount = HelperMethods.CompareStringLists(viewColumnsServiceAccount, referenceListViewColumnLableServiceAccount, false, true);
                    Assert.IsTrue(resultServiceAccount && viewColumnsServiceAccount.Count == referenceListViewColumnLableServiceAccount.Count);

                    #endregion

                    #region Vendors

                    HelperMethods.Grid_SwitchView(xrmApp, client, "Vendors");
                    test.Log(Status.Info, "Checking Vendors View");
                    xrmApp.ThinkTime(1000);
                    // View Column Lable verified list
                    List<string> referenceListViewColumnLableVendors = new List<string> {
                    "Account Name",
                    "Product Price List",
                     "Main Phone",
                     "Address 1: City",
                     "Primary Contact",
                     "Email (Primary Contact)"

                    };
                    //Get View Columns Lable List
                    List<string> viewColumnsVendors = HelperMethods.getGridHeaderList(xrmApp, client, "Vendors");
                    HelperMethods.ScreenShot(xrmApp, client, "Vendors view", childFolderPath);

                    var resultVendors = HelperMethods.CompareStringLists(viewColumnsVendors, referenceListViewColumnLableVendors, false, true);
                    Assert.IsTrue(resultVendors && viewColumnsVendors.Count == referenceListViewColumnLableVendors.Count);
                    #endregion

                    Flag = true;
                    if (Flag)
                    {
                        test.Log(Status.Info, "All Views  Validated Successfully ");
                        test.Pass("Test Passed");
                        HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "ValidateViews");
                        Assert.IsTrue(Flag);
                    }
                    else
                    {
                        test.Fail("Test Failed");
                        Assert.IsTrue(Flag);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }

               
    }
    public class InfoConstants
    {
        public string login = "1.Login into Dynamics365 Apps by using valid credentials.";
        public string OpenSales = "2.Navigated to Sale App in Dynamics365 CRM.";
        public string createAccount = "3.Created new customers account in Sales app is Successfully.";
        public string tabsVerify = "3.Tabs verified in account form.";
        public string acountCoreFieldsVerify = "3.Mandatory field are verified in accounts form.";
        public string phoneNumberValidation = "3.Phone number validated in account form";
        public string GlobalSearch = "3.Search input in global search";
        public string CreateContact = "3.Creating new contact in sales app.";

       



    }
}