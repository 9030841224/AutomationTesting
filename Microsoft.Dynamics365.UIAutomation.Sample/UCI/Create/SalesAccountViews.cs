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
    /// <summary>
    /// Summary description for SalesAccountViews
    /// </summary>
    [TestClass]
    public class SalesAccountViews : BaseClass
    {        
        [TestMethod]
        public void Validate_Account_Views_Filter_Smoke_Test()
        {
            var test = extent.CreateTest("Validate_Account_Views_Filter_Smoke_Test");
            var client = new WebClient(TestSettings.Options);
            bool Flag = false;
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "Validate_Account_Views_Filter_Smoke_Test";
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

                    #region Predefined/Ref list for view filters

                    advancedFilterTest[] AccountsBeingFollowed ={
                    new advancedFilterTest(0, "dropDowns", "AND"),
                    new advancedFilterTest(0, "fields", "Regarding (Follows)"),
                    new advancedFilterTest(1, "dropDowns", "Contains data"),
                    new advancedFilterTest(2, "dropDowns", "AND"),
                    };

                    advancedFilterTest[] AccountsIFollow =
                    {
                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Regarding (Follows)"),
                        new advancedFilterTest(1, "dropDowns", "Contains data"),
                        new advancedFilterTest(2, "dropDowns", "AND"),
                        new advancedFilterTest(1, "fields", "Owner"),
                        new advancedFilterTest(2, "fields", "Equals current user"),
                    };

                    advancedFilterTest[] AccountInfluencedDealsthatWeWon =
                   {
                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Connected To (Connections)"),
                        new advancedFilterTest(1, "dropDowns", "Contains data"),
                        new advancedFilterTest(2, "dropDowns", "AND"),
                        new advancedFilterTest(1, "fields", "Role (To)"),
                        new advancedFilterTest(2, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Influencer"),
                        new advancedFilterTest(3, "fields", "Connected From (Opportunities)"),
                        new advancedFilterTest(3, "dropDowns", "Contains data"),
                        new advancedFilterTest(4, "dropDowns", "AND"),
                        new advancedFilterTest(4, "fields", "Status"),
                        new advancedFilterTest(5, "fields", "Equals"),
                        new advancedFilterTest(1, "values", "Won"),
                        new advancedFilterTest(6, "fields", "Actual Close Date"),
                        new advancedFilterTest(7, "fields", "Last x months"),
                        new advancedFilterTest(0, "FreeTextValue", "12"),



                    };

                    advancedFilterTest[] AccountsNoCampaignActivitiesinLast3Months =

                    {
                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(1, "dropDowns", "Or"),
                        new advancedFilterTest(0, "fields", "Last Date Included in Campaign"),
                        new advancedFilterTest(1, "fields", "Older than x months"),
                        new advancedFilterTest(0, "FreeTextValue", "3"),
                        new advancedFilterTest(2, "fields", "Last Date Included in Campaign"),
                        new advancedFilterTest(3, "fields", "Does not contain data"),
                        //new advancedFilterTest(1, "dropDowns", "Equals"),
                        

                    };

                    advancedFilterTest[] AccountsRespondedtoCampaignsinLast6Months =

                    {
                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Party (Activity Parties)"),
                        new advancedFilterTest(1, "dropDowns", "Contains data"),
                        new advancedFilterTest(2, "dropDowns", "AND"),
                        new advancedFilterTest(1, "fields", "Activity (Campaign Responses)"),
                        new advancedFilterTest(3, "dropDowns", "Contains data"),
                        new advancedFilterTest(4, "dropDowns", "AND"),
                        new advancedFilterTest(2, "fields", "Received On"),
                        new advancedFilterTest(3, "fields", "Last x months"),
                        new advancedFilterTest(0, "FreeTextValue", "6"),
                        new advancedFilterTest(4, "fields", "Participation Type"),
                        new advancedFilterTest(5, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Customer"),

                    };

                    advancedFilterTest[] ActiveAccounts =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Status"),
                        new advancedFilterTest(1, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Active"),

                     };

                    advancedFilterTest[] AllAccounts =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),


                     };

                    advancedFilterTest[] Customers =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Status"),
                        new advancedFilterTest(1, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Active"),
                        new advancedFilterTest(2, "fields", "Relationship Type"),
                        new advancedFilterTest(3, "fields", "Equals"),
                        new advancedFilterTest(1, "values", "Customer"),


                    };

                    advancedFilterTest[] ExcludedAccountsCampaigns =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Customers (Bulk Operation Logs)"),
                        new advancedFilterTest(1, "dropDowns", "Contains data"),
                        new advancedFilterTest(2, "dropDowns", "AND"),
                        new advancedFilterTest(1, "fields", "Reason Id"),
                        new advancedFilterTest(2, "fields", "Does not equal"),
                        new advancedFilterTest(0, "FreeTextValue", "0"),
                        new advancedFilterTest(3, "fields", "Reason Id"),
                        new advancedFilterTest(4, "fields", "Contains data"),


                    };

                    advancedFilterTest[] InactiveAccounts =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Status"),
                        new advancedFilterTest(1, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Inactive"),


                    };

                    advancedFilterTest[] MyActiveAccounts =

                   {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Owner"),
                        new advancedFilterTest(1, "fields", "Equals current user"),
                        new advancedFilterTest(2, "fields", "Status"),
                        new advancedFilterTest(3, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Active"),



                    };

                    advancedFilterTest[] MyConnections =

                   {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Status"),
                        new advancedFilterTest(1, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Active"),
                        new advancedFilterTest(2, "fields", "Connected To (Connections)"),
                        new advancedFilterTest(1, "dropDowns", "Contains data"),
                        new advancedFilterTest(2, "dropDowns", "AND"),
                        new advancedFilterTest(3, "fields", "Connected From"),
                        new advancedFilterTest(4, "fields", "Equals current user"),
                        new advancedFilterTest(5, "fields", "Status"),
                        new advancedFilterTest(6, "fields", "Equals"),
                        new advancedFilterTest(1, "values", "Active"),



                    };

                    advancedFilterTest[] SelectedAccountsCampaigns =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Customers (Bulk Operation Logs)"),
                        new advancedFilterTest(1, "dropDowns", "Contains data"),
                        new advancedFilterTest(2, "dropDowns", "AND"),
                        new advancedFilterTest(1, "fields", "Reason Id"),
                        new advancedFilterTest(2, "fields", "Equals"),
                        new advancedFilterTest(0, "FreeTextValue", "0"),
                        new advancedFilterTest(3, "fields", "Reason Id"),
                        new advancedFilterTest(4, "fields", "Contains data"),



                    };

                    advancedFilterTest[] ServiceAccount =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Status"),
                        new advancedFilterTest(1, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Active"),
                        new advancedFilterTest(2, "fields", "Billing Account"),
                        new advancedFilterTest(3, "fields", "Contains data"),




                    };

                    advancedFilterTest[] Vendors =

                    {

                        new advancedFilterTest(0, "dropDowns", "AND"),
                        new advancedFilterTest(0, "fields", "Status"),
                        new advancedFilterTest(1, "fields", "Equals"),
                        new advancedFilterTest(0, "values", "Active"),
                        new advancedFilterTest(2, "fields", "Relationship Type"),
                        new advancedFilterTest(3, "fields", "Equals"),
                        new advancedFilterTest(1, "values", "Vendor"),



                    };

                    #endregion

                    #region list of views in Accounts
                    Dictionary<string, advancedFilterTest[]> sysViewDictionary = new Dictionary<string, advancedFilterTest[]>();
                    sysViewDictionary.Add("Accounts Being Followed", AccountsBeingFollowed);
                    sysViewDictionary.Add("Accounts I Follow", AccountsIFollow);
                    sysViewDictionary.Add("Accounts: Influenced Deals That We Won", AccountInfluencedDealsthatWeWon);
                    sysViewDictionary.Add("Accounts: No Campaign Activities in Last 3 Months", AccountsNoCampaignActivitiesinLast3Months);
                    sysViewDictionary.Add("Accounts: Responded to Campaigns in Last 6 Months", AccountsRespondedtoCampaignsinLast6Months);
                    sysViewDictionary.Add("Active Accounts", ActiveAccounts);
                    sysViewDictionary.Add("All Accounts", AllAccounts);
                    sysViewDictionary.Add("Customers", Customers);
                    sysViewDictionary.Add("Excluded Accounts Campaigns", ExcludedAccountsCampaigns);
                    sysViewDictionary.Add("Inactive Accounts", InactiveAccounts);
                    sysViewDictionary.Add("My Active Accounts", MyActiveAccounts);
                    sysViewDictionary.Add("My Connections", MyConnections);
                    sysViewDictionary.Add("Selected Accounts Campaigns", SelectedAccountsCampaigns);
                    sysViewDictionary.Add("Service Account", ServiceAccount);
                    sysViewDictionary.Add("Vendors", Vendors);

                    #endregion

                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);

                    foreach (var sysView in sysViewDictionary)
                    {
                        HelperMethods.Grid_SwitchView(xrmApp, client, sysView.Key, true);
                        test.Log(Status.Info, "Switch View: " + sysView.Key);
                        HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);

                        HelperMethods.AdvancedFilter_Open(xrmApp, client);
                        test.Log(Status.Info, sysView.Key);

                        HelperMethods.ScreenShot(xrmApp, client, sysView.Key.Replace(':', '_'), childFolderPath);

                        var fields = HelperMethods.AdvancedFilter_GetFields(xrmApp, client);
                        var values = HelperMethods.AdvancedFilter_GetTagItems(xrmApp, client);
                        var dropDowns = HelperMethods.AdvancedFilter_GetDropDowns(xrmApp, client);
                        var freeTextField = HelperMethods.AdvancedFilter_GetFreeTextFields(xrmApp, client);

                        foreach (var views in sysViewDictionary[sysView.Key])
                        {
                            bool advFindTestPassed = HelperMethods.AdvancedFilter_TestFor(xrmApp, views, fields, values, dropDowns, freeTextField);
                            Assert.IsTrue(advFindTestPassed);
                        }

                        HelperMethods.AdvancedFilter_Cancel(xrmApp, client);
                        test.Log(Status.Info, "Verified Successfully: " + sysView.Key);
                        HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    }
                    Flag = true;
                    if (Flag)
                    {
                        test.Log(Status.Info, "All Views Validated Successfully ");
                        test.Pass("Test Passed");
                        HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "");
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
     
        [TestMethod]
        public void Validate_Account_Views_Order_SmokeTest()
        {
            var test = extent.CreateTest("Validate_Account_Views_Order_SmokeTest");
            var client = new WebClient(TestSettings.Options);
            bool Flag = false;
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "ValidateAccountOrderViews";
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

                    List<string> referenceList = new List<string> {
                            "Accounts Being Followed",
                            "Accounts I Follow",
                            "Accounts: Influenced Deals That We Won",
                            "Accounts: No Campaign Activities in Last 3 Months",
                            "Accounts: Responded to Campaigns in Last 6 Months",
                            "Active Accounts",
                            "All Accounts",
                            "Customers",
                            "Excluded Accounts Campaigns",
                            "Inactive Accounts",
                            "My Active Accounts",
                            "My Connections",
                            "Selected Accounts Campaigns",
                            "Service Account",
                            "Vendors",
                        };

                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    List<string> accountsViewList = HelperMethods.GetGridViewList(xrmApp, client, true);

                    if (accountsViewList[0] == "My Active Accounts")
                    {
                        accountsViewList.RemoveAt(0);
                    }

                    //Expected Results Step 1 - Views match the reference list, in order
                    bool viewsMatch = HelperMethods.CompareStringLists(referenceList, accountsViewList, true, true);
                    Assert.IsTrue(viewsMatch);
                    Flag = true;
                    if (Flag)
                    {
                        test.Log(Status.Info, "Account Views Order Validated Successfully");
                        test.Pass("Test Passed");
                        HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "ValidateAccountOrderViews");
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
}
