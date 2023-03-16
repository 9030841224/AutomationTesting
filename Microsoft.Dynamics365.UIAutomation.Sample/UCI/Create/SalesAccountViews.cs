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
    public class SalesAccountViews
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
            var htmlReporter = new ExtentHtmlReporter(@"D:\New folder\EasyRepro-develop\Microsoft.Dynamics365.UIAutomation.Sample\ExtentReports\" + r.Next(10, 100) + "index.HTML");
            extent.AttachReporter(htmlReporter);
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

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            extent.Flush();
        }

    }
}
