using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using AventStack.ExtentReports;
using System.IO;
using static Microsoft.Dynamics365.UIAutomation.Api.UCI.HelperMethods;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI.Create
{
    [TestClass]
    public class SaleProcessFlowTest : BaseClass
    {
        
        [TestMethod]
        public void CreateLead()
        {
            var test = extent.CreateTest("CreateNewlead");
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "CreateLead";
                    if (!Directory.Exists(parentFolder))
                    {
                        Directory.CreateDirectory(parentFolder);
                    }
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
                    xrmApp.Navigation.OpenSubArea("Sales", "Leads");
                    test.Log(Status.Info, "Navigating to lead");
                    xrmApp.CommandBar.ClickCommand("New");
                    test.Log(Status.Info, "Creating new Lead");
                    String BusinessName = "Lead";
                    HelperMethods.CreateLead(xrmApp, BusinessName);
                    xrmApp.Entity.Save();
                    test.Log(Status.Info, "Lead Created");
                    HelperMethods.ScreenShot(xrmApp, client, "A. LeadCreated", childFolderPath);
                    Assert.IsTrue(true);
                    xrmApp.CommandBar.ClickCommand("Qualify");
                    test.Log(Status.Info, "Lead Qualified");
                    HelperMethods.ScreenShot(xrmApp, client, "B.Lead Qualified", childFolderPath);
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);

                    xrmApp.Entity.SelectTab("Product line items");
                    xrmApp.Entity.SetValue(new LookupItem() { Name = "pricelevelid", Value = "France Bill Rates" });
                    xrmApp.ThinkTime(4000);

                    int totalEquCnt = HelperMethods.GetSubGridItemsCount(xrmApp, client, "opportunityproductsGrid", 2000, HelperMethods.GridType.pcfGrid);
                    xrmApp.ThinkTime(2000);
                    Assert.IsTrue(totalEquCnt >= 0);

                    test.Log(Status.Info, "Products added");
                    xrmApp.Entity.SelectTab("Quotes");
                    HelperMethods.Clickcommands(xrmApp, client, Commands.NewQuoteFromQuoteTab);
                    test.Log(Status.Info, "Create New Quote");
                    HelperMethods.ScreenShot(xrmApp, client, "C. Quote Created", childFolderPath);
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    HelperMethods.ClickCommand(xrmApp, client, "Activate Quote");
                    test.Log(Status.Info, "Activate Quote");
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    HelperMethods.ScreenShot(xrmApp, client, "E. Quote Activated", childFolderPath);
                    xrmApp.CommandBar.ClickCommand("Create Order");
                    test.Log(Status.Info, "Create Order");
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    HelperMethods.Clickcommands(xrmApp, client, Commands.OKFromPopUp);
                    HelperMethods.ScreenShot(xrmApp, client, "F. Order Created", childFolderPath);
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    xrmApp.CommandBar.ClickCommand("Create Invoice");
                    test.Log(Status.Info, "Create Invoice");
                    HelperMethods.ScreenShot(xrmApp, client, "G. Create Invoice", childFolderPath);
                    var isCreated = HelperMethods.FindElementXpath_notificationWrapper(xrmApp, client);
                    HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    if (isCreated.Count != 0)
                    {
                        test.Log(Status.Fail, "All Mandatory fields need to be provide");                       
                         HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                        Assert.IsTrue(false);
                    }
                    xrmApp.ThinkTime(4000);
                    HelperMethods.ClickCommand(xrmApp, client, "Confirm Invoice");
                    test.Log(Status.Info, "Confirm Invoice");
                    HelperMethods.ScreenShot(xrmApp, client, "H. Confirm Invoice", childFolderPath);
                    xrmApp.ThinkTime(4000);
                    HelperMethods.ClickCommand(xrmApp, client, "Invoice Paid");
                    test.Log(Status.Info, "Invoice Paid");
                    HelperMethods.Clickcommands(xrmApp, client, Commands.OKFromPopUp);
                    HelperMethods.ScreenShot(xrmApp, client, "I. Invoice Paid", childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Opportunities");
                    HelperMethods.ShowLayoutCancel(xrmApp, client);
                    xrmApp.ThinkTime(5000);

                    ////HelperMethods.Grid_SwitchView(xrmApp, client, "All Opportunities");
                    ////HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                    ////HelperMethods.OpenFirstRecord(xrmApp, client);
                    
                    HelperMethods.CopyScreenShotsIntoWord(xrmApp, client, childFolderPath, "CopyScreenshotsIntoWord");
                    test.Log(Status.Pass, "Sales Process Completed");

                }
            }
            catch
            {
                test.Fail("Test Failed");
                Assert.IsTrue(false);
            }
        }
    }
}
