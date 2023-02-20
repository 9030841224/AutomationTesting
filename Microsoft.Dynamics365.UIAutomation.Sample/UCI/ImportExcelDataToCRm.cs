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

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class ImportExcelDataToCRm
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        string parentFolder = "D:\\New folder\\EasyRepro-develop\\TestResults";
       

        [TestMethod]
        public void ImportExcelDataToAccounts()
        {
            var client = new WebClient(TestSettings.Options);
            InfoConstants Info = new InfoConstants();
            try
            {
                using (var xrmApp = new XrmApp(client))
                {
                    string childFolder = "ImportExcelDataToAccounts";
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

                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                    HelperMethods.ScreenShot(xrmApp, client, Info.login, childFolderPath);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    HelperMethods.ScreenShot(xrmApp, client, Info.OpenSales, childFolderPath);
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");


                    HelperMethods.UploadExcel(xrmApp, client);

                    Assert.IsTrue(true);
                }
            }
            catch(Exception ex)
            {
                Assert.IsTrue(false);
            }
        }
        
    }
}
