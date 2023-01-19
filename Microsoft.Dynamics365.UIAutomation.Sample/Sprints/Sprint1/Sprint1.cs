using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using OpenQA.Selenium;
using System.Collections.Generic;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Sprints.Sprint1
{
    [TestClass]
    public class Sprint1
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void CreateNewAccount()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                xrmApp.CommandBar.ClickCommand("New");
                String BusinessName = "Test Account";
                HelperMethods.CreateTestAccount(xrmApp, BusinessName);
                xrmApp.Entity.Save();
            }
        }

        [TestMethod]
        public void AcountTabsVerify()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");

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
                Assert.IsTrue(result && referenceList.Count == tabList.Count);
            }

        }

        [TestMethod]
        public void AcountCoreFieldsVerify()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");
                xrmApp.Entity.SelectTab("Summary");
                HelperMethods.WaitForInvisibilityOfProgressIndicator(xrmApp, client);
                xrmApp.ThinkTime(3000);

                // Core field verified list
                List<string> referenceListcorefield = new List<string> {
                    "Account Name"
                    //"Relationship Type"
                };

                // NonCore field verified list
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
            }
        }

        [TestMethod]
        public void AcountPhoneNumberValidation()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                HelperMethods.Grid_SwitchView(xrmApp, client,"My Active Accounts");

                HelperMethods.OpenFirstRecord(xrmApp, client);

                var phone = xrmApp.Entity.GetValue("telephone1");
                bool validatePhoneNumber = HelperMethods.ValidatePhoneNumber(phone);
                Assert.IsTrue(validatePhoneNumber);
            }
        }
        [TestMethod]

        public void CreateContact()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Contacts");
                
                xrmApp.CommandBar.ClickCommand("New");
                String BusinessName = "Contact";
                HelperMethods.CreateContact(xrmApp, BusinessName);

                xrmApp.Entity.Save();

                //xrmApp.CommandBar.ClickCommand("Delete");

            }
        }
    }
}