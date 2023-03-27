using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class BaseClass
    {
        protected static ExtentReports extent;
        protected readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        protected readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        protected readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        protected readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        //protected readonly string parentFolder = "D:\\New folder\\EasyRepro-develop\\TestResults";
        protected readonly string parentFolder = System.Configuration.ConfigurationManager.AppSettings["ParentFolder"];
        //public const string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\\Projects\\Accounts.xlsx; Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'";
        protected readonly string connString = System.Configuration.ConfigurationManager.AppSettings["OLEDConnectionString"];

        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            extent = new ExtentReports();
            Random r = new Random();
            var htmlReporter = new ExtentHtmlReporter(@"D:\New folder\EasyRepro-develop\Microsoft.Dynamics365.UIAutomation.Sample\ExtentReports\" + r.Next(10, 100) + "index.HTML");
            extent.AttachReporter(htmlReporter);
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            extent.Flush();
        }
    }
}
