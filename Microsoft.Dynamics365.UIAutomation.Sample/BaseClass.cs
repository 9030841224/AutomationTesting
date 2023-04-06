using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class BaseClass
    {        
        protected readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        protected readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        protected readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        protected readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        protected readonly static string parentFolder = System.Configuration.ConfigurationManager.AppSettings["ParentFolder"];
        protected readonly string connString = System.Configuration.ConfigurationManager.AppSettings["OLEDConnectionString"];
        protected readonly string BaseURL = System.Configuration.ConfigurationManager.AppSettings["BaseURL"];
        protected static ExtentReports extent;

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

        public static string CreateFolder(string ChildFolderName)
        {
            string childFolder = ChildFolderName;
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

            return childFolderPath;
        }
    }
}
