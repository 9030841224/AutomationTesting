using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Net.Http;

namespace Microsoft.Dynamics365.UIAutomation.Sample.API
{
    [TestClass]
    public class AccountAPITest
    {
        [TestMethod]
        [TestCategory("Create account")]
        public void CreateAccountAPITest()
        {
            HttpClient client = new HttpClient();
            HttpResponseMessage response = client.PostAsync("api/products", "");
            response.EnsureSuccessStatusCode();


        }
    }

    public class Product
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Category { get; set; }
    }
}
