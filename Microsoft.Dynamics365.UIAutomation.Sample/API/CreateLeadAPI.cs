using Microsoft.Dynamics365.UIAutomation.Sample.Models.RequestModel;
using Microsoft.Dynamics365.UIAutomation.Sample.Models.ResponseModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Net.Http;
using System.Text;

namespace Microsoft.Dynamics365.UIAutomation.Sample.API
{
    [TestClass]
    public class CreateLeadAPI : BaseClass
    {
       
        public async void CreateLeadAPITest()
        {
            //HttpClient client = new HttpClient();

            ////Setting request model values here..
            //var leadDt = new LeadDetailsRequest() { Topic = "", FirstName = "", LastName = "" };
            //HttpRequestMessage content = new HttpRequestMessage(HttpMethod.Post, "contacts")
            //{
            //    Content = new StringContent(leadDt.ToString(), Encoding.UTF8, "application/json")
            //};

            ////API call here..
            //HttpResponseMessage response = await client.PostAsync(BaseURL + "/Lead", content);
            //var Actual = await response.Content.ReadAsStringAsync();

            ////Set Expected Outcome here..
            //var Expected = new LeadDetailsResponse()
            //{
            //    Status = "Success",
            //    Message = "Lead Created in Sales"
            //};

            ////Compare both Actual vs Expected..
            //bool checker = compareData(Actual, Expected);

            //Assert.IsTrue(checker);

        }
        private bool compareData(string actual, LeadDetailsResponse expected)
        {
            throw new NotImplementedException();
        }
    }

}
