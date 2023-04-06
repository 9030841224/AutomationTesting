using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Models.ResponseModel
{
    public class LeadDetailsResponse
    {
        public string Status { get; set; } = "Success";

        public string Message { get; set; } = "Lead Created";
    }
}
