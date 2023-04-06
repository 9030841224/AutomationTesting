//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.IdentityModel.Clients.ActiveDirectory;
////using Salam.CRM.API.Models;
//using System.Configuration;
//using System.Diagnostics;
//using System.IO;
//using System.Net;
//using System.Net.Http;
//using System.Net.Http.Headers;
////using System.Runtime.Caching;


//namespace Microsoft.Dynamics365.UIAutomation.Sample.Authentication
//{
//	public class CRMConnection
//	{
//		public static async Task<HttpClient> GetD365Client()
//		{
//			try
//			{
//				int timeoutValue = Convert.ToInt32(ConfigurationManager.AppSettings["Timeout"]);

//				var accessToken = await GenerateAccessToken();
//				//var accessToken = await GetAccessToken();
//				//Enabling TLS12 Protocal
//				//System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

//				var client = new HttpClient();
//				client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
//				client.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
//				client.DefaultRequestHeaders.Add("OData-Version", "4.0");
//				client.DefaultRequestHeaders.Add("Accept", "application/json");
//				client.DefaultRequestHeaders.Add("Prefer", "odata.include-annotations=\"*\"");
//				client.Timeout = new TimeSpan(0, 0, timeoutValue);

//				return client;
//			}
//			catch (Exception ex)
//			{
//				throw;
//			}
//		}
//		public static async Task<AuthenticationResult> GetAccessToken()
//		{
//			try
//			{
//				string userName = ConfigurationManager.AppSettings["CRMUser"].ToString();
//				string password = ConfigurationManager.AppSettings["CRMUserPassword"].ToString();
//				string decryptedPassword = EncryptionManager.Decrypt(password);

//				string crmUrl = ConfigurationManager.AppSettings["CRMAPIURL"].ToString();
//				string applicationId = ConfigurationManager.AppSettings["APPID"].ToString();

//				AuthenticationParameters ap = AuthenticationParameters.CreateFromResourceUrlAsync(new Uri(crmUrl)).Result;
//				string resourceUrl = ap.Resource;
//				string authorityUrl = ap.Authority;
//				AuthenticationContext authContext = new AuthenticationContext(authorityUrl, false);
//				UserCredential cred = new UserPasswordCredential(userName, decryptedPassword);

//				AuthenticationResult authToken = await authContext.AcquireTokenAsync(resourceUrl, applicationId, cred);
//				//try
//				//{
//				//    authToken = await authContext.AcquireTokenSilentAsync(resourceUrl, applicationId);
//				//}
//				//catch (AdalException adalException)
//				//{
//				//    if (adalException.ErrorCode == AdalError.FailedToAcquireTokenSilently
//				//        || adalException.ErrorCode == AdalError.InteractionRequired)
//				//    {
//				//        authToken = await authContext.AcquireTokenAsync(resourceUrl, applicationId, cred);
//				//    }
//				//}
//				return authToken;
//			}
//			catch (Exception ex)
//			{
//				throw;
//			}
//		}		
//		private static async Task<string> GenerateAccessToken()
//		{
//			try
//			{
//				int cacheValue = Convert.ToInt32(ConfigurationManager.AppSettings["CacheTime"]);
//				ObjectCache cache = MemoryCache.Default;
//				var authResult = cache.Get("AccessToken") as AuthenticationResult;

//				if (authResult != null && authResult.ExpiresOn > System.DateTimeOffset.UtcNow && !string.IsNullOrEmpty(authResult.AccessToken))
//				{
//					return authResult.AccessToken;
//				}

//				AuthenticationResult authenticationResult = await GetAccessToken();
//				var accessToken = authenticationResult.AccessToken;
//				var tokenExpiredOnMin = authenticationResult.ExpiresOn.ToString();
//				var extendedLifeTimeToken = authenticationResult.ExtendedLifeTimeToken.ToString();

//				CacheItemPolicy policy = new CacheItemPolicy { AbsoluteExpiration = DateTime.Now.AddMinutes(cacheValue) };

//				if (cache.Get("AccessToken") != null)
//				{
//					MemoryCache.Default.Remove("AccessToken");
//				}
//				cache.Add("AccessToken", authenticationResult, policy);

//				var strLogTexts = "New Access Token : " + accessToken;
//				strLogTexts += Environment.NewLine + "Token Expired Time : " + tokenExpiredOnMin;
//				strLogTexts += Environment.NewLine + "Current UTC time : " + System.DateTimeOffset.UtcNow;
//				strLogTexts += Environment.NewLine + "ExtendedLifeTimeToken : " + extendedLifeTimeToken;
//				strLogTexts += Environment.NewLine + "AuthenticationResult of Cache : " + authResult;
//				strLogTexts += Environment.NewLine + "Current Server Time : " + DateTime.Now;

//				SqlErrorLogging sqlErrorLoggings = new SqlErrorLogging();
//				ApiError apiErrors = new ApiError()
//				{
//					Message = strLogTexts,
//					RequestUri = string.Empty,
//					RequestMethod = "GET",
//					TimeUtc = DateTime.Now,
//					RequestBody = string.Empty
//				};
//				sqlErrorLoggings.InsertErrorLog(apiErrors);
//				return accessToken;
//			}
//			catch (Exception Ex)
//			{
//				throw;
//			}

//		}
//	}
//}
