/*
*
* Copyright (c) Microsoft Corporation.
* All rights reserved.
*
* This code is licensed under the MIT License.
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files(the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions :
*
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
* THE SOFTWARE.
*
*/

using System;
using System.Configuration;
using System.Collections.Generic;
using System.Threading.Tasks;

using Microsoft.InformationProtection;
using Microsoft.Identity.Client;
using System.Linq;

namespace MipSdkDotNetQuickstart
{
    public class AuthDelegateImplementation : IAuthDelegate
    {
        // Set the redirect URI from the AAD Application Registration.
        private static readonly string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private static readonly bool isMultitenantApp = Convert.ToBoolean(ConfigurationManager.AppSettings["ida:IsMultitenantApp"]);
        private static readonly string tenant = ConfigurationManager.AppSettings["ida:TenantGuid"];
        private ApplicationInfo appInfo;

        // Microsoft Authentication Library IPublicClientApplication
        private IPublicClientApplication _app;

        // Define MSAL scopes.
        // As of the 1.7 release, the two services backing the MIP SDK, RMS and MIP Sync Service, provide resources instead of scopes.
        // The List<string> entities below will be used to map the resources to scopes and to pass those scopes to Azure AD via MSAL.

        /// <summary>
        /// Delegated scopes for Azure RMS.
        /// </summary>
        private List<string> aadrmScopes = new List<string>()
        {
            "https://aadrm.com/user_impersonation" // Delegated API permission, valid for public client applications.         
        };

        /// <summary>
        /// Delegated scope for the Microsoft Information Protection Sync Service.
        /// </summary>
        private List<string> syncServiceScopes = new List<string>()
        {
            "https://psor.o365syncservice.com/UnifiedPolicy.User.Read" // Delegated API permission, allows policy sync for authenticated user.            
        };

        /// <summary>
        /// Delegated permission for Graph. This is *not* required by MIP SDK and is used only in this sample app.
        /// </summary>
        List<string> graphScope = new List<string>()
        {
            "https://graph.microsoft.com/User.Read" // Graph scope for triggering sign in for user identity discovery.            
        };

        public AuthDelegateImplementation(ApplicationInfo appInfo)
        {
            this.appInfo = appInfo;
        }

        /// <summary>
        /// AcquireToken is called by the SDK when auth is required for an operation. 
        /// Adding or loading an IFileEngine is typically where this will occur first.
        /// The SDK provides all three parameters below.Identity from the EngineSettings.
        /// Authority and resource are provided from the 401 challenge.
        /// The SDK cares only that an OAuth2 token is returned.How it's fetched isn't important.
        /// In this sample, we fetch the token using Active Directory Authentication Library(ADAL).
        /// </summary>
        /// <param name="identity"></param>
        /// <param name="authority"></param>
        /// <param name="resource"></param>
        /// <returns>The OAuth2 token for the user</returns>
        public string AcquireToken(Identity identity, string authority, string resource, string claims)
        {            
                return AcquireTokenAsync(authority, resource, claims, isMultitenantApp).Result.AccessToken;            
        }

        /// <summary>
        /// Implements token acquisition logic via the Microsoft Authentication Library.
        /// 
        /// /// </summary>
        /// <param name="identity"></param>
        /// <param name="authority"></param>
        /// <param name="resource"></param>
        /// <param name="claims"></param>
        /// <returns></returns>
        public async Task<AuthenticationResult> AcquireTokenAsync(string authority, string resource, string claims, bool isMultiTenantApp = true)
        {
            AuthenticationResult result = null;

            // Create an auth context using the provided authority and token cache
            if (isMultitenantApp)
                _app = PublicClientApplicationBuilder.Create(appInfo.ApplicationId)
                    .WithAuthority(authority)
                    .WithDefaultRedirectUri()
                    .Build();
            else
            {
                if (authority.ToLower().Contains("common"))
                {
                    authority = authority.Remove(authority.Length - 6, 6);
                }
                _app = PublicClientApplicationBuilder.Create(appInfo.ApplicationId)
                    .WithAuthority(authority + tenant)
                    .WithDefaultRedirectUri()
                    .Build();

            }
            var accounts = (_app.GetAccountsAsync()).GetAwaiter().GetResult();

            List<string> scopes = new List<string>();

            if (resource.ToLower().Contains("aadrm"))
            {
                scopes = aadrmScopes;
            }
            else if (resource.ToLower().Contains("graph"))
            {
                scopes = graphScope;
            }
            else
            {
                scopes = syncServiceScopes;
            }

            try
            {
                result = await _app.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }

            catch (MsalUiRequiredException)
            {
                result = _app.AcquireTokenInteractive(scopes)
                    .WithAccount(accounts.FirstOrDefault())
                    .WithPrompt(Prompt.SelectAccount)
                    .ExecuteAsync()
                    .ConfigureAwait(false)
                    .GetAwaiter()
                    .GetResult();
            }

            // Return the token. The token is sent to the resource.                           
            return result;
        }

        /// <summary>
        /// The GetUserIdentity() method is used to pre-identify the user and obtain the UPN. 
        /// The UPN is later passed set on FileEngineSettings for service location.
        /// </summary>
        /// <returns>Microsoft.InformationProtection.Identity</returns>
        public Identity GetUserIdentity()
        {
            AuthenticationResult result = AcquireTokenAsync("https://login.microsoftonline.com/common", "https://graph.microsoft.com", null).Result;
            return new Identity(result.Account.Username);
        }
    }
}
