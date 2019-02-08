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
using Microsoft.InformationProtection;
using Microsoft.IdentityModel.Clients.ActiveDirectory;


namespace MipSdkDotNetQuickstart
{
    public class AuthDelegateImplementation : IAuthDelegate
    {
        // Set the redirect URI from the AAD Application Registration.
        private static readonly string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private ApplicationInfo appInfo;        
        private TokenCache tokenCache = new TokenCache();
        
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
        public string AcquireToken(Identity identity, string authority, string resource)
        {
            try
            { 
                // Create an auth context using the provided authority and token cache
                AuthenticationContext authContext = new AuthenticationContext(authority, tokenCache);
                               
                // Attempt to acquire a token for the given resource, using the ApplicationId, redirectUri, and Identity
                var result = authContext.AcquireTokenAsync(resource, appInfo.ApplicationId, new Uri(redirectUri), new PlatformParameters(PromptBehavior.Auto, null), new UserIdentifier(identity.Email, UserIdentifierType.RequiredDisplayableId)).Result;
                
                // Return the token. The token is sent to the resource.
                return result.AccessToken;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// The GetUserIdentity() method is used to pre-identify the user and obtain the UPN. 
        /// The UPN is later passed set on FileEngineSettings for service location.
        /// </summary>
        /// <returns>Microsoft.InformationProtection.Identity</returns>
        public Identity GetUserIdentity()
        {
            try
            {
                string resource = "https://graph.microsoft.com/";
              
                AuthenticationContext authContext = new AuthenticationContext("https://login.windows.net/common", tokenCache);
                var result = authContext.AcquireTokenAsync(resource, appInfo.ApplicationId, new Uri(redirectUri), new PlatformParameters(PromptBehavior.Always, null)).Result;
                return new Identity(result.UserInfo.DisplayableId);

            }
            catch(Exception ex)
            {
                throw ex;
            }
        }


    }
}
