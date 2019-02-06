using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.InformationProtection;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace MipSdkDotNetQuickstart
{
    public class AuthDelegateImplementation : IAuthDelegate
    {
        private ApplicationInfo appInfo;
        private readonly string redirectUri = "mip-sdk://authorization";
        private TokenCache tokenCache = new TokenCache();


        public AuthDelegateImplementation(ApplicationInfo appInfo)
        {
            this.appInfo = appInfo;          
        }

        public string AcquireToken(Identity identity, string authority, string resource)
        {
            try
            { 
                AuthenticationContext authContext = new AuthenticationContext(authority, tokenCache);
                var result = authContext.AcquireTokenAsync(resource, appInfo.ApplicationId, new Uri(redirectUri), new PlatformParameters(PromptBehavior.Auto, null), new UserIdentifier(identity.Email, UserIdentifierType.RequiredDisplayableId)).Result;
                return result.AccessToken;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Use this method to force auth to get a token and store identity. 
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
