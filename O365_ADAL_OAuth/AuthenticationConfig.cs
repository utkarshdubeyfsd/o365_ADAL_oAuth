using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace O365_ADAL_OAuth
{
    /// <summary>
    /// Description of the configuration of an AzureAD public client application (desktop/mobile application). This should
    /// match the application registration done in the Azure portal
    /// </summary>
    public class AuthenticationConfig
    {
        /// <summary>
        /// instance of Azure AD, for example public Azure or a Sovereign cloud (Azure China, Germany, US government, etc ...)
        /// </summary>
        public string Instance { get; set; } = "https://login.microsoftonline.com/{0}";

        /// <summary>
        /// The Tenant is:
        /// - either the tenant ID of the Azure AD tenant in which this application is registered (a guid)
        /// or a domain name associated with the tenant
        /// - or 'organizations' (for a multi-tenant application)
        /// </summary>
        public string Tenant { get; set; }

        /// <summary>
        /// login name or email id or user prinicipal name of sign-in user
        /// </summary>
        public string userPrincipalName { get; set; }

        /// <summary>
        /// library id of that library where file is going to upload
        /// it can be retrieve by using SharePoint Graph API, GET method: https://graph.microsoft.com/v1.0/sites/root/drives 
        /// </summary>
        public string LibraryId { get; set; }

        /// <summary>
        /// Guid used by the application to uniquely identify itself to Azure AD
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// URL of the authority
        /// </summary>
        public string Authority
        {
            get
            {
                return String.Format(CultureInfo.InvariantCulture, Instance, Tenant);
            }
        }

#if !VariationWithCertificateCredentials
        /// <summary>
        /// Client secret (application password)
        /// </summary>
        /// <remarks>Daemon applications can authenticate with AAD through two mechanisms: ClientSecret
        /// (which is a kind of application password: this property)
        /// or a certificate previously shared with AzureAD during the application registration 
        /// (and identified by the CertificateName property belows)
        /// <remarks> 
        public string ClientSecret { get; set; }
#else
        /// <summary>
        /// Name of a certificate in the user certificate store
        /// </summary>
        /// <remarks>Daemon applications can authenticate with AAD through two mecanisms: ClientSecret
        /// (which is a kind of application password: the property above)
        /// or a certificate previously shared with AzureAD during the application registration 
        /// (and identified by this CertificateName property)
        /// <remarks> 
        public string CertificateName { get; set; }
#endif
        /// <summary>
        /// Reads the configuration from a json file
        /// </summary>
        /// <param name="path">Path to the configuration json file</param>
        /// <returns>AuthenticationConfig read from the json file</returns>
        public static AuthenticationConfig ReadFromJsonFile(string path)
        {
            IConfigurationRoot Configuration;

            var builder = new ConfigurationBuilder()
             .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(path);

            Configuration = builder.Build();
            return Configuration.Get<AuthenticationConfig>();
        }
    }
}
