using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System.Security.Cryptography.X509Certificates;

static X509Certificate2 LoadCertificate(string certificateThumbprint)
{
    // Note: https://jan-v.nl/post/loading-certificates-with-azure-functions/
    // Must Update Function App Settings 
    var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
    store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
    var certificateCollection = store.Certificates.Find(X509FindType.FindByThumbprint, certificateThumbprint, false);
    store.Close();

    return certificateCollection.First();
}

var host = new HostBuilder()
.ConfigureServices((hostingContext, services) =>
{
    services.AddPnPCore();
    // Add PnP Core SDK
        services.AddPnPCore(options =>
        {
            // Add the base site url
            options.Sites.Add("Default", new PnPCoreSiteOptions
            {
                SiteUrl = Environment.GetEnvironmentVariable("SiteUrl")
            });
        });

        // Manual configure the used authentication
        services.AddPnPCoreAuthentication(options =>
        {
            // Load the certificate that will be used to authenticate
            //var cert = LoadCertificate(certificatePath);

            // Configure certificate based auth
            options.Credentials.Configurations.Add("CertAuth", new PnPCoreAuthenticationCredentialConfigurationOptions
            {
                ClientId = Environment.GetEnvironmentVariable("ClientId"),
                TenantId = Environment.GetEnvironmentVariable("TenantId"),
                X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
                {
                    Certificate = LoadCertificate(Environment.GetEnvironmentVariable("Thumbprint")),
                }
            });

            // Configure the default authentication provider
            options.Credentials.DefaultConfiguration = "CertAuth";
            
            // Connect the configured authentication method to the configured site
            options.Sites.Add("Default", new PnPCoreAuthenticationSiteOptions
            {
                AuthenticationProviderName = "CertAuth",
            });

            options.Credentials.DefaultConfiguration = "CertAuth";
        });
})
.ConfigureFunctionsWorkerDefaults()
.Build();

host.Run();
