namespace Test
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Exchange.WebServices.Data;
    using Microsoft.Identity.Client;
    using Newtonsoft.Json;
    using Xunit.Abstractions;

    public class M365Context
    {
        public UserInformation AppAccount { get; set; }
    }

    public class UserInformation
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public string AdminUrl { get; set; }
        public string WebUrl { get; set; }

        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string ClientCertificate { get; set; }
        public string ClientCertificatePassword { get; set; }
        public string TenantId { get; set; }
        public string Scope { get; set; }
        public string Authority { get; set; }
        public string RedirectUrl { get; set; }
        public string ProxyUrl { get; set; }
    }

    static class TestUtility
    {
        public static T Get<T>()
        {
            var fileName = "TestData/" + typeof(T).Name + ".tests.json";
            using (var streamReader = new StreamReader(fileName, Encoding.UTF8))
            using (var reader = new JsonTextReader(streamReader))
            {
                JsonSerializer jsonSerializer = JsonSerializer.CreateDefault(new JsonSerializerSettings());
                return (T)jsonSerializer.Deserialize(reader, typeof(T));
            }
        }

        //public static SecureString ToSecureString(this string item)
        //{
        //    if (!string.IsNullOrEmpty(item))
        //    {
        //        var secureString = new SecureString();
        //        foreach (var c in item)
        //        {
        //            secureString.AppendChar(c);
        //        }

        //        return secureString;
        //    }

        //    return null;
        //}
    }

    public abstract class BaseTokenTest
    {
        protected readonly ITestOutputHelper output;

        protected M365Context m365Context = null;

        public BaseTokenTest(ITestOutputHelper output)
        {
            m365Context = TestUtility.Get<M365Context>();
            this.output = output;
        }

        private ExchangeService CreateExchangeServiceWithToken(UserInformation userInformation, string token)
        {
            var service = new ExchangeService();
            service.Url = new Uri(userInformation.WebUrl);
            service.Credentials = new OAuthCredentials(token);
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userInformation.Username);

            if (!string.IsNullOrEmpty(userInformation.WebUrl))
            {
                service.WebProxy = new System.Net.WebProxy(userInformation.ProxyUrl);
            }

            return service;
        }

        public ExchangeService CreateEXOServiceWithUserCredentialByMSAL(UserInformation userInformation)
        {
            var app = PublicClientApplicationBuilder
                .Create(userInformation.ClientId)
                .WithTenantId(userInformation.TenantId)
                .WithHttpClientFactory(new HttpClientFactoryWithProxy(userInformation.ProxyUrl))
                .Build();

            string[] scopes = { new Uri(userInformation.WebUrl).GetLeftPart(UriPartial.Authority) + "/.default" };

            var token = app.AcquireTokenByUsernamePassword(scopes, userInformation.Username, userInformation.Password).ExecuteAsync().ConfigureAwait(true).GetAwaiter().GetResult();

            return CreateExchangeServiceWithToken(userInformation, token.AccessToken);
        }

        public ExchangeService CreateEXOServiceWithUserNameByMSAL(UserInformation userInformation)
        {
            var app = PublicClientApplicationBuilder
                .Create(userInformation.ClientId)
                .WithTenantId(userInformation.TenantId)
                .WithHttpClientFactory(new HttpClientFactoryWithProxy(userInformation.ProxyUrl))
                .WithRedirectUri(userInformation.RedirectUrl)
                .Build();

            string[] scopes = { new Uri(userInformation.WebUrl).GetLeftPart(UriPartial.Authority) + "/.default" };

            var token = app.AcquireTokenInteractive(scopes).WithLoginHint(userInformation.Username).ExecuteAsync().ConfigureAwait(true).GetAwaiter().GetResult();


            return CreateExchangeServiceWithToken(userInformation, token.AccessToken);
        }

        public ExchangeService CreateEXOServiceWithAppCredentialsByMSAL(UserInformation userInformation)
        {
            var app = ConfidentialClientApplicationBuilder
                .Create(userInformation.ClientId)
                .WithCertificate(new System.Security.Cryptography.X509Certificates.X509Certificate2(userInformation.ClientCertificate, userInformation.ClientCertificatePassword))
                .WithTenantId(userInformation.TenantId)
                .WithHttpClientFactory(new HttpClientFactoryWithProxy(userInformation.ProxyUrl))
                .Build();

            string[] scopes = { new Uri(userInformation.WebUrl).GetLeftPart(UriPartial.Authority) + "/.default" };

            var token = app.AcquireTokenForClient(scopes).ExecuteAsync().ConfigureAwait(true).GetAwaiter().GetResult();

            return CreateExchangeServiceWithToken(userInformation, token.AccessToken);
        }

        class HttpClientFactoryWithProxy : IMsalHttpClientFactory //, IHttpClientFactory
        {
            private readonly HttpClient httpClient;

            public HttpClientFactoryWithProxy(string proxyUrl) 
            {
                if (!string.IsNullOrEmpty(proxyUrl))
                {
                    httpClient = new HttpClient(new HttpClientHandler()
                    {
                        Proxy = new System.Net.WebProxy(proxyUrl),
                        UseProxy = true
                    });
                }
                else
                {
                    httpClient = new HttpClient();
                }
            }

            public HttpClient GetHttpClient()
            {
                return httpClient;
            }
        }
    }
}
