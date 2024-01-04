using Azure.Identity;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.Models;

namespace ConsoleApp_MicrosoftGraphSDK_Test
{
    /// <summary>
    /// Delegated Permission
    /// </summary>
    public class App2
    {
        public string AppName { get; set; }
        public App2(string appName)
        {
            AppName = appName;
        }
        public async Task RunAsync()
        {
            //发送请求时需要携带的密钥，基本上就是POST请求中需要的请求体。
            string scopesText = Helper.GetJsonConfig(AppName, "Scopes");
            string[] scopes = new string[] { scopesText };
            string clientId = Helper.GetJsonConfig(AppName, "ClientId");
            string tenantId = Helper.GetJsonConfig(AppName, "TenantId");
            string username = Helper.GetJsonConfig(AppName, "Username");
            string password = Helper.GetJsonConfig(AppName, "Password");

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };
            var userCredential = new UsernamePasswordCredential(username,password,tenantId,clientId,options);

            GraphServiceClient graphClient = new GraphServiceClient(userCredential, scopes);
            try
            {
                //await deleteFileById(graphClient);
                await getRootChildrenListAsync(graphClient);

            }
            //读取ODataError错误中的详细信息才能了解为什么请求失败。
            catch (ODataError odataError)
            {
                Console.WriteLine(string.Format("Error Code: {0}", odataError.Error.Code));
                Console.WriteLine(string.Format("Error Message: {0}", odataError.Error.Message));
                throw;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
                throw;
            }
        }

        public async Task getRootChildrenListAsync(GraphServiceClient graphClient)
        {
            //这一步获取网盘的ID，然后后面通过ID来获取根目录下的文件。即使只有一个网盘，也需要这样才行。
            Drive drive = await graphClient.Me.Drive.GetAsync();
            //var i=drive.Items; //值为null
            //var r = drive.Root;  //值为null

            //获取指定网盘ID的root/children,根目录下的所有文件和文件夹。
            DriveItem rootItem = await graphClient
                .Drives[drive.Id].Root
                .GetAsync(conf =>
                {
                    conf.QueryParameters.Expand = new[] { "children" };
                });

            //显示获取的子文件夹或文件
            foreach (DriveItem child in rootItem?.Children)
            {
                Console.WriteLine("Id= " + child.Id);
                Console.WriteLine("Name= " + child.Name);
                Console.WriteLine("CDT= " + child.CreatedDateTime);
                Console.WriteLine("ChildCount= " + child.Folder?.ChildCount);
                Console.WriteLine("Create By= " + child.CreatedBy.User.DisplayName);
                Console.WriteLine("Package=" + child.Package);
                Console.WriteLine("WebUrl=" + child.WebUrl);
                Console.WriteLine();
            }
        }
    }
}
