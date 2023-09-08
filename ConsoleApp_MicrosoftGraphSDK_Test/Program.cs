using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System.Threading.Channels;

namespace ConsoleApp_MicrosoftGraphSDK_Test
{
    public class Program
    {
        static async Task Main(string[] args)
        {   
            //发送请求时需要携带的密钥，基本上就是POST请求中需要的请求体。
            string scopesText= getJsonConfig("Scopes");
            string[] scopes = new string[] { scopesText };
            string clientId = getJsonConfig("ClientId");
            string tenantId = getJsonConfig("TenantId");
            string clientSecret = getJsonConfig("ClientSecret");

            // using Azure.Identity;
            ClientSecretCredentialOptions options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };
            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            try
            {

                //获取指定用户的网盘，然后通过drive.Id获取网盘ID。
                var drive = await graphClient.Users["a37cb0b6-562d-422d-bdb3-2063e6867316"].Drive.GetAsync();

                //获取指定网盘ID的root/children,根目录下的所有文件和文件夹
                DriveItem rootItem = await graphClient
                    .Drives["b!h2paA_qdHkWMFMnaxQ6505czytzKNNJBhQafMxIUPLnZtSSEzeB5Q6gCbA9lBz0K"].Root
                    .GetAsync(conf => 
                    { 
                        conf.QueryParameters.Expand = new[] { "children" }; 
                    });

                //显示获取的子文件夹或文件
                foreach(DriveItem child in rootItem?.Children)
                {
                    Console.WriteLine("Id= "+child.Id);
                    Console.WriteLine("Name= "+child.Name);
                    Console.WriteLine("CDT= "+child.CreatedDateTime);
                    Console.WriteLine("ChildCount= "+child.Folder?.ChildCount);
                    Console.WriteLine("Create By= " + child.CreatedBy.User.DisplayName);
                    Console.WriteLine("Package="+child.Package);
                    Console.WriteLine("WebUrl="+child.WebUrl);
                    Console.WriteLine();
                }

                //根据ID获取指定文件夹下的文件
                //ID:01C4GJWGN44II7IQNCGZFKHLTSVBFC6AYX
                DriveItem publicShareFolderItem = await graphClient
                    .Drives["b!h2paA_qdHkWMFMnaxQ6505czytzKNNJBhQafMxIUPLnZtSSEzeB5Q6gCbA9lBz0K"]
                    .Items["01C4GJWGN44II7IQNCGZFKHLTSVBFC6AYX"]
                    .GetAsync(conf =>
                    {
                        conf.QueryParameters.Expand = new[] { "children" };
                    });
                Console.WriteLine("PublicShare folder size= "+publicShareFolderItem.Size);
                Console.WriteLine("PublicShare folder count= " + publicShareFolderItem.Children.Count);
                foreach (DriveItem child in publicShareFolderItem?.Children)
                {
                    Console.WriteLine("Id= " + child.Id);
                    Console.WriteLine("Name= " + child.Name);
                    Console.WriteLine("CDT= " + child.CreatedDateTime);
                    Console.WriteLine("Create By= " + child.CreatedBy.User.DisplayName);
                    Console.WriteLine("WebUrl=" + child.WebUrl); //该URL不能给其他用户直接访问，其他用户需要登录Offic365才能访问。
                    Console.WriteLine("Size=" + child.Size);
                    Console.WriteLine();
                }

                //根据文件路径获取文件信息
                DriveItem fileItem = await graphClient
                        .Drives["b!h2paA_qdHkWMFMnaxQ6505czytzKNNJBhQafMxIUPLnZtSSEzeB5Q6gCbA9lBz0K"]
                        .Root.ItemWithPath("/PublicShare/model-10.jpg").GetAsync();
                Console.WriteLine("Id= " + fileItem.Id);
                Console.WriteLine("Name= " + fileItem.Name);
                Console.WriteLine("CDT= " + fileItem.CreatedDateTime);
                Console.WriteLine("Create By= " + fileItem.CreatedBy.User.DisplayName);
                Console.WriteLine("WebUrl=" + fileItem.WebUrl); //该URL不能给其他用户直接访问，其他用户需要登录Offic365才能访问。
                Console.WriteLine("Size=" + fileItem.Size);
                Console.WriteLine();

                //根据ID获取文件信息并下载该文件
                //ID:01C4GJWGIXFT6RLTDCSJCYZEVPPFMYLW6M
                var remoteFileStream = await graphClient
                        .Drives["b!h2paA_qdHkWMFMnaxQ6505czytzKNNJBhQafMxIUPLnZtSSEzeB5Q6gCbA9lBz0K"]
                        .Items["01C4GJWGIXFT6RLTDCSJCYZEVPPFMYLW6M"].Content.GetAsync();
                
                string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                using (FileStream localFileStream = File.Create(Path.Combine(userFolder,"Pictures/01.jpg")))
                {
                    CopyStream(remoteFileStream, localFileStream);
                }
                Console.WriteLine();

                //上传文件

            }
            //读取ODataError错误中的详细信息才能了解为什么请求失败。
            catch (ODataError odataError)
            {
                Console.WriteLine(string.Format("Error Code: {0}", odataError.Error.Code));
                Console.WriteLine(string.Format("Error Message: {0}", odataError.Error.Message));
                throw;
            }
            catch(Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
                throw;
            }
        }

        //获取本地文件夹中储存的密钥，防止密钥泄露到GitHub上。
        public static string getJsonConfig(string key)
        {
            string userFolder=Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string filePath = Path.Combine(userFolder, "OneDrive\\脚本代码\\AccountSecrets.json");
            IConfigurationBuilder builder = new ConfigurationBuilder().AddJsonFile(filePath);
            IConfiguration config = builder.Build();
            var section= config.GetSection("MsGraphSdkApp1").GetSection(key);
            return section.Value;
        }

        //用于把OneDrive的文件流写入到本地文件
        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }
    }
}