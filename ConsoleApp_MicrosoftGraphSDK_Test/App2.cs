using Azure.Identity;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.Security;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;

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
            var userCredential = new UsernamePasswordCredential(username, password, tenantId, clientId, options);

            GraphServiceClient graphClient = new GraphServiceClient(userCredential, scopes);
            try
            {
                //await deleteFileById(graphClient);
                //await getRootChildrenListAsync(graphClient);
                //await uploadFileToFolderById(graphClient);
                await uploadBigFile(graphClient, @"C:\Users\jinge\Pictures\ImportedPhoto_1739953924163.jpg");

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

        /// <summary>
        /// 根据文件夹ID，上传文件到指定文件夹
        /// </summary>
        /// <param name="graphClient"></param>
        /// <returns></returns>
        public async Task uploadFileToFolderById(GraphServiceClient graphClient)
        {
            string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            using (FileStream localFileStream = File.OpenRead(Path.Combine(userFolder, "Pictures/TestForUpload-1.jpg")))
            {
                var uploadedItem = await graphClient
                    .Drives["b!giXh4U7g8UyCVVetRQoTyNCYhPEGiWpFimImDbiPA913eE1reuE0TqFkDQLDmZny"]
                    .Items["01ZVDGC6N6Y2GOVW7725BZO354PWSELRRZ"]
                    .ItemWithPath("TestForUpload-1.jpg") //别忘了指定上传之后的文件名
                    .Content.PutAsync(localFileStream);
                Console.WriteLine(uploadedItem.Id);
                Console.WriteLine(uploadedItem.WebUrl);
            }
        }

        /// <summary>
        /// 上传大文件，支持断点续传。
        /// 官方文档：https://learn.microsoft.com/en-us/graph/sdks/large-file-upload?tabs=csharp
        /// </summary>
        /// <param name="graphClient"></param>
        /// <returns></returns>
        public async Task uploadBigFile(GraphServiceClient graphClient, string filePath)
        {
            //string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            //string filePath = Path.Combine(userFolder, "Pictures/3P2A9987.JPG");
            string fileName = Path.GetFileName(filePath);
            //这样写using，则fileStream对象会在代码块的末尾被释放，也就是当前方法的末尾。
            using var fileStream = File.OpenRead(filePath);

            // Use properties to specify the conflict behavior
            // using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
            var uploadSessionRequestBody = new DriveUpload.CreateUploadSessionPostRequestBody
            {
                Item = new DriveItemUploadableProperties
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" },
                    },
                },
            };

            // Create the upload session
            // itemPath does not need to be a path to an existing item
            var myDrive = await graphClient.Me.Drive.GetAsync();
            var uploadSession = await graphClient.Drives[myDrive?.Id]
                .Items["root"]
                .ItemWithPath(fileName)
                .CreateUploadSession
                .PostAsync(uploadSessionRequestBody);

            // Max slice size must be a multiple of 320 KiB
            int maxSliceSize = 320 * 1024;
            var fileUploadTask = new LargeFileUploadTask<DriveItem>(
                uploadSession, fileStream, maxSliceSize, graphClient.RequestAdapter);

            var totalLength = fileStream.Length;
            // Create a callback that is invoked after each slice is uploaded
            IProgress<long> progress = new Progress<long>(prog =>
            {
                Console.WriteLine($"Uploaded {prog} bytes of {totalLength} bytes");
            });

            try
            {
                // Upload the file
                var uploadResult = await fileUploadTask.UploadAsync(progress);

                Console.WriteLine(uploadResult.UploadSucceeded ?
                    $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                    "Upload failed");
            }
            catch (ODataError ex)
            {
                Console.WriteLine($"Error uploading: {ex.Error?.Message}");
            }
        }
    }
}
