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
        //rf2j账号（新账号）
        static string AppNameKey1 = "MsGraphSdkApp1";
        //xltrg账号（旧账号）
        static string AppNameKey2 = "MsGraphSdkApp_xltrg";

        static async Task Main(string[] args)
        {
            //App1 app = new App1(AppNameKey1);
            //await app.RunAsync();

            App2 app = new App2(AppNameKey2);
            await app.RunAsync();
        }


    }
}