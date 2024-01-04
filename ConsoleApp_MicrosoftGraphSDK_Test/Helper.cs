using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp_MicrosoftGraphSDK_Test
{
    public class Helper
    {
        //获取本地文件夹中储存的密钥，防止密钥泄露到GitHub上。
        public static string GetJsonConfig(string appName, string key)
        {
            string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string filePath = Path.Combine(userFolder, "OneDrive\\脚本代码\\AccountSecrets.json");
            IConfigurationBuilder builder = new ConfigurationBuilder().AddJsonFile(filePath);
            IConfiguration config = builder.Build();
            var section = config.GetSection(appName).GetSection(key);
            return section.Value;
        }

        //用于把OneDrive的文件流复制到本地文件流，用于写入文件
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
