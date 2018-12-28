using CommandLine;
using Serilog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Support
{
    public class Options
    {
        [Option(
            "address",
            Required = true,
            HelpText = "Service address, ex https://xxx.visualstudio.com")]
        public String ServiceAddress { get; set; }

        [Option(
            "tokenfile",
            Required = true,
            HelpText = "File that contains access token to perform the migration, useful if you have token stored in a file in a protected directory accessible only by the account that is runnign the service")]
        public String AccessTokenFile { get; set; }

        [Option(
            "teamproject",
            Required = true,
            HelpText = "Name of the teamproject")]
        public String TeamProject { get; set; }

        internal string GetAccessToken()
        {
            if (!String.IsNullOrEmpty(AccessTokenFile))
            {
                if (!File.Exists(AccessTokenFile))
                {
                    Log.Logger.Error("Unable to find AccessTokenFile {AccessTokenFile} with access token ", AccessTokenFile);
                    throw new ConfigurationErrorsException("Unable to find AccessTokenFile");
                }
                return File.ReadAllText(AccessTokenFile);
            }

            Log.Logger.Error("Access token should be specified, you can use any of the supported method (Access Token File)", AccessTokenFile);
            throw new ConfigurationErrorsException("No Access Token specified");
        }
    }
}
