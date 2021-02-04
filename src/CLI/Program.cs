using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;

namespace tools
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var switchMappings = new Dictionary<string, string>()
           {
               { "-c", "connectionString" },
               { "-m", "method" },
               { "-p", "path" },
               { "--connectionString", "connectionString" },
               { "--method", "method" },
               { "--path", "path" },
               { "--alt6", "key6" }
           };
            var builder = new ConfigurationBuilder();
            builder.AddCommandLine(args, switchMappings);

            var config = builder.Build();

            var connectionString = config["connectionString"];
            if (string.IsNullOrEmpty(connectionString))
            {
                throw new Exception("Miss connection string.");
            }

            var method = config["method"];
            if (string.IsNullOrEmpty(method))
            {
                throw new Exception("Miss method.");
            }

            var path = config["path"];
            if (string.IsNullOrEmpty(path))
            {
                throw new Exception("Miss path.");
            }

            CrmServiceClient svc = new CrmServiceClient(connectionString);

            // Verify that you are connected.
            if (svc != null && svc.IsReady)
            {
                switch (method.ToLower())
                {
                    case "exportwebtemplates":
                        ExportWebTemplates(svc, path);
                        break;

                    case "importwebtemplates":
                        ImportWebTemplates(svc, path);
                        break;

                    case "":
                        break;

                    default:
                        return;
                }
            }
            else
            {
                ShowCrmErrors(svc);

                return;
            }
        }

        private static void ShowCrmErrors(CrmServiceClient svc)
        {
            // Display the last error.
            Console.WriteLine("An error occurred: {0}", svc.LastCrmError);

            // Display the last exception message if any.
            Console.WriteLine(svc.LastCrmException.Source);
            Console.WriteLine(svc.LastCrmException.StackTrace);

            throw new Exception(svc.LastCrmException.Message);
        }

        private static void ImportWebTemplates(CrmServiceClient svc, string sourcePath)
        {
            var files = new DirectoryInfo(sourcePath).GetFiles();
            foreach (var file in files)
            {
                if (file.Extension != ".html")
                {
                    continue;
                }
                var raw = File.ReadAllText(file.FullName);

                var id = raw.Split(' ')[1];
                var source = raw.Substring(47);

                var updateData = new Dictionary<string, CrmDataTypeWrapper>();
                updateData.Add("adx_source", new CrmDataTypeWrapper(source, CrmFieldType.String));
                var updateAccountStatus = svc.UpdateEntity("adx_webtemplate", "adx_webtemplateid", Guid.Parse(id), updateData);

                if (updateAccountStatus != true)
                {
                    ShowCrmErrors(svc);
                }
            }
        }

        private static void ExportWebTemplates(CrmServiceClient svc, string exportPath = "src/webtemplates")
        {
            string fetchXML =
                @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' returntotalrecordcount='true' >
                    <entity name='adx_webtemplate'>
                        <attribute name='adx_name' />
                        <attribute name='adx_webtemplateid' />
                        <attribute name='adx_source' />
                    </entity>
                 </fetch>";
            var queryResult = svc.GetEntityDataByFetchSearchEC(fetchXML);
            if (queryResult != null)
            {
                foreach (var entity in queryResult.Entities)
                {
                    var name = entity.Attributes["adx_name"];
                    var id = entity.Attributes["adx_webtemplateid"];
                    if (entity.Attributes.ContainsKey("adx_source"))
                    {
                        Console.WriteLine($"Working on WebTemplate: {name}");

                        var headline = $"<!-- {id} -->";
                        var source = headline + Environment.NewLine + entity.Attributes["adx_source"];
                        Directory.CreateDirectory(exportPath);
                        File.WriteAllText($"{exportPath}/{name}.html", source);
                    }
                }
                Console.WriteLine(string.Format("WebTemplates Records Count : {0}", queryResult.TotalRecordCount));
            }
        }
    }
}