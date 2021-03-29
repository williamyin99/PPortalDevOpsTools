using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
               { "--path", "path" }
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

                    case "exportwebfiles":
                        ExportWebFiles(svc, path);
                        break;

                    case "exportcontentsnippets":
                        ExportContentSnippets(svc, path);
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

        private static void ExportWebTemplates(CrmServiceClient svc, string exportPath = "src/webtemplates")
        {
            Directory.CreateDirectory(exportPath);
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
                    var name = entity.Attributes["adx_name"].ToString();
                    var id = entity.Attributes["adx_webtemplateid"];
                    if (entity.Attributes.ContainsKey("adx_source"))
                    {
                        Console.WriteLine($"Working on WebTemplate: {name}");

                        var headline = $"<!-- {id} -->";
                        var source = entity.Attributes["adx_source"].ToString();
                        if (!source.Contains(headline))
                        {
                            source = headline + Environment.NewLine + source;
                        }

                        if (name.Contains('/'))
                        {
                            name = name.Replace('/', '-');
                        }
                        File.WriteAllText($"{exportPath}/{name}.html", source);
                    }
                }
                Console.WriteLine(string.Format("WebTemplates Records Count : {0}", queryResult.TotalRecordCount));
            }
        }

        private static void ExportWebFiles(CrmServiceClient svc, string exportPath, string filterId = "")
        {
            var filter = "";
            if (!string.IsNullOrEmpty(filterId))
            {
                filter = $"";
            }

            string fetchXML =
                $@"
                <fetch>
                  <entity name='annotation' >
                    <attribute name='documentbody' />
                    <attribute name='objectid' />
                    <link-entity name='adx_webfile' from='adx_webfileid' to='objectid' link-type='inner' alias='f' >
                      <attribute name='adx_webfileid' />
                      <attribute name='adx_name' />
                      <attribute name='adx_partialurl' />
                      <attribute name='adx_websiteid' />
                        {filter}
                    </link-entity>
                  </entity>
                </fetch>";

            var queryResult = svc.GetEntityDataByFetchSearchEC(fetchXML);
            if (queryResult != null)
            {
                Directory.Delete(exportPath, true);
                Directory.CreateDirectory(exportPath);
                var mapList = new List<dynamic>();
                foreach (var entity in queryResult.Entities)
                {
                    var fileName = ((entity.Attributes["f.adx_partialurl"] as AliasedValue).Value as string).Split('/').Last();
                    var id = (entity.Attributes["f.adx_webfileid"] as AliasedValue).Value;
                    var str = entity.Attributes["documentbody"] as string;
                    var docBody = Convert.FromBase64String(str);
                    if (docBody != null)
                    {
                        Console.WriteLine($"Working on WebFile: {fileName}");
                        using (var fileStream = new FileStream($"{exportPath}/{fileName}", FileMode.Create))
                        {
                            fileStream.Write(docBody, 0, docBody.Length);
                            mapList.Add(new { FileName = fileName, OriginName = (entity.Attributes["f.adx_name"] as AliasedValue).Value, WebFileId = id });
                        }
                    }
                }
                File.WriteAllText(exportPath + "/mapList.json", JsonConvert.SerializeObject(mapList));
                Console.WriteLine(string.Format("WebFiles Records Count : {0}", queryResult.TotalRecordCount));
            }
        }

        private static void ExportContentSnippets(CrmServiceClient svc, string exportPath)
        {
            Directory.CreateDirectory(exportPath);
            string fetchXML =
                @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' returntotalrecordcount='true' >
                    <entity name='adx_contentsnippet'>
                        <attribute name='adx_name' />
                        <attribute name='adx_contentsnippetid' />
                        <attribute name='adx_value' />
                        <attribute name='adx_type' />
                    </entity>
                 </fetch>";

            var queryResult = svc.GetEntityDataByFetchSearchEC(fetchXML);
            if (queryResult != null)
            {
                foreach (var entity in queryResult.Entities)
                {
                    var name = entity.Attributes["adx_name"].ToString();
                    var id = entity.Attributes["adx_contentsnippetid"];
                    if (entity.Attributes.ContainsKey("adx_value"))
                    {
                        Console.WriteLine($"Working on ContentSnippets: {name}");

                        var headline = $"<!-- {id} -->";
                        var source = entity.Attributes["adx_value"].ToString();
                        if (!source.Contains(headline))
                        {
                            source = headline + Environment.NewLine + source;
                        }

                        if (name.Contains('/'))
                        {
                            name = name.Replace('/', '-');
                        }
                        File.WriteAllText($"{exportPath}/{name}.html", source);
                    }
                }
                Console.WriteLine(string.Format("ContentSnippets Records Count : {0}", queryResult.TotalRecordCount));
            }
        }
    }
}