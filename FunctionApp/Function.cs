using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;

namespace FunctionApp
{
    public static class Function
    {
        [FunctionName("Function")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("New modifications was detected.");
            // Get request body
            string httpPostData = string.Empty;
            var content = await req.Content.ReadAsStringAsync();
            var reader = new StreamReader(content);
            if (reader != null)
            {
                httpPostData = reader.ReadToEnd();
            }
            XmlDocument xmlDoc = new XmlDocument();
            if (!string.IsNullOrWhiteSpace(httpPostData))
            {
                xmlDoc.LoadXml(httpPostData);
                //print Xml to Azure Function log
                using (var stringWriter = new StringWriter())
                using (var xmlTextWriter = XmlWriter.Create(stringWriter))
                {
                    xmlDoc.WriteTo(xmlTextWriter);
                    xmlTextWriter.Flush();
                    log.Info(stringWriter.GetStringBuilder().ToString());
                }

                string webUrl = xmlDoc.GetElementsByTagName("WebUrl")[0].InnerText;
                string listTitle = xmlDoc.GetElementsByTagName("ListTitle")[0].InnerText;
                string itemId = xmlDoc.GetElementsByTagName("ListItemId")[0].InnerText;

                string listTitleLog = "Log";

                // Input Parameters  
                string userName = "***********";
                string password = "******";

                string columnTitle = "Title";

                string objectif = "New element";

                // PnP component to set context  
                OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();

                try
                {
                    using (var clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(webUrl, userName, password))
                    {
                        // Retrieve list item
                        clientContext.Load(clientContext.Web);
                        clientContext.ExecuteQuery();
                        List list = clientContext.Web.GetListByTitle(listTitleLog);
                        clientContext.Load(list);
                        clientContext.ExecuteQuery();

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem oListItem;
                        oListItem = list.AddItem(itemCreateInfo);
                        oListItem[columnTitle] = objectif;
                        oListItem.Update();
                        clientContext.Load(list);
                        clientContext.ExecuteQuery();
                    }
                }
                catch (Exception ex)
                {
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Error Message: " + ex.Message);
                }
            }

            return req.CreateResponse(HttpStatusCode.OK, "L'azure function est terminé");
        }
    }
}
