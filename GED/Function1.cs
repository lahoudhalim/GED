using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
//using Microsoft.AspNetCore.Http;


namespace GED
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static HttpResponseMessage Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, ILogger log)
        {
            string httpPostData = string.Empty;

            var reader = new StreamReader(req.Content.ReadAsStreamAsync().Result);
            if (reader != null)
            {
                httpPostData = reader.ReadToEnd();
            }



            if (!string.IsNullOrWhiteSpace(httpPostData))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(httpPostData);



                // Read data from event payload.
                string webUrl = xmlDoc.GetElementsByTagName("WebUrl")[0].InnerText;
                int listItemId = int.Parse(xmlDoc.GetElementsByTagName("ListItemId")[0].InnerText);
            }
            ClientContext ctx = SPConnection.GetSPOLContext("https://ghtpdfr.sharepoint.com/sites/ged");
            SPFunctions.getStatus(ctx, 2, "Library%20GED");

            return req.CreateResponse(HttpStatusCode.OK, "Succeeded");

        }
    }

}