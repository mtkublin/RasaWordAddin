using System.Collections.Generic;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        public void GetProjsList(RestClient client, RibbonDropDown TestProjectDropDown)
        {
            var Request = new RestRequest("api/projects", Method.GET);
            IRestResponse Response = client.Execute(Request);
            string JSONresultDoc = Response.Content.ToString();

            List<string> ProjList = JsonConvert.DeserializeObject<List<string>>(JSONresultDoc);

            foreach (string itemName in ProjList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                TestProjectDropDown.Items.Add(item);
            }
        }

        public void TESTGetProjsList(RestClient client, RibbonDropDown TestProjectDropDown, RibbonDropDown TestModelDropDown)
        {
            GetProjsList(client, TestProjectDropDown);
            
            string ProjectName = TestProjectDropDown.SelectedItem.Label;
            Globals.ThisAddIn.TESTGetModelsList(client, ProjectName, TestModelDropDown);
        }

        public void GetModelsList(RestClient client, string ProjectName, string NewModelName)
        {
            var newRequest = new RestRequest("api/models/{project}", Method.GET);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            IRestResponse newResponse = client.Execute(newRequest);
            string newJSONresultDoc = newResponse.Content.ToString();

            List<string> ModelList = JsonConvert.DeserializeObject<List<string>>(newJSONresultDoc);

            //if (NewModelName is in ModelList)
            //{

            //}
        }

        public void TESTGetModelsList(RestClient client, string ProjectName, RibbonDropDown TestModelDropDown)
        {
            var newRequest = new RestRequest("api/models/{project}", Method.GET);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            IRestResponse newResponse = client.Execute(newRequest);
            string newJSONresultDoc = newResponse.Content.ToString();
            List<string> ModelList = JsonConvert.DeserializeObject<List<string>>(newJSONresultDoc);

            TestModelDropDown.Items.Clear();

            foreach (string itemName in ModelList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = itemName;
                TestModelDropDown.Items.Add(item);
            }

            string ModelName = TestModelDropDown.SelectedItem.Label;
            Globals.ThisAddIn.UpdateInterpreter(client, ProjectName, ModelName);
        }

        public void UpdateInterpreter(RestClient client, string ProjectName, string ModelName)
        {
            var newRequest = new RestRequest("api/interpreter/{project}/{model}", Method.POST);
            newRequest.AddParameter("project", ProjectName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("project", ProjectName);
            newRequest.AddParameter("model", ModelName, ParameterType.UrlSegment);
            newRequest.AddUrlSegment("model", ModelName);
            IRestResponse newResponse = client.Execute(newRequest);
        }
    }
}
