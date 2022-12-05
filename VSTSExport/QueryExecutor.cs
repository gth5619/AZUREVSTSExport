using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;

namespace VSTSExport
{
    public class QueryExecutor
    {
        private readonly Uri uri;
        private readonly string personalAccessToken;


        public QueryExecutor(string orgName, string personalAccessToken)
        {
            this.uri = new Uri(ConfigurationManager.AppSettings["uri"] + orgName);
            this.personalAccessToken = personalAccessToken;
        }

        /// <summary>
        ///     Execute a WIQL (Work Item Query Language) query to return a list of open bugs.
        /// </summary>
        /// <param name="project">The name of your project within your organization.</param>
        /// <returns>A list of <see cref="WorkItem"/> objects representing all the open bugs.</returns>
        public async Task<IList<WorkItem>> QueryOpenBugs(string project)
        {
            var credentials = new VssBasicCredential(string.Empty, this.personalAccessToken);
            string AreaPath = ConfigurationManager.AppSettings["AreaPath"];
            string IterationPath = ConfigurationManager.AppSettings["IterationPath"];
            string IterationPath1 = ConfigurationManager.AppSettings["IterationPath1"];
            string AssignedTo = ConfigurationManager.AppSettings["AssignedTo"];
            // create a wiql object and build our query
            var wiql = new Wiql()
            {
                // NOTE: Even if other columns are specified, only the ID & URL are available in the WorkItemReference
                Query = "Select [Id] " +
                        "From WorkItems " +
                        "Where ([Work Item Type] =  'Product Backlog Item'  or  [Work Item Type] =  'Bug' or  [Work Item Type] =  'Defect' or  [Work Item Type] =  'Enabler')" +
                        "And [System.TeamProject] = '" + project + "' " +
                        "And [System.AssignedTo] = '" + AssignedTo + "' " +
                        "And [System.AreaPath] = '" + AreaPath + "' " +
                        "And ( [System.IterationPath] = '" + IterationPath + "' " +
                        "or [System.IterationPath] = '" + IterationPath1 + "') " +
                        // "And [System.State] <> 'Closed' " +
                        "Order By [State] Asc, [Changed Date] Desc",
            };
            // Defect Task  enabler

            try
            {
                // create instance of work item tracking http client
                using (var httpClient = new WorkItemTrackingHttpClient(this.uri, credentials))
                {
                    // execute the query to get the list of work items in the results
                    var result = await httpClient.QueryByWiqlAsync(wiql).ConfigureAwait(false);

                    var ids = result.WorkItems.Select(item => item.Id).ToArray();

                    // some error handling
                    if (ids.Length == 0)
                    {
                        return Array.Empty<WorkItem>();
                    }

                    // build a list of the fields we want to see
                    var fields = new[] { "System.Id", "System.Title", "System.State", "System.WorkItemType" };

                    List<WorkItem> workItems = new List<WorkItem>();

                    decimal pageing = Math.Ceiling((decimal)ids.Count() / 200);

                    for (int i = 0; i < pageing; i++)
                    {
                        var input = ids.Skip(i * 200).Take(200);
                        workItems.AddRange(await httpClient.GetWorkItemsAsync(input, fields, result.AsOf).ConfigureAwait(false));
                    }

                    return workItems;
                    // get work items for the ids found in query
                    // return await httpClient.GetWorkItemsAsync(ids, fields, result.AsOf).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        ///     Execute a WIQL (Work Item Query Language) query to print a list of open bugs.
        /// </summary>
        /// <param name="project">The name of your project within your organization.</param>
        /// <returns>An async task.</returns>
        public async Task PrintOpenBugsAsync(string project)
        {
            StringBuilder sbVSTSReport = new StringBuilder();
            try
            {
                var workItems = await this.QueryOpenBugs(project).ConfigureAwait(false);
                Console.WriteLine("Query Results: {0} items found", workItems.Count);
                List<string[]> lst = new List<string[]>();

                string[] arr = new string[5] { "SNO", "ID", "WorkItemType", "Work Item Title", "Status" };
                lst.Add(arr);

                // loop though work items and write to console
                int Counter = 1;
                foreach (var workItem in workItems)
                {
                    arr = new string[5] { Counter.ToString(), workItem.Id.ToString(), workItem.Fields["System.WorkItemType"].ToString(), workItem.Fields["System.Title"].ToString(), workItem.Fields["System.State"].ToString() };
                    lst.Add(arr);
                    Counter++;
                }

                File.WriteAllLines(@"C:\temp\Tosca\Report Analysis\report " + DateTime.Now.ToString("dd-MM-yyyy  hh-mm-ss") + ".csv", lst.Select(x => string.Join(",", x)));
                Console.ReadKey();
            }
            catch (Exception ex)
            {

            }
        }
    }
}