using System.Threading.Tasks;
using System.Configuration;

namespace VSTSExport
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            QueryExecutor obje = new QueryExecutor(ConfigurationManager.AppSettings["boardName"], ConfigurationManager.AppSettings["PAT"]);
            await obje.PrintOpenBugsAsync(ConfigurationManager.AppSettings["project"]); 
        }
    }
}