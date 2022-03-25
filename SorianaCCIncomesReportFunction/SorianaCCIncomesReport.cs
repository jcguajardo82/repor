using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;

using SorianaCCIncomesReportFunction.Services;

namespace SorianaIncomesCCReporterFunction
{
    public class SorianaCCIncomesReport
    {     
        [FunctionName("SorianaCCIncomesReport")]
        //public static void Run([TimerTrigger(" 0 00 21 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest request)
        {
            try
            {
                CCIncomesReportService Report = new CCIncomesReportService();
                Report.GenerarReportes();

                return new OkObjectResult("");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
