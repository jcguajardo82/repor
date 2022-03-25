using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

using Soriana.PPS.Reportes.PaymentStoreRpt.Services;

namespace Soriana.PPS.Reportes.PaymentStoreRpt
{
    public class PaymentStoreRptFunction
    {
        [FunctionName("PaymentStorePPSReport")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest request) // 0 0 10 * * *
        //public static void Run([TimerTrigger(" 0 00 22 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        {
            try
            {
                PaymentStoreRptService Report = new PaymentStoreRptService();
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
