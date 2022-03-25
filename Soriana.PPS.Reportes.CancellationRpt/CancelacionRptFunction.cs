using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Soriana.PPS.Reportes.CancellationRpt.Services;

namespace Soriana.PPS.Reportes.CancellationRpt
{
    public class CancelacionRptFunction
    {
        [FunctionName("CancelacionesPPSReport")]
        //public static void Run([TimerTrigger(" 0 30 21 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest request)
        {
            try
            {
                CancelacionesRptService Report = new CancelacionesRptService();
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
