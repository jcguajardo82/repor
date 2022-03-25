using System;
using Microsoft.Azure.WebJobs;

using Soriana.PPS.Reportes.CanalCompraRpt.Services;
using Microsoft.Extensions.Logging;

namespace Soriana.PPS.Reportes.CanalCompraRpt
{
    public class Function1
    {
        [FunctionName("CanalCompraPPSReport")]
        //public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest request)
        public static void Run([TimerTrigger(" 0 45 21 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        {
            try
            {
                CanalCompraRptService Report = new CanalCompraRptService();
                Report.GenerarReportes();

                //return new OkObjectResult("");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
