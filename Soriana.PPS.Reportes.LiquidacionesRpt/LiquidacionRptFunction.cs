using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;

using Soriana.PPS.Reportes.LiquidacionesRpt.Services;

namespace Soriana.PPS.Reportes.LiquidacionesRpt
{
    public class LiquidacionRptFunction
    {
        [FunctionName("LiquidacionPPSReport")]
        public static void Run([TimerTrigger(" 0 15 21 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)   
        {
            try
            {
                //public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest request)
                LiquidacionesRptService Report = new LiquidacionesRptService();
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
