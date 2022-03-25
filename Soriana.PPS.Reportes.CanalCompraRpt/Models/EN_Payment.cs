using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EN
{
    public class StorePaymentRequest
    {
        public string Cve_Operacion { get; set; }
        public List<StorePayment> PaymentsStore { get; set; }
        public SaveStorePaymentRequest SavePayment { get; set; }
    }

    public class StorePayment
    {
        public string Id_Cve_RefPago { get; set; }
        public string ImportePago { get; set; }
        public string Sucursal { get; set; }
        public string Cajero { get; set; }
        public string Caja { get; set; }
        public string Transaccion { get; set; }
        public string FechaPago { get; set; }
    }

    public class SaveStorePaymentRequest
    {
        public string Id_Cve_RefPago { get; set; }
        public string ClientID { get; set; }
        public string OrderAmount { get; set; }
    }
}
