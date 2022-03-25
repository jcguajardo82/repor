using System;
using System.Collections.Generic;
using System.Text;

namespace SorianaCCIncomesReportFunction.Models
{
    public class OrdersAppModel
    {
        /*public string Id_Num_Orden { get; set; }
        public string Id_num_un { get; set; }
        public string Fec_movto { get; set; }
        public string id_num_cte { get; set; }
        public string id_num_apl { get; set; }
        public string id_num_formapago { get; set; }
        public string id_num_srventrega { get; set; }
        public string tipoOrden { get; set; }*/

        public string OrderNo { get; set; } = "";
        public string StatusUe { get; set; } = "";
        public string CreatedBy { get; set; } = "";
        public string Id_Num_Apl { get; set; } = "";
        public string MethodPayment { get; set; } = "";
        public string DeliveryType { get; set; } = "";
    }
}
