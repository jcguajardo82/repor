using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EN
{
    public class SantanderRes_Model
    {
        public int Id_Num_Transaccion { get; set; }
        public string MerchTxnRef { get; set; }
        public string Desc_RespCode { get; set; }
        public string Cve_FolCPagos { get; set; }
        public string Cve_Autz { get; set; }
        public string Cve_RespCode { get; set; }
        public string Cve_CdError { get; set; }
        public string Desc_NbError { get; set; }
        public string Cve_Hora { get; set; }
        public string Cve_Fecha { get; set; }
        public string Desc_NbEmp { get; set; }
        public string Desc_NbMerchant { get; set; }
        public string Cve_TipoFP { get; set; }
        public string Desc_TpOper { get; set; }
        public string Desc_CcNomFP { get; set; }
        public string Num_FP { get; set; }
        public string Importe { get; set; }
        public string Cve_IdURL { get; set; }
        public string Cve_TokenFP { get; set; }
        public string Cve_eMailCte { get; set; }
    }

    public class Mensaje_RespuestaDB
    {
        public bool Bit_Error { get; set; }
        public string Desc_MensajeError { get; set; }    
    }

    public class MensajeRespuestaWS
    {
        public bool Bit_Error { get; set; }
        public string Peticion { get; set; }
        public string Desc { get; set; }
    }

    public class TestConn
    {
        public string Fecha { get; set; }
        public bool TestConnection { get; set; }
    }

    public class OrdenCancelacion_Model
    {
        public string Id_Num_Transaccion { get; set; }
        public string Id_Num_Orden { get; set; }
        public string Imp_CierrePreventa { get; set; }
    }
}
