using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EN
{
    public class TEF_File
    {
        public TEF_Encabezado Encabezado { get; set; }
        public TEF_EncabezadoAfiliacion EncabezadoAfiliacion { get; set; }
        public List<TEF_DetalleTransaccion> DetalleTransaccion { get; set; }
        public TEF_DetalleAfiliacion TrailerAfiliacion { get; set; }
        public TEF_DetalleAfiliacion TrailerArchivo { get; set; }
    }

    public class TEF_Encabezado
    {
        public string NoRegistroEnc { get; set; }
        public string TipoRegistroEnc { get; set; }
        public string FechaTransmisión { get; set; }
        public string NombreArchivo { get; set; }
    }

    public class TEF_EncabezadoAfiliacion
    {
        public string NoRegistroAfil { get; set; }
        public string TipoRegistroAfil { get; set; }
        public string AfiliacionBancaria { get; set; }
        public string RellenoEnc { get; set; }
    }

    public class TEF_DetalleTransaccion
    {
        public string NoRegistroTran { get; set; }
        public string TipoRegistroTran { get; set; }
        public string FechaTransaccion { get; set; }
        public string TipoTransaccion { get; set; }
        public string ReferenciaTransaccion { get; set; }
        public string ReferenciaTarjeta { get; set; }
        public string Autorizacion { get; set; }
        public string Importe { get; set; }
        public string HoraTransaccion { get; set; }
        public string ReferenciaPromoFinanciamiento { get; set; }
        public string CantidadPlazoDiferidos { get; set; }
        public string Relleno { get; set; }
    }

    public class TEF_DetalleAfiliacion
    {
        public string NoRegistroAfiliacion { get; set; }
        public string TipoRegistroAfiliacion { get; set; }
        public string AfiliacionBancaria { get; set; }
        public string CantidadControlVentas { get; set; }
        public string TotalControlVentas { get; set; }
        public string CantidadControlDevoluciones { get; set; }
        public string TotalControlDevoluciones { get; set; }
        public string CantidadControlCancelaciones { get; set; }
        public string TotalControlCancelaciones { get; set; }
        public string Relleno { get; set; }

    }
}
