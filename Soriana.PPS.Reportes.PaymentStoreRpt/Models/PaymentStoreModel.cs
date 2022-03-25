using System;
using System.Collections.Generic;
using System.Text;

namespace Soriana.PPS.Reportes.PaymentStoreRpt.Models
{
    public class PaymentStoreModel
    {
        public string OrderReferenceNumber { get; set; }
        public string OrderAmount { get; set; }
        public string LineaCaptura { get; set; }
        public string Estatus { get; set; }
        public string CreatedDate { get; set; }
    }

    public class PaymentStoreModelResponse
    {
        public string OrdenID { get; set; } = "";
        public string IDtransaccion { get; set; } = "";
        public string FechaCreacion { get; set; } = "";
        public string HoraCreacion { get; set; } = "";
        public string NoAfiliacion { get; set; } = "";
        public string Adquirente { get; set; } = "";
        public string Catalogo { get; set; } = "";
        public string TipoEntrega { get; set; } = "";
        public string CanalCompra { get; set; } = "";
        public string FormaPago { get; set; } = "";
        public string NoTienda { get; set; } = "";
        public string NombreTienda { get; set; } = "";
        public string noCajero { get; set; } = "";
        //fecha creacion de la orden
        //hora de creacion de la orden
        public string montoPagado { get; set; } = "";
        public string precioTotalOrden { get; set; } = "";
        public string fechaPago { get; set; } = "";
        public string formaPago { get; set; } = "";
        public string Banco { get; set; } = "";
        public string NoAutorizacion { get; set; } = "";
        public string BIN { get; set; } = "";
        public string Sufijo { get; set; } = "";
        public string TipoTarjeta { get; set; } = "";
        public string Marca { get; set; } = "";
        public string formato { get; set; } = "";
        public string ciudadEstatusPago { get; set; } = "";
        public string CostoEnvio { get; set; } = "";
        public string MSI { get; set; } = "";
        public string PuntosAplicados { get; set; } = "";
        public string PromocionesAplicadas { get; set; } = "";
        public string NombrePersonaRegistrada { get; set; } = "";
        public string Apellido_P { get; set; } = "";
        public string Apellido_M { get; set; } = "";
        public string NoTarjetalealtad { get; set; } = "";
        public string Correo { get; set; } = "";
        public string EstatusOrden { get; set; } = "";
        public string EstatusEnvío { get; set; } = "";
        public string almacenSurtio { get; set; } = "";
        public string loyalty { get; set; } = "";
        public string CreteOrderStore { get; set; } = "";
        public string HoraOrderStore { get; set; } = "";
    }

    public class JsonRespoonseModel
    {
        public string orderReferenceNumber { get; set; }
        public string orderAmount { get; set; }
        public string orderDateTime { get; set; }
        public string orderSaleChannel { get; set; }
        public string paymentType { get; set; }
        public string paymentProcessor { get; set; }
        public string paymentToken { get; set; }
        public string customerEmail { get; set; }
        public string customerCity { get; set; }
        public string customerState { get; set; }
        public string customerLoyaltyCardId { get; set; }
        public string customerLoyaltyRedeemElectronicMoney { get; set; }
        public string customerLoyaltyRedeemPoints { get; set; }
        public string customerLoyaltyRedeemMoney { get; set; }
        public string TransactionAuthorizationId { get; set; }
        public string TransactionStatus { get; set; }
        public string TransactionReferenceID { get; set; }
        public string AffiliationType { get; set; }
        public string IsAuthenticated { get; set; }
        public string IsAuthorized { get; set; }
        public string Apply3DS { get; set; }
        public string MerchandiseType { get; set; }
        public List<shipments> shipments { get; set; }
    }

    public class shipments
    {
        public string shippingStoreId { get; set; }
        public string shippingDeliveryDesc { get; set; }
        public string shippingPaymentImport { get; set; }
        public string shippingPaymentInstallments { get; set; }
        public string shippingReferenceNumber { get; set; }
        public string shippingFirstName { get; set; }
        public string shippingLastName { get; set; }
        public List<items> Items { get; set; }
    }

    public class items
    {
        public string shippingItemId { get; set; } = "";
        public string shippingItemName { get; set; } = "";
        public string shippingItemCategory { get; set; } = "";
        public string ShippingItemTotal { get; set; } = "";
    }

    public class AprobacionesMarcas
    {
        public string canalCompra { get; set; } = "";
        public string tipoTarjeta { get; set; } = "";
        public string marca { get; set; } = "";
        public string totalOrdenes { get; set; } = "";
        public string ordenesAprobadas { get; set; } = "";
        public string porcentajeAprobacion { get; set; } = "";
        public string ordenesRechazadas { get; set; } = "";
        public string porcentajeRechazo { get; set; } = "";
        public string monto { get; set; } = "";
    }

    public class Aprobaciones
    {
        public string Id_Num_Orden { get; set; }
        public string id_num_apl { get; set; }
        public string id_num_formapago { get; set; }
        public string nom_pagOrig { get; set; }
        public string tipoOrden { get; set; }
        public string imp_preciounit { get; set; }
    }
}
