using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using Newtonsoft.Json;
using Renci.SshNet;
using Soriana.PPS.Common.Constants;

using Soriana.PPS.Reportes.ReportesMasivos.Services;

using EN;
using System.Collections.Generic;
using System.Text;
using Jitbit.Utils;
using SorianaCCIncomesReportFunction.Models;
using System.Text.RegularExpressions;

namespace Soriana.PPS.Reportes.CancellationRpt.Services
{
    public class CancelacionesRptService
    {
        DA_Reportes Reportes = new DA_Reportes();

        public void GenerarReportes()
        {
            var Creditos = BL_Creditos();
            CreateExcel_Creditos(Creditos);
        }

        #region Datos Reportes
        #region Omonel
        public List<ProcesadorPagosBase> BL_Omonel(string Method)
        {
            try
            {
                List<ProcesadorPagosBase> LstppsBase = new List<ProcesadorPagosBase>();
                DataSet ds = Reportes.DA_ReporteBaseOmonel();

                string FechaOrden = string.Empty;
                string HoraOrden = string.Empty;

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

                        #region Datos Orden
                        ppsBase.PaymentToken = row["PaymentToken"].ToString();
                        ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();
                        #endregion

                        #region Tokenizacion
                        ppsBase.Bank = row["Bank"].ToString();                                          //Banco
                        ppsBase.BinCode = row["BinCode"].ToString();
                        ppsBase.MaskCard = row["MaskCard"].ToString().Substring(15);                    // Sufijo hay q recortar
                        ppsBase.TypeOfCard = row["TypeOfCard"].ToString();                              //Tipo Tarjeta
                        ppsBase.PaymentMethod = row["PaymentMethod"].ToString();                        //Marca
                        ppsBase.CustomerFirstName = row["CustomerFirstName"].ToString();
                        ppsBase.CustomerLastName = row["CustomerLastName"].ToString();
                        #endregion

                        #region Lealtad
                        ppsBase.CustomerLoyaltyRedeemPoints = row["CustomerLoyaltyRedeemPoints"].ToString(); //punto aplicados
                        ppsBase.CustomerLoyaltyRedeemMoney = row["CustomerLoyaltyRedeemMoney"].ToString();  //efectivo disponble /dinero en efectivo                              
                        ppsBase.CustomerLoyaltyCardId = row["CustomerLoyaltyCardId"].ToString();
                        #endregion

                        #region Reverso
                        //if (Method == "Reverso")
                        //{
                        //    DataSet dsReverso = Reportes.DA_ReversoBalance(row["OrderReferenceNumber"].ToString(), "Omonel");

                        //    foreach (DataTable dtRev in dsReverso.Tables)
                        //    {
                        //        foreach (DataRow rowRev in dtRev.Rows)
                        //        {
                        //            ReverseModel reverso = new ReverseModel();

                        //            ppsBase.FechaReversoAutorizacion = reverso.FechaReverso;
                        //            ppsBase.HoraReversoAutorizacion = reverso.HoraReverso;
                        //            ppsBase.MontoReverso = reverso.MontoRverso;
                        //            ppsBase.IDTransaccionReverso = reverso.IDReverso;
                        //        }
                        //    }
                        //}
                        #endregion

                      
                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;
                        #endregion

                        #region ApprovalCode
                        var ApprovalCode = BL_GetApprovalCode(row["OrderReferenceNumber"].ToString());
                        ppsBase.TransactionAuthorizationId = ApprovalCode;
                        #endregion

                        #region OUE
                        var oue = BL_Ordenes_APP(row["OrderReferenceNumber"].ToString());

                        if (oue.Id_Num_Apl == "22")
                        {
                            if (row["OrderSaleChannel"].ToString() == "1")
                                ppsBase.OrderSaleChannel = "SFWEB";
                            else
                                ppsBase.OrderSaleChannel = row["OrderSaleChannel"].ToString();                  //Canal Compra
                        }
                        else
                        {
                            ppsBase.OrderSaleChannel = "APP";                 //Canal Compra
                            ppsBase.TipoMobile = oue.CreatedBy;
                        }

                        var Delivery = Regex.Replace(oue.DeliveryType, @"[^a-zA-z0-9 ]+", "");
                        Delivery = Delivery.Replace("Envo", "Envio");

                        ppsBase.DeliveryType = Delivery;
                        #endregion

                        #region Omonel Transaction
                        var OmonelResponse = OmonelLiquidacion(row["OrderReferenceNumber"].ToString());

                        ppsBase.TransactionAuthorizationId = OmonelResponse.Cve_Autz;

                        if (OmonelResponse.Cve_Autz == "000000")
                        {
                            ppsBase.TransactionStatus = "DECLINED";
                        }
                        else
                        {
                            ppsBase.TransactionStatus = "AUTHORIZED";
                        }
                        #endregion

                        #region Cancellation                        
                        if (Method == "Creditos")
                        {
                            #region Cancellation Dev
                            /*decimal TotalPrice = 0;
                            decimal TotalPiezas = 0;
                            string Consignacion = string.Empty;
                            string productName = string.Empty;

                            var Cancelacion = BL_CancelDevolucion(ppsBase.OrderReferenceNumber);

                            if (Cancelacion.Cancelacion.OrderId != "" || Cancelacion.Devolucion.OrderId != "")
                            {
                                var Productos = BL_ArticulosByOrder(ppsBase.OrderReferenceNumber);

                                foreach (var prod in Productos)
                                {
                                    TotalPrice = decimal.Parse(prod.Price) + TotalPrice;
                                    Consignacion = prod.ProductId + ", " + Consignacion;
                                    TotalPiezas = TotalPiezas + 1;
                                    productName = productName + ", " + prod.ProductName;
                                }

                                ppsBase.NombreCancelacion = ppsBase.ShippingFirstName + " " + ppsBase.CustomerLastName;
                                ppsBase.Motivo = Cancelacion.Cancelacion.cancellationReason.Trim();

                                if (Cancelacion.Cancelacion.fec_movto == "")
                                {
                                    ppsBase.FechaCancel = "";
                                    ppsBase.HoraCancel = "";
                                }
                                else
                                {
                                    var fechaCancel = Cancelacion.Cancelacion.fec_movto;
                                    DateTime FecCancel = DateTime.Parse(fechaCancel);
                                    ppsBase.FechaCancel = FecCancel.ToString("MMMM");
                                    ppsBase.HoraCancel = fechaCancel.Substring(10);
                                }

                                ppsBase.MontoCancel = TotalPrice.ToString();
                                ppsBase.ConsignacionIDCancelada = Consignacion;
                                ppsBase.NoPiezasConsignacionCancelacion = TotalPiezas.ToString();
                                ppsBase.FechaINgresoRMA = Cancelacion.Cancelacion.fec_movto;
                                ppsBase.MontoConsignacionIDCancelada = TotalPrice.ToString();

                                ppsBase.ConsignaciónIDDevolucin = Consignacion;
                                ppsBase.DetalleConsignacionIngresada = productName;
                                ppsBase.NoPzasConsignacionDevolucion = TotalPiezas.ToString();

                                if (Cancelacion.Devolucion.fec_movto == "")
                                {
                                    ppsBase.FechaDevolucion = "";
                                    ppsBase.HoraDevolucion = "";
                                }
                                else
                                {
                                    var fechaDevolucion = Cancelacion.Devolucion.fec_movto;
                                    DateTime FecDev = DateTime.Parse(fechaDevolucion);

                                    ppsBase.FechaDevolucion = FecDev.ToString("MMMM");
                                    ppsBase.HoraDevolucion = fechaDevolucion.Substring(10);
                                }

                                ppsBase.FechaDevolucion = Cancelacion.Devolucion.fec_movto;
                                ppsBase.MontoDevolucionConsignacion = TotalPrice.ToString();

                                ppsBase.FechaReembolso = Cancelacion.Devolucion.fec_movto;
                                ppsBase.HoraReembolso = "";
                                ppsBase.FormaPagoRembolso = "";
                                ppsBase.Bin_Reembolso = row["BinCode"].ToString();
                                ppsBase.SufijoReembolso = row["MaskCard"].ToString().Substring(15);
                                ppsBase.ReembolsoAutomatico = "True";
                                ppsBase.ReembolsoManual = "";
                             
                            }*/
                            #endregion
                        }
                        #endregion

                        #region GetTrace Omonel                       
                        var datosExtra = Bl_TracePaymentOmonel(ppsBase.OrderReferenceNumber);

                        if (datosExtra.orderReferenceNumber != null)
                        {
                            if (oue.Id_Num_Apl == "22")
                            {
                                FechaOrden = datosExtra.orderDateTime.Substring(0, 10);
                                HoraOrden = datosExtra.orderDateTime.Substring(11, 5);
                            }
                            else
                            {
                                FechaOrden = datosExtra.orderDateTime.Substring(0, 10);
                                HoraOrden = datosExtra.orderDateTime.Substring(11, 5);
                            }

                            ppsBase.OrderDate = FechaOrden;
                            ppsBase.OrderHour = HoraOrden;

                            ppsBase.TransactionReferenceID = datosExtra.TransactionReferenceID;
                            ppsBase.AffiliationType = datosExtra.AffiliationType;
                            ppsBase.IsAuthenticated = datosExtra.IsAuthenticated;
                            ppsBase.IsAuthorized = datosExtra.IsAuthorized;
                            ppsBase.Apply3DS = datosExtra.Apply3DS;
                            ppsBase.MerchandiseType = datosExtra.MerchandiseType;
                            ppsBase.clientEmail = datosExtra.customerEmail;
                            ppsBase.IDTransaccionReembolso = datosExtra.TransactionReferenceID; ;

                            if (datosExtra.paymentType == "WALLET")
                            {
                                ppsBase.paymentTypeJson = "PAYPAL";
                            }
                            else
                            {
                                bool flagOmonel = datosExtra.paymentToken.Contains("OMONEL");

                                if (datosExtra.paymentToken.Contains("OMONEL"))
                                {
                                    ppsBase.paymentTypeJson = "OMONEL";
                                    ppsBase.Adquirente = "OMONEL";
                                }

                                else
                                {
                                    ppsBase.Adquirente = "EVO PAYMENT";
                                    ppsBase.paymentTypeJson = datosExtra.paymentType;
                                }
                            }

                            if (datosExtra.shipments.Count > 0)
                            {
                                foreach (var ship in datosExtra.shipments)
                                {
                                    decimal monto = decimal.Parse(ship.shippingPaymentImport);

                                    ppsBase.orderAmount = "$ " + monto.ToString("0.00");
                                    //ppsBase.orderAmount = ship.shippingPaymentImport;
                                    ppsBase.shippingDeliveryDesc = ship.shippingDeliveryDesc;
                                    ppsBase.shippingPaymentImport = ship.shippingPaymentImport;
                                    ppsBase.ShippingFirstName = ship.shippingFirstName;
                                    ppsBase.ShippingLastName = ship.shippingLastName;
                                    ppsBase.shippingPaymentInstallments = ship.shippingPaymentInstallments;

                                    #region Datos Liquidacion
                                    if (OmonelResponse.Cve_Autz != "000000" && OmonelResponse.Cve_Autz != "")
                                    {
                                        ppsBase.FechaLiquidacion = FechaOrden;
                                        ppsBase.HoraLiquidacion = HoraOrden;
                                        ppsBase.MontoLiquidacion = datosExtra.shipments[0].shippingPaymentImport;
                                        ppsBase.LiquidacionManual = "";
                                        ppsBase.LiquidacionAutomatica = "True";
                                        ppsBase.IDTransaccionLiquidacion = datosExtra.TransactionReferenceID;
                                        ppsBase.TransactionStatus = "AUTHORIZED";
                                    }
                                    #endregion

                                    foreach (var items in ship.Items)
                                    {
                                        ppsBase.shippingItemCategory = items.shippingItemCategory;
                                        ppsBase.shippingItemId = items.shippingItemId;
                                        ppsBase.shippingItemName = items.shippingItemName;
                                        ppsBase.ShippingItemTotal = items.ShippingItemTotal;
                                    }

                                    if (ship.shippingReferenceNumber == "001-1")
                                    {
                                        ppsBase.ShippingDeliveryDesc = "SETC";
                                        ppsBase.Catalogo = "SETC";
                                        ppsBase.AffiliationType = "8655759";
                                        ppsBase.Adquirente = "GETNET";

                                        LstppsBase.Add(ppsBase);
                                    }
                                    else
                                    {
                                        ppsBase.ShippingDeliveryDesc = "MG";
                                        ppsBase.Catalogo = "MG";
                                        ppsBase.AffiliationType = "1045441";
                                        ppsBase.Adquirente = "EVO Payment";

                                        LstppsBase.Add(ppsBase);
                                    }
                                }
                            }
                        }
                        #endregion



                    }
                }

                return LstppsBase;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public JsonResponseModel Bl_TracePaymentOmonel(string OrderReferenceNumber)
        {
            try
            {
                JsonResponseModel Response = new JsonResponseModel();
                string JsonRequest = string.Empty;

                DataSet ds = Reportes.DA_TracePaymentOmonel(OrderReferenceNumber);

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        JsonRequest = row["RequestJson"].ToString();
                        Response = JsonConvert.DeserializeObject<JsonResponseModel>(JsonRequest);
                    }
                }

                return Response;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Omonel_Auth OmonelLiquidacion(string OrderReferenceNumber)
        {
            try
            {
                Omonel_Auth Response = new Omonel_Auth();

                DataSet ds = Reportes.DA_Omonel_Autorizacion(OrderReferenceNumber);

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        Response.Cve_Autz = row["Cve_Autz"].ToString();
                        Response.ShippingReferenceNumber = row["ShippingReferenceNumber"].ToString();
                    }
                }

                return Response;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Creditos
        public List<ProcesadorPagosBase> BL_Creditos()
        {
            List<ProcesadorPagosBase> LstppsBase = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstOmonel = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstAPP = new List<ProcesadorPagosBase>();

            string FechaOrden = string.Empty;
            string HoraOrden = string.Empty;

            try
            {
                DataSet ds = Reportes.DA_OrdenesCanceladas();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

                        var Datos = BL_CancelDevolucion(row["OrderReferenceNumber"].ToString());

                        LstppsBase.Add(Datos);
                    }
                }

                #region Omonel
                lstOmonel = BL_Omonel("Creditos");

                foreach (var omonel in lstOmonel)
                {
                    LstppsBase.Add(omonel);
                }
                #endregion

                return LstppsBase;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public ProcesadorPagosBase BL_CancelDevolucion(string OrderReferenceNumber)
        {
            try
            {
                DataSet ds = Reportes.DA_Datos_by_Order(OrderReferenceNumber);
                ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();
                string FechaOrden = string.Empty;
                string HoraOrden = string.Empty;

                foreach (DataTable dt in ds.Tables)
                {
                    foreach(DataRow row in dt.Rows)
                    {
                        
                        #region Mapping
                        #region Datos Orden
                        ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();          //Orden ID
                        ppsBase.PaymentTransactionID = row["PaymentTransactionID"].ToString();          //Transacción
                        ppsBase.OrderSaleChannel = row["OrderSaleChannel"].ToString();                  //Canal Compra
                        #endregion

                        #region Tokenizacion
                        ppsBase.Bank = row["Bank"].ToString();                                          //Banco
                        ppsBase.BinCode = row["BinCode"].ToString();
                        ppsBase.MaskCard = row["MaskCard"].ToString().Substring(15);                    // Sufijo hay q recortar
                        ppsBase.TypeOfCard = row["TypeOfCard"].ToString();                              //Tipo Tarjeta
                        ppsBase.PaymentMethod = row["PaymentMethod"].ToString();                        //Marca
                        ppsBase.CustomerFirstName = row["CustomerFirstName"].ToString();
                        ppsBase.CustomerLastName = row["CustomerLastName"].ToString();
                        #endregion

                        #region shipping
                        ppsBase.ShippingStoreId = row["shippingStoreId"].ToString();                    //Tipo ALmacen
                        ppsBase.ShippingReferenceNumber = row["ShippingReferenceNumber"].ToString();    //Consignación ID

                        if (row["ShippingReferenceNumber"].ToString() == "001-1")
                        {
                            ppsBase.ShippingDeliveryDesc = "SETC";
                            ppsBase.Catalogo = "SETC";
                            ppsBase.AffiliationType = "8655759";
                        }

                        else
                        {
                            ppsBase.ShippingDeliveryDesc = "MG";
                            ppsBase.Catalogo = "MG";
                            ppsBase.AffiliationType = "";
                        }
                        #endregion

                        #region Lealtad
                        ppsBase.CustomerLoyaltyRedeemPoints = row["CustomerLoyaltyRedeemPoints"].ToString(); //punto aplicados
                        ppsBase.CustomerLoyaltyRedeemMoney = row["CustomerLoyaltyRedeemMoney"].ToString();  //efectivo disponble /dinero en efectivo                              
                        ppsBase.CustomerLoyaltyCardId = row["CustomerLoyaltyCardId"].ToString();
                        #endregion                   

                        #region ApprovalCode
                        var ApprovalCode = BL_GetApprovalCode(row["OrderReferenceNumber"].ToString());
                        ppsBase.TransactionAuthorizationId = ApprovalCode;
                        #endregion

                        #region OUE
                        var oue = BL_Ordenes_APP(row["OrderReferenceNumber"].ToString());

                        if (oue.Id_Num_Apl == "22" || oue.Id_Num_Apl == "")
                        {
                            if (row["OrderSaleChannel"].ToString() == "1")
                                ppsBase.OrderSaleChannel = "SFWEB";
                            else
                                ppsBase.OrderSaleChannel = row["OrderSaleChannel"].ToString();                  //Canal Compra
                        }
                        else
                        {
                            ppsBase.OrderSaleChannel = "APP";                 //Canal Compra
                            ppsBase.TipoMobile = oue.CreatedBy;
                        }

                        var Delivery = Regex.Replace(oue.DeliveryType, @"[^a-zA-z0-9 ]+", "");
                        Delivery = Delivery.Replace("Envo", "Envio");

                        ppsBase.DeliveryType = Delivery;
                        #endregion

                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;
                        #endregion

                        #region GetTrace
                        var DatosExtra = BL_TracePayment(ppsBase.OrderReferenceNumber, ppsBase.Catalogo);

                        if (DatosExtra != null)
                        {
                            if (DatosExtra.orderReferenceNumber != null)
                            {
                                if (oue.Id_Num_Apl == "22" || oue.Id_Num_Apl == "")
                                {
                                    FechaOrden = DatosExtra.orderDateTime.Substring(0, 10);
                                    HoraOrden = DatosExtra.orderDateTime.Substring(11, 5);
                                }
                                else
                                {
                                    FechaOrden = DatosExtra.orderDateTime.Substring(0, 10);
                                    HoraOrden = DatosExtra.orderDateTime.Substring(11, 5);
                                }

                                ppsBase.OrderDate = DatosExtra.orderDateTime;
                                decimal monto = decimal.Parse(DatosExtra.orderAmount);

                                ppsBase.orderAmount = "$ " + monto.ToString("0.00");
                                ppsBase.orderAmount = DatosExtra.orderAmount;
                                ppsBase.TransactionReferenceID = DatosExtra.TransactionReferenceID;
                                ppsBase.Apply3DS = DatosExtra.Apply3DS;
                                ppsBase.MerchandiseType = DatosExtra.MerchandiseType;
                                ppsBase.TransactionStatus = DatosExtra.TransactionStatus;
                                ppsBase.clientEmail = DatosExtra.customerEmail;

                                if (DatosExtra.paymentType == "WALLET")
                                {
                                    ppsBase.paymentTypeJson = "PAYPAL";
                                }
                                else
                                {
                                    bool flagOmonel = DatosExtra.paymentToken.Contains("OMONEL");

                                    if (DatosExtra.paymentToken.Contains("OMONEL"))
                                    {
                                        ppsBase.paymentTypeJson = "OMONEL";
                                        ppsBase.Adquirente = "OMONEL";
                                    }
                                    else
                                    {
                                        ppsBase.paymentTypeJson = DatosExtra.paymentType;
                                        ppsBase.Adquirente = "EVO PAYMENT";
                                    }

                                }

                                if (DatosExtra.shipments.Count > 0)
                                {
                                    ppsBase.shippingDeliveryDesc = DatosExtra.shipments[0].shippingDeliveryDesc;
                                    ppsBase.shippingPaymentImport = DatosExtra.shipments[0].shippingPaymentImport;
                                    ppsBase.shippingPaymentInstallments = DatosExtra.shipments[0].shippingPaymentInstallments;
                                    ppsBase.ShippingItemTotal = DatosExtra.shipments[0].Items[0].ShippingItemTotal;
                                    ppsBase.ShippingFirstName = DatosExtra.shipments[0].shippingFirstName;
                                    ppsBase.ShippingLastName = DatosExtra.shipments[0].shippingLastName;
                                }
                            }
                        }

                        #endregion

                        #region Cancellation
                        decimal TotalPrice = 0;
                        decimal TotalPiezas = 0;
                        string Consignacion = string.Empty;
                        string productName = string.Empty;

                        ppsBase.NombreCancelacion = ppsBase.ShippingFirstName + " " + ppsBase.CustomerLastName;
                        ppsBase.Motivo = Cancelacion.Cancelacion.cancellationReason.Trim();  //DONDE

                        ppsBase.MontoCancel = TotalPrice.ToString();   //OK
                        ppsBase.ConsignacionIDCancelada = Consignacion;   //OK
                        ppsBase.NoPiezasConsignacionCancelacion = TotalPiezas.ToString();   //OK
                        ppsBase.FechaINgresoRMA = Cancelacion.Cancelacion.fec_movto;
                        ppsBase.MontoConsignacionIDCancelada = TotalPrice.ToString();       //OK

                        ppsBase.ConsignaciónIDDevolucin = Consignacion;
                        ppsBase.DetalleConsignacionIngresada = productName;                 
                        ppsBase.NoPzasConsignacionDevolucion = TotalPiezas.ToString();

                        ppsBase.FechaDevolucion = Cancelacion.Devolucion.fec_movto;      //FEC_MOVTO
                        ppsBase.MontoDevolucionConsignacion = TotalPrice.ToString();     //MONTO REFOUND

                        ppsBase.FechaReembolso = Cancelacion.Devolucion.fec_movto;       //FECMOVYO
                        ppsBase.HoraReembolso = "";
                        ppsBase.FormaPagoRembolso = "";
                        ppsBase.Bin_Reembolso = row["BinCode"].ToString();
                        ppsBase.SufijoReembolso = row["MaskCard"].ToString().Substring(15);
                        ppsBase.ReembolsoAutomatico = "True";
                        ppsBase.ReembolsoManual = "";
                        ppsBase.IDTransaccionReembolso = DatosExtra.TransactionReferenceID;

                        /*
                        var Cancelacion = BL_CancelDevolucion(ppsBase.OrderReferenceNumber);

                        if (Cancelacion.Cancelacion.OrderId != "" || Cancelacion.Devolucion.OrderId != "")
                        {
                            var Productos = BL_ArticulosByOrder(ppsBase.OrderReferenceNumber);

                            foreach (var prod in Productos)
                            {
                                TotalPrice = decimal.Parse(prod.Price) + TotalPrice;
                                Consignacion = prod.ProductId + ", " + Consignacion;
                                TotalPiezas = TotalPiezas + 1;
                                productName = productName + ", " + prod.ProductName;
                            }

                            ppsBase.NombreCancelacion = ppsBase.ShippingFirstName + " " + ppsBase.CustomerLastName;
                            ppsBase.Motivo = Cancelacion.Cancelacion.cancellationReason.Trim();

                            if (Cancelacion.Cancelacion.fec_movto == "")
                            {
                                ppsBase.FechaCancel = "";
                                ppsBase.HoraCancel = "";
                            }
                            else
                            {
                                var fechaCancel = Cancelacion.Cancelacion.fec_movto;
                                DateTime FecCancel = DateTime.Parse(fechaCancel);
                                ppsBase.FechaCancel = FecCancel.ToString("MMMM");
                                ppsBase.HoraCancel = fechaCancel.Substring(10);
                            }

                            ppsBase.MontoCancel = TotalPrice.ToString();
                            ppsBase.ConsignacionIDCancelada = Consignacion;
                            ppsBase.NoPiezasConsignacionCancelacion = TotalPiezas.ToString();
                            ppsBase.FechaINgresoRMA = Cancelacion.Cancelacion.fec_movto;
                            ppsBase.MontoConsignacionIDCancelada = TotalPrice.ToString();


                            ppsBase.ConsignaciónIDDevolucin = Consignacion;
                            ppsBase.DetalleConsignacionIngresada = productName;
                            ppsBase.NoPzasConsignacionDevolucion = TotalPiezas.ToString();

                            if (Cancelacion.Devolucion.fec_movto == "")
                            {
                                ppsBase.FechaDevolucion = "";
                                ppsBase.HoraDevolucion = "";
                            }
                            else
                            {
                                var fechaDevolucion = Cancelacion.Devolucion.fec_movto;
                                DateTime FecDev = DateTime.Parse(fechaDevolucion);

                                ppsBase.FechaDevolucion = FecDev.ToString("MMMM");
                                ppsBase.HoraDevolucion = fechaDevolucion.Substring(10);
                            }


                            ppsBase.FechaDevolucion = Cancelacion.Devolucion.fec_movto;
                            ppsBase.MontoDevolucionConsignacion = TotalPrice.ToString();

                            ppsBase.FechaReembolso = Cancelacion.Devolucion.fec_movto;
                            ppsBase.HoraReembolso = "";
                            ppsBase.FormaPagoRembolso = "";
                            ppsBase.Bin_Reembolso = row["BinCode"].ToString();
                            ppsBase.SufijoReembolso = row["MaskCard"].ToString().Substring(15);
                            ppsBase.ReembolsoAutomatico = "True";
                            ppsBase.ReembolsoManual = "";
                            ppsBase.IDTransaccionReembolso = DatosExtra.TransactionReferenceID;
                        }*/
                        #endregion

                       
                        #endregion

                    }
                }

                return ppsBase;             
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<DetalleProducto> BL_ArticulosByOrder(string OrderReferenceNumber)
        {
            try
            {
                List<DetalleProducto> lstDetalleProd = new List<DetalleProducto>();
                string OrderNo = string.Empty;

                DataSet ds = Reportes.DA_SliptOrder(OrderReferenceNumber);

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        OrderNo = row["OrderNo"].ToString();
                    }
                }

                DataSet dsArt = Reportes.DA_ArticulosbyOrder(OrderNo);

                foreach (DataTable dtArt in dsArt.Tables)
                {
                    if (dtArt.TableName == "Table")
                    {
                        foreach (DataRow rowArt in dtArt.Rows)
                        {
                            DetalleProducto prod = new DetalleProducto
                            {
                                CodeBarra = rowArt["CodeBarra"].ToString(),
                                ProductName = rowArt["ProductName"].ToString(),
                                Quantity = rowArt["Quantity"].ToString(),
                                Price = rowArt["Price"].ToString(),
                                ProductId = rowArt["ProductId"].ToString()
                            };

                            lstDetalleProd.Add(prod);
                        }
                    }

                }

                return lstDetalleProd;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
        #endregion

        #region Excel
        public void CreateExcel_Creditos(List<ProcesadorPagosBase> lstPpsBase)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "Creeditos";
                #endregion

                #region Mapping         
                foreach (var item in lstPpsBase)
                {
                    eftExport.AddRow();

                    eftExport["Orden ID"] = item.OrderReferenceNumber;
                    eftExport["ID Transaccion"] = item.TransactionReferenceID;
                    eftExport["Fecha Creacion"] = item.OrderDate;
                    eftExport["Catalogo"] = item.Catalogo;
                    eftExport["Forma Pago"] = item.paymentTypeJson;
                    eftExport["Monto Total Orden"] = item.orderAmount;
                    eftExport["Canal Compra"] = item.OrderSaleChannel;
                    eftExport["3D/Safe Key"] = item.Apply3DS;
                    eftExport["Estatus Pago"] = item.TransactionStatus;
                    eftExport["Banco"] = item.Bank;
                    eftExport["BinCode"] = item.BinCode;
                    eftExport["Sufijo"] = item.MaskCard;
                    eftExport["Tipo Tarjeta"] = item.TypeOfCard;
                    eftExport["Marca"] = item.PaymentMethod;
                    eftExport["No. Autorizacion"] = item.TransactionAuthorizationId;
                    eftExport["MSI"] = item.shippingPaymentInstallments;
                    eftExport["Nombre Persona Registrada"] = item.ShippingFirstName;
                    eftExport["Apellido P"] = item.ShippingLastName;
                    eftExport["Apellido M"] = "";
                    eftExport["No Tarjeta Lealtad"] = item.CustomerLoyaltyCardId;
                    eftExport["Metodo de Envio"] = item.DeliveryType;
                    eftExport["Correo"] = item.clientEmail;
                    eftExport["Nombre de quien Cancela"] = item.NombreCancelacion;
                    eftExport["Motivo Cancelacion"] = item.Motivo;
                    eftExport["Fecha Cancelacion"] = item.FechaCancel;
                    eftExport["Hora Cancelacion"] = item.HoraCancel;
                    eftExport["Monto Cancelacion"] = item.MontoCancel;
                    eftExport["Consignacion ID Cancelada"] = item.ConsignacionIDCancelada;
                    eftExport["Monto Consignacion ID Cancelada"] = item.MontoConsignacionIDCancelada;
                    eftExport["No Piezas Consignacion"] = item.NoPiezasConsignacionCancelacion;
                    eftExport["Fecha INgreso RMA"] = item.FechaINgresoRMA;
                    eftExport["Consignación ID Devolucion"] = "";
                    eftExport["Detalle de la Consignacin Ingresada"] = "";
                    eftExport["No Pzas Consignacion"] = item.NoPzasConsignacionDevolucion;
                    eftExport["Fecha Devolucion"] = item.FechaDevolucion;
                    eftExport["Hora Devolucion"] = item.HoraDevolucion;
                    eftExport["Monto Devolucion Consignacion"] = item.MontoDevolucionConsignacion;
                    eftExport["Fecha Reembolso"] = item.FechaReembolso;
                    eftExport["Hora Reembolso"] = item.HoraReembolso;
                    eftExport["Forma PagoRembolso"] = item.FormaPagoRembolso;
                    eftExport["Bin Reembolso"] = item.Bin_Reembolso;
                    eftExport["Sufijo Reembolso"] = item.SufijoReembolso;
                    eftExport["Reembolso Manual"] = item.ReembolsoManual;
                    eftExport["Reembolso Automatico"] = item.ReembolsoAutomatico;
                    eftExport["ID Transaccion Reembolso"] = item.IDTransaccionReembolso;
                    eftExport["No Tarjeta Lealtad"] = item.CustomerLoyaltyCardId;
                    eftExport["Puntos"] = "";
                    eftExport["Tipo de Reembolso Programa Lealtad"] = "";
                }
                #endregion

                #region Autosize
                #endregion

                byte[] buffer = eftExport.ExportToBytes();

                FtpUpload(nombreArchivo, buffer, ".xls", true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Private Method
        public JsonResponseModel BL_TracePayment(string OrderReferenceNumber, string Catalogo)
        {
            try
            {
                string ceros = "";
                string JsonRequest = string.Empty;
                string JsonRequest_SETC = string.Empty;
                string JsonReques_MG = string.Empty;
                JsonResponseModel Response = new JsonResponseModel();

                DataSet ds = Reportes.DA_TracePayment(OrderReferenceNumber);

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        JsonRequest = row["RequestJson"].ToString();
                        Response = JsonConvert.DeserializeObject<JsonResponseModel>(JsonRequest);

                        foreach (var sh in Response.shipments)
                        {
                            if (sh.shippingReferenceNumber == "001-1")
                            {
                                JsonRequest_SETC = JsonRequest;
                            }
                            else
                            {
                                JsonReques_MG = JsonRequest;
                            }
                        }

                        JsonRequest = string.Empty;
                    }
                }

                if (JsonRequest_SETC != "" || JsonReques_MG != "")
                {
                    if (Catalogo == "SETC")
                    {
                        Response = JsonConvert.DeserializeObject<JsonResponseModel>(JsonRequest_SETC);
                    }
                    else if (Catalogo == "MG")
                    {
                        Response = JsonConvert.DeserializeObject<JsonResponseModel>(JsonReques_MG);
                    }

                    shipments shipment = new shipments();
                    items item = new items();
                    List<items> lstItems = new List<items>();

                    if (Response != null)
                    {
                        foreach (var sh in Response.shipments)
                        {
                            shipment.shippingDeliveryDesc = sh.shippingDeliveryDesc;
                            shipment.shippingPaymentImport = sh.shippingPaymentImport;
                            shipment.shippingPaymentInstallments = sh.shippingPaymentInstallments;
                            shipment.shippingFirstName = sh.shippingFirstName;
                            shipment.shippingLastName = sh.shippingLastName;

                            foreach (var row2 in sh.Items)
                            {
                                item.shippingItemCategory = row2.shippingItemCategory;
                                if (row2.shippingItemCategory == "Costo de envio")
                                {
                                    item.shippingItemId = row2.shippingItemId;
                                    item.shippingItemName = row2.shippingItemName;
                                    item.ShippingItemTotal = row2.ShippingItemTotal;

                                    lstItems.Add(item);

                                    shipment.Items = lstItems;
                                }
                                else
                                {
                                    item.shippingItemId = "";
                                    item.shippingItemName = "";
                                    item.ShippingItemTotal = "";

                                    lstItems.Add(item);

                                    shipment.Items = lstItems;
                                }
                            }
                        }
                        Response.shipments.Clear();
                        Response.shipments.Add(shipment);
                    }

                }

                return Response;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string BL_GetApprovalCode(string OrderReferenceNumber)
        {
            string dsS = string.Empty;
            string ApprovalCode = string.Empty;
            var dsJsonResponse = GetJsonResponsebyOrder_CCIncomes(OrderReferenceNumber);

            foreach (DataTable dt in dsJsonResponse.Tables)
            {
                foreach (DataRow row in dt.Rows)
                {
                    dsS = row["ResponseJson"].ToString();
                }
            }

            if (dsS != "")
            {
                if (dsS.Contains("DECLINED"))
                    ApprovalCode = "";
                else
                {
                    var approval = JsonConvert.DeserializeObject<Soriana.PPS.Common.DTO.ClosureOrder.ApprovalCodeModel>(dsS);

                    if (approval.ResponseObject.processorInformation.approvalCode != "")
                        ApprovalCode = approval.ResponseObject.processorInformation.approvalCode;
                }
            }

            return ApprovalCode;
        }

        public OrdersAppModel BL_Ordenes_APP(string OrderReferenceNumber)
        {
            try
            {
                string OrderNo = string.Empty;
                OrdersAppModel ordersApp = new OrdersAppModel();
                DataSet ds = Reportes.DA_SliptOrder(OrderReferenceNumber);

                if (OrderReferenceNumber == "368057")
                {
                    var a = OrderReferenceNumber;
                }


                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        OrderNo = row["OrderNo"].ToString();
                    }
                }

                if (OrderNo != "")
                {
                    DataSet ds2 = Reportes.DA_OrdenesAPP(OrderNo);

                    foreach (DataTable dt in ds2.Tables)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            ordersApp.OrderNo = row["OrderNo"].ToString();
                            ordersApp.StatusUe = row["StatusUe"].ToString();
                            ordersApp.CreatedBy = row["CreatedBy"].ToString();
                            ordersApp.Id_Num_Apl = row["Id_Num_Apl"].ToString();
                            ordersApp.MethodPayment = row["MethodPayment"].ToString();
                            ordersApp.DeliveryType = row["DeliveryType"].ToString();
                        }
                    }
                }

                return ordersApp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public ShipmentDataEstatus BL_EstatusShipment(string OrderReferenceNumber)
        {
            try
            {
                ShipmentDataEstatus Estatus = new ShipmentDataEstatus();
                DataSet ds = Reportes.DA_EstatusShipment(OrderReferenceNumber);

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        Estatus.CarrierName = row["Tipo-Almacen"].ToString();
                        //Estatus.OrderId = row["OrderId"].ToString(); ;
                        //Estatus.shipmentAlias = row["shipmentAlias"].ToString(); ;
                        Estatus.status = row["estatus"].ToString();
                        Estatus.DeliveryType = row["DeliveryType"].ToString();
                    }
                }

                return Estatus;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private DataSet GetJsonResponsebyOrder_CCIncomes(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.Int);
                        param.Value = OrderReferenceNumber;

                        using (System.Data.SqlClient.SqlDataAdapter dataAdapter = new System.Data.SqlClient.SqlDataAdapter(cmd))
                            dataAdapter.Fill(ds);
                    }
                }

                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region FTP
        public static void FtpUpload(string nombrearchivo, byte[] buffer, string TipoArchivo, bool copyDirectory = false)
        {
            #region Definiciones
            DateTime fecha = DateTime.Now;
            string strDate = fecha.ToString("ddMMyyyy");
            string userName = Environment.GetEnvironmentVariable("userNameFTP");
            string password = Environment.GetEnvironmentVariable("password");
            string server = Environment.GetEnvironmentVariable("server");
            string puerto = Environment.GetEnvironmentVariable("puerto");
            string pathUpload = Environment.GetEnvironmentVariable("pathUpload");
            string fullName = pathUpload + nombrearchivo;
            string fullPathUpload = pathUpload + Path.GetFileName(fullName) + strDate + TipoArchivo;
            #endregion

            try
            {
                using (SftpClient client = new SftpClient(new PasswordConnectionInfo(server, Convert.ToInt32(puerto), userName, password)))
                {
                    client.Connect();
                    client.WriteAllBytes(fullPathUpload, buffer);
                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void FtpUpload(string nombrearchivo, bool copyDirectory = false)
        {
            #region Definiciones
            DateTime fecha = DateTime.Now;
            string strDate = fecha.ToString("ddMMyyyy");
            string userName = Environment.GetEnvironmentVariable("userNameFTP");
            string password = Environment.GetEnvironmentVariable("password");
            string server = Environment.GetEnvironmentVariable("server");
            string puerto = Environment.GetEnvironmentVariable("puerto");
            string pathUpload = Environment.GetEnvironmentVariable("pathUpload");
            string ZipCopyGenerationPath = "";
            string fullNameTemp = nombrearchivo + ".csv";
            string fullName = pathUpload + nombrearchivo;
            string fullPathUpload = pathUpload + Path.GetFileName(fullName) + strDate + ".zip";
            string zipFullName = nombrearchivo + strDate + ".zip";
            #endregion

            if (File.Exists(zipFullName))
                File.Delete(zipFullName);


            using (ZipArchive zip = ZipFile.Open($"{zipFullName}", ZipArchiveMode.Create))
            {
                zip.CreateEntryFromFile($@"{fullNameTemp}", $"{ Path.GetFileName(fullNameTemp)}");
            }

            using (SftpClient client = new SftpClient(new PasswordConnectionInfo(server, Convert.ToInt32(puerto), userName, password)))
            {
                client.Connect();
                using (Stream stream = File.OpenRead(zipFullName))
                {
                    client.UploadFile(stream, fullPathUpload);
                }
                client.Disconnect();
            }

            if (copyDirectory && !String.IsNullOrEmpty(ZipCopyGenerationPath))
            {
                var zipCopyFullName = ZipCopyGenerationPath + nombrearchivo + ".zip";
                File.Copy(zipFullName, zipCopyFullName, true);
            }
        }
        #endregion
        #endregion
    }
}
