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
using Jitbit.Utils;
using SorianaCCIncomesReportFunction.Models;
using System.Text.RegularExpressions;

namespace SorianaCCIncomesReportFunction.Services
{
    public class CCIncomesReportService 
    {
        DA_Reportes Reportes = new DA_Reportes();

        public void GenerarReportes()
        {
            var Autorizaciones = BL_AutorizacionBancaria();
            CreateExel_AutorizacionBancaria(Autorizaciones);

            //CreateReportTC();
        }

        #region Datos Reportes
        #region Autorizacion Bancaria
        public List<ProcesadorPagosBase> BL_AutorizacionBancaria()
        {
            List<ProcesadorPagosBase> lstPPS = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstOmonel = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstAPP = new List<ProcesadorPagosBase>();

            try
            {
                DataSet ds = Reportes.DA_ReporteBase();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

                        #region Mapping
                        #region Datos Orden
                        ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();          //Orden ID
                        ppsBase.PaymentTransactionID = row["PaymentTransactionID"].ToString();          //Transacción       
                        if (row["OrderSaleChannel"].ToString() == "1")
                            ppsBase.OrderSaleChannel = "SFWEB";
                        else
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
                            ppsBase.Adquirente = "GETNET";
                        }

                        else
                        {
                            ppsBase.ShippingDeliveryDesc = "MG";
                            ppsBase.Catalogo = "MG";
                            ppsBase.AffiliationType = "1045441";
                            ppsBase.Adquirente = "EVO Payment";
                        }
                        #endregion

                        #region Lealtad
                        ppsBase.CustomerLoyaltyRedeemPoints = row["CustomerLoyaltyRedeemPoints"].ToString(); //punto aplicados
                        ppsBase.CustomerLoyaltyRedeemMoney = row["CustomerLoyaltyRedeemMoney"].ToString();  //efectivo disponble /dinero en efectivo                              
                        ppsBase.CustomerLoyaltyCardId = row["CustomerLoyaltyCardId"].ToString();
                        #endregion

                        #region GetTrace
                        var DatosExtra = BL_TracePayment(ppsBase.OrderReferenceNumber, ppsBase.Catalogo);

                        if (DatosExtra.orderReferenceNumber != null)
                        {
                            string FechaOrden = DatosExtra.orderDateTime.Substring(0, 10);
                            string HoraOrden = DatosExtra.orderDateTime.Substring(11, 5);

                            ppsBase.OrderDate = FechaOrden;
                            ppsBase.OrderHour = HoraOrden;

                            decimal monto = decimal.Parse(DatosExtra.orderAmount); 

                            ppsBase.orderAmount = "$ " + monto.ToString("0.00"); 
                            ppsBase.TransactionReferenceID = DatosExtra.TransactionReferenceID;
                            ppsBase.IsAuthenticated = DatosExtra.IsAuthenticated;
                            ppsBase.IsAuthorized = DatosExtra.IsAuthorized;
                            ppsBase.Apply3DS = DatosExtra.Apply3DS;
                            ppsBase.MerchandiseType = DatosExtra.MerchandiseType;
                            ppsBase.PaymentTransactionService = DatosExtra.TransactionStatus;
                            ppsBase.TransactionAuthorizationId = "ID" + DatosExtra.TransactionAuthorizationId + " ";

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
                                ppsBase.shippingItemCategory = DatosExtra.shipments[0].Items[0].shippingItemCategory;
                                ppsBase.shippingItemId = DatosExtra.shipments[0].Items[0].shippingItemId;
                                ppsBase.shippingItemName = DatosExtra.shipments[0].Items[0].shippingItemName;
                                ppsBase.ShippingItemTotal = DatosExtra.shipments[0].Items[0].ShippingItemTotal;
                            }
                        }
                        #endregion

                        #region ApprovalCode
                        var ApprovalCode = BL_GetApprovalCode(row["OrderReferenceNumber"].ToString());

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
                            ppsBase.OrderSaleChannel = "APP";                                                //Canal Compra
                            ppsBase.TipoMobile = oue.CreatedBy;
                        }                     
                        #endregion

                        #region Emisor
                        var EmisorResponse = BL_DatosEmisor(ppsBase.OrderReferenceNumber);

                        ppsBase.DecisionEmisor = EmisorResponse.DecisionEmisor;
                        ppsBase.CveReespuestaEmisor = ApprovalCode; // EmisorResponse.CveReespuestaEmisor;
                        ppsBase.DescReespuestaEmisor = EmisorResponse.DescReespuestaEmisor;

                        if (EmisorResponse.DecisionEmisor == "AUTHORIZED")
                            ppsBase.DescReespuestaEmisor = "AUTHORIZED";
                        #endregion
                        #endregion

                        lstPPS.Add(ppsBase);
                    }
                }

                #region Omonel
                lstOmonel = BL_Omonel("AutBancaria");

                foreach (var omonel in lstOmonel)
                {
                    lstPPS.Add(omonel);
                }
                #endregion

                //lstAPP = BL_APP("AutBancaria")
                //{

                //}
             
                return lstPPS;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

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

        public ResponseEmisor BL_DatosEmisor(string OrderReferenceNumber)
        {
            try
            {
                ResponseEmisor responseEmisor = new ResponseEmisor();
                EN.ApprovalCodeModel aproval = new EN.ApprovalCodeModel();
                DataSet ds = Reportes.DA_DatosEmisor(OrderReferenceNumber);

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        responseEmisor.DecisionEmisor = row["StatusEmisor"].ToString();
                        responseEmisor.DescReespuestaEmisor = row["ReasonCodes"].ToString();                    
                    }
                }

                return responseEmisor;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

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

                        #region Emisor
                        if (Method == "AutBancaria")
                        {
                            var EmisorResponse = BL_DatosEmisor(ppsBase.OrderReferenceNumber);

                            ppsBase.DecisionEmisor = EmisorResponse.DecisionEmisor;
                            ppsBase.CveReespuestaEmisor = EmisorResponse.CveReespuestaEmisor;
                            ppsBase.DescReespuestaEmisor = EmisorResponse.DescReespuestaEmisor;

                            if (EmisorResponse.DecisionEmisor == "AUTHORIZED")
                                ppsBase.DescReespuestaEmisor = "AUTHORIZED";
                        }
                        #endregion

                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;
                        #endregion

                        #region ApprovalCode
                        var ApprovalCode = BL_GetApprovalCode(row["OrderReferenceNumber"].ToString());                     
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
                            ppsBase.IDTransaccionReembolso = datosExtra.TransactionReferenceID;
                            ppsBase.TransactionAuthorizationId = datosExtra.TransactionAuthorizationId;

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
                                    if(OmonelResponse.Cve_Autz != "000000"  &&  OmonelResponse.Cve_Autz != "")
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

                foreach(DataTable dt in ds.Tables)
                {
                    foreach(DataRow row in dt.Rows)
                    {
                        Response.Cve_Autz = row["Cve_Autz"].ToString();
                        Response.ShippingReferenceNumber = row["ShippingReferenceNumber"].ToString();
                    }
                }

                return Response;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Datos APP
        /*public List<ProcesadorPagosBase> BL_APP(string Method)
        {
            try
            {
                List<ProcesadorPagosBase> LstppsBase = new List<ProcesadorPagosBase>();
                DataSet ds = new DataSet();

                ds = Reportes.DA_OrdenesAPP();

                foreach(DataTable dt in ds.Tables)
                {
                    foreach(DataRow row in dt.Rows)
                    {
                        var OrderReferenceNumber = row["OrderNo"].ToString();

                        var app = BL_APP_Transaction(OrderReferenceNumber, Method);

                        if(app.OrderReferenceNumber != "")
                            LstppsBase.Add(app);
                    }
                }

                return LstppsBase;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public ProcesadorPagosBase BL_APP_Transaction(string OrderReferenceNumber, string Method)
        {
            try
            {
                ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();
                int strOrderLenght = OrderReferenceNumber.Length;
                OrderReferenceNumber = OrderReferenceNumber.Substring(0, strOrderLenght - 2);
                string FechaOrden = string.Empty;
                string HoraOrden = string.Empty;

                DataSet ds = Reportes.DA_APP_by_Order(OrderReferenceNumber);

                foreach(DataTable dt in ds.Tables)
                {
                    foreach(DataRow row in dt.Rows)
                    {   
                        if(row["OrderReferenceNumber"].ToString() != "")
                        {
                            #region Mapping
                            #region Order
                            ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();          //Orden ID
                            ppsBase.PaymentTransactionID = row["PaymentTransactionID"].ToString();          //Transacción
                            ppsBase.OrderAmount = row["OrderAmount"].ToString();                            //Monto Total Orden
                            #endregion

                            #region Tokenizacion
                            ppsBase.Bank = row["Bank"].ToString();                                          //Banco
                            ppsBase.BinCode = row["BinCode"].ToString();
                            ppsBase.MaskCard = row["MaskCard"].ToString().Substring(15);                    // Sufijo hay q recortar
                            ppsBase.TypeOfCard = row["TypeOfCard"].ToString();                              //Tipo Tarjeta
                            ppsBase.PaymentMethod = row["PaymentMethod"].ToString();                        //Marca
                            #endregion

                            #region shipping
                            ppsBase.ShippingStoreId = row["shippingStoreId"].ToString();                    //Tipo ALmacen
                            ppsBase.ShippingReferenceNumber = row["ShippingReferenceNumber"].ToString();    //Consignación ID

                            if (row["ShippingReferenceNumber"].ToString() == "001-1")
                            {
                                ppsBase.ShippingDeliveryDesc = "SETC";
                                ppsBase.Catalogo = "SETC";
                                ppsBase.AffiliationType = "8655759";
                                ppsBase.Adquirente = "GETNET";
                            }

                            else
                            {
                                ppsBase.ShippingDeliveryDesc = "MG";
                                ppsBase.Catalogo = "MG";
                                ppsBase.AffiliationType = "1045441";
                                ppsBase.Adquirente = "EVO Payment";
                            }
                            #endregion

                            #region Lealtad
                            ppsBase.CustomerLoyaltyRedeemPoints = row["CustomerLoyaltyRedeemPoints"].ToString(); //punto aplicados
                            ppsBase.CustomerLoyaltyRedeemMoney = row["CustomerLoyaltyRedeemMoney"].ToString();  //efectivo disponble /dinero en efectivo                              
                            ppsBase.CustomerLoyaltyCardId = row["CustomerLoyaltyCardId"].ToString();
                            #endregion

                            #region Reverso
                            if (Method == "Reverso")
                            {
                                DataSet dsReverso = Reportes.DA_ReversoBalance(row["OrderReferenceNumber"].ToString());

                                foreach (DataTable dtRev in dsReverso.Tables)
                                {
                                    foreach (DataRow rowRev in dtRev.Rows)
                                    {
                                        ReverseModel reverso = new ReverseModel();

                                        ppsBase.FechaReversoAutorizacion = reverso.FechaReverso;
                                        ppsBase.HoraReversoAutorizacion = reverso.HoraReverso;
                                        ppsBase.MontoReverso = reverso.MontoRverso;
                                        ppsBase.IDTransaccionReverso = reverso.IDReverso;
                                    }
                                }
                            }
                            #endregion

                            #region ApprovalCode
                            var ApprovalCode = BL_GetApprovalCode(row["OrderReferenceNumber"].ToString());
                            ppsBase.TransactionAuthorizationId = ApprovalCode;
                            #endregion

                            #region Emisor
                            if (Method == "AutBancaria")
                            {
                                var EmisorResponse = BL_DatosEmisor(ppsBase.OrderReferenceNumber);

                                ppsBase.DecisionEmisor = EmisorResponse.DecisionEmisor;
                                ppsBase.CveReespuestaEmisor = ApprovalCode;
                                ppsBase.DescReespuestaEmisor = EmisorResponse.DescReespuestaEmisor;

                                if (EmisorResponse.DecisionEmisor == "AUTHORIZED")
                                    ppsBase.DescReespuestaEmisor = "AUTHORIZED";
                                else
                                {
                                    var emisor = EmisorResponse.DecisionEmisor;
                                }
                                    
                            }
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

                            #region GetTrace
                            var DatosExtra = BL_TracePayment(ppsBase.OrderReferenceNumber, ppsBase.Catalogo);

                            if (DatosExtra != null)
                            {
                                if (DatosExtra.orderReferenceNumber != null)
                                {
                                    if (oue.Id_Num_Apl == "22")
                                    {
                                        FechaOrden = DatosExtra.orderDateTime.Substring(0, 10);
                                        HoraOrden = DatosExtra.orderDateTime.Substring(11, 5);
                                    }
                                    else
                                    {
                                        var fec = DatosExtra.orderDateTime.Replace("-", "");
                                        FechaOrden = fec.Substring(0, 8);
                                        HoraOrden = fec.Substring(9, 5);
                                    }

                                    ppsBase.OrderDate = FechaOrden;
                                    ppsBase.OrderHour = HoraOrden;
                                    ppsBase.orderAmount = DatosExtra.orderAmount;
                                    ppsBase.TransactionReferenceID = DatosExtra.TransactionReferenceID;
                                    ppsBase.IsAuthenticated = DatosExtra.IsAuthenticated;
                                    ppsBase.IsAuthorized = DatosExtra.IsAuthorized;
                                    ppsBase.Apply3DS = DatosExtra.Apply3DS;
                                    ppsBase.MerchandiseType = DatosExtra.MerchandiseType;
                                    ppsBase.clientEmail = DatosExtra.customerEmail;
                                    ppsBase.TransactionStatus = DatosExtra.TransactionStatus;
                                    ppsBase.PaymentTransactionService = DatosExtra.TransactionStatus;

                                    #region Datos Liquidacion
                                    if (DatosExtra.TransactionStatus == "AUTHORIZED")
                                    {
                                        ppsBase.FechaLiquidacion = FechaOrden;
                                        ppsBase.HoraLiquidacion = HoraOrden;
                                        ppsBase.MontoLiquidacion = DatosExtra.shipments[0].shippingPaymentImport;
                                        ppsBase.LiquidacionManual = "";
                                        ppsBase.LiquidacionAutomatica = "True";
                                        ppsBase.IDTransaccionLiquidacion = DatosExtra.TransactionReferenceID;
                                    }
                                    #endregion

                                    if (DatosExtra.paymentType == "WALLET")
                                    {
                                        ppsBase.paymentTypeJson = "PAYPAL";
                                    }
                                    else
                                    {
                                        bool flagOmonel = DatosExtra.paymentToken.Contains("OMONEL");

                                        if (DatosExtra.paymentToken.Contains("OMONEL"))
                                            ppsBase.paymentTypeJson = "OMONEL";
                                        else
                                            ppsBase.paymentTypeJson = DatosExtra.paymentType;
                                    }

                                    if (DatosExtra.shipments.Count > 0)
                                    {
                                        ppsBase.shippingDeliveryDesc = DatosExtra.shipments[0].shippingDeliveryDesc;
                                        ppsBase.shippingPaymentImport = DatosExtra.shipments[0].shippingPaymentImport;
                                        ppsBase.shippingPaymentInstallments = DatosExtra.shipments[0].shippingPaymentInstallments;
                                        ppsBase.ShippingFirstName = DatosExtra.shipments[0].shippingFirstName;
                                        ppsBase.ShippingItemTotal = DatosExtra.shipments[0].Items[0].ShippingItemTotal;
                                    }
                                }
                            }
                            #endregion

                            #region Cancellation                        
                            if (Method == "Creditos")
                            {
                                #region Cancellation Dev
                                decimal TotalPrice = 0;
                                decimal TotalPiezas = 0;
                                string Consignacion = string.Empty;
                                string productName = string.Empty;

                                ppsBase.OrderDate = DatosExtra.orderDateTime;

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
                                    ppsBase.IDTransaccionReembolso = DatosExtra.TransactionReferenceID; ;
                                }
                                #endregion
                            }
                            #endregion

                            #region Estatus Shipment
                            var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                            ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                            ppsBase.EstatusEnvio = estatusShipment.status;
                            #endregion
                            #endregion
                        }
                    }
                }

                return ppsBase;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }*/
        #endregion

        #region Liquidacion
        public List<ProcesadorPagosBase> BL_Liquidaciones()
        {
            List<ProcesadorPagosBase> LstppsBase = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstOmonel = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstAPP = new List<ProcesadorPagosBase>();
            string FechaOrden = string.Empty;
            string HoraOrden = string.Empty;

            try
            {
                DataSet ds = Reportes.DA_ReporteBase();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

                        #region Mapping
                        #region Datos Orden
                        ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();          //Orden ID
                        ppsBase.PaymentTransactionID = row["PaymentTransactionID"].ToString();          //Transacción                      
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
                            ppsBase.Adquirente = "GETNET";
                        }
                        else
                        {
                            ppsBase.ShippingDeliveryDesc = "MG";
                            ppsBase.Catalogo = "MG";
                            ppsBase.AffiliationType = "1045441";
                            ppsBase.Adquirente = "EVO Payment";
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
                            ppsBase.OrderSaleChannel = "APP";                                               //Canal Compra
                            ppsBase.TipoMobile = oue.CreatedBy;
                        }

                        var Delivery = Regex.Replace(oue.DeliveryType, @"[^a-zA-z0-9 ]+", "");
                        Delivery = Delivery.Replace("Envo", "Envio");

                        ppsBase.DeliveryType = Delivery;
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

                                ppsBase.OrderDate = FechaOrden;
                                ppsBase.OrderHour = HoraOrden;
                                decimal monto = decimal.Parse(DatosExtra.orderAmount);

                                ppsBase.orderAmount = "$ " + monto.ToString("0.00");
                                ppsBase.MerchandiseType = DatosExtra.MerchandiseType;
                                ppsBase.clientEmail = DatosExtra.customerEmail;
                                ppsBase.TransactionStatus = DatosExtra.TransactionStatus;

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
                                    ppsBase.shippingItemCategory = DatosExtra.shipments[0].Items[0].shippingItemCategory;
                                    ppsBase.shippingItemId = DatosExtra.shipments[0].Items[0].shippingItemId;
                                    ppsBase.shippingItemName = DatosExtra.shipments[0].Items[0].shippingItemName;
                                    ppsBase.ShippingItemTotal = DatosExtra.shipments[0].Items[0].ShippingItemTotal;

                                    ppsBase.ShippingFirstName = DatosExtra.shipments[0].shippingFirstName;
                                    ppsBase.ShippingLastName = DatosExtra.shipments[0].shippingLastName;

                                    #region Datos Liquidacion
                                    if(DatosExtra.TransactionStatus == "AUTHORIZED")
                                    {
                                        decimal MontoLiquidacion = decimal.Parse(DatosExtra.shipments[0].shippingPaymentImport);
                                        ppsBase.FechaLiquidacion = FechaOrden;
                                        ppsBase.HoraLiquidacion = HoraOrden;
                                        ppsBase.MontoLiquidacion = "$ " + MontoLiquidacion.ToString("0.00");
                                        ppsBase.LiquidacionManual = "";
                                        ppsBase.LiquidacionAutomatica = "True";
                                        ppsBase.IDTransaccionLiquidacion = DatosExtra.TransactionReferenceID;
                                    }                                  
                                    #endregion
                                }
                            }
                        }

                        #endregion
              
                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;
                        #endregion
                        #endregion

                        if (DatosExtra != null)
                            LstppsBase.Add(ppsBase);
                    }
                }

                #region Omonel
                lstOmonel = BL_Omonel("Liquidacion");

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
        #endregion

        #region Reverso
        public List<ProcesadorPagosBase> BL_Reversos()
        {
            List<ProcesadorPagosBase> LstppsBase = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstOmonel = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstAPP = new List<ProcesadorPagosBase>();

            string FechaOrden = string.Empty;
            string HoraOrden = string.Empty;

            try
            {
                DataSet ds = Reportes.DA_ReporteBase();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

                        #region Mapping
                        #region Datos Orden
                        ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();          //Orden ID
                        ppsBase.PaymentTransactionID = row["PaymentTransactionID"].ToString();          //Transacción
                        //ppsBase.OrderSaleChannel = row["OrderSaleChannel"].ToString();                  //Canal Compra
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
                            ppsBase.Adquirente = "GETNET";
                        }

                        else
                        {
                            ppsBase.ShippingDeliveryDesc = "MG";
                            ppsBase.Catalogo = "MG";
                            ppsBase.AffiliationType = "1045441";
                            ppsBase.Adquirente = "EVO Payment";
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

                                ppsBase.OrderDate = FechaOrden;
                                ppsBase.OrderHour = HoraOrden;
                                decimal monto = decimal.Parse(DatosExtra.orderAmount);

                                ppsBase.orderAmount = "$ " + monto.ToString("0.00");
                                ppsBase.MerchandiseType = DatosExtra.MerchandiseType;
                                ppsBase.clientEmail = DatosExtra.customerEmail;
                                ppsBase.TransactionStatus = DatosExtra.TransactionStatus;

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
                                    ppsBase.shippingItemCategory = DatosExtra.shipments[0].Items[0].shippingItemCategory;
                                    ppsBase.shippingItemId = DatosExtra.shipments[0].Items[0].shippingItemId;
                                    ppsBase.shippingItemName = DatosExtra.shipments[0].Items[0].shippingItemName;
                                    ppsBase.ShippingItemTotal = DatosExtra.shipments[0].Items[0].ShippingItemTotal;

                                    ppsBase.ShippingFirstName = DatosExtra.shipments[0].shippingFirstName;
                                    ppsBase.ShippingLastName = DatosExtra.shipments[0].shippingLastName;

                                    #region Datos Reverson
                                    ppsBase.FechaReversoAutorizacion = "";
                                    ppsBase.HoraReversoAutorizacion = "";
                                    ppsBase.MontoReverso = "";
                                    ppsBase.IDTransaccionReverso = "";
                                    #endregion
                                }
                            }
                        }

                        #endregion

                        #region Reverso
                        //DataSet dsReverso = Reportes.DA_ReversoBalance(row["OrderReferenceNumber"].ToString(), "Refound");

                        //foreach (DataTable dtRev in dsReverso.Tables)
                        //{
                        //    foreach (DataRow rowRev in dtRev.Rows)
                        //    {
                        //        ReverseModel reverso = new ReverseModel();

                        //        string FechaReverso = rowRev["fec_movto"].ToString();
                        //        string HoraReverson = rowRev["fec_movto"].ToString();
                        //        decimal MontoReverso = decimal.Parse(rowRev["TotalReverse"].ToString());

                        //        ppsBase.FechaReversoAutorizacion = FechaReverso.Substring(0,10);
                        //        ppsBase.HoraReversoAutorizacion = HoraReverson.Substring(11, 5);
                        //        ppsBase.MontoReverso = "$ " + MontoReverso.ToString("0.00"); 
                        //        ppsBase.IDTransaccionReverso = rowRev["idPayment"].ToString();
                        //    }
                        //}
                        #endregion
                   
                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;
                        #endregion

                        #endregion

                        if (DatosExtra != null)
                            LstppsBase.Add(ppsBase);
                    }
                }

                #region Omonel
                lstOmonel = BL_Omonel("Reverso");

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
                DataSet ds = Reportes.DA_ReporteBase();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

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
                                    //ppsBase.shippingItemCategory = DatosExtra.shipments[0].Items[0].shippingItemCategory;
                                    //ppsBase.shippingItemId = DatosExtra.shipments[0].Items[0].shippingItemId;
                                    //ppsBase.shippingItemName = DatosExtra.shipments[0].Items[0].shippingItemName;
                                    ppsBase.ShippingItemTotal = DatosExtra.shipments[0].Items[0].ShippingItemTotal;
                                    ppsBase.ShippingFirstName = DatosExtra.shipments[0].shippingFirstName;
                                    ppsBase.ShippingLastName = DatosExtra.shipments[0].shippingLastName;
                                }
                            }
                        }

                        #endregion

                        #region Cancellation
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
                            ppsBase.IDTransaccionReembolso = DatosExtra.TransactionReferenceID;
                        }
                        #endregion

                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;*/
                        #endregion
                        #endregion

                        LstppsBase.Add(ppsBase);
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

        public CancelDevolucionModel BL_CancelDevolucion(string OrderReferenceNumber)
        {
            try
            {
                CancelDevolucionModel Response = new CancelDevolucionModel();
                CancellationModel Cancelacion = new CancellationModel();
                DevolucionModel Devolucion = new DevolucionModel();

                DataSet ds = Reportes.DA_CancelDevolucion(OrderReferenceNumber, "cancelar");

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        Cancelacion.OrderId = row["OrderId"].ToString();
                        Cancelacion.clientEmail = row["clientEmail"].ToString();
                        Cancelacion.accion = row["accion"].ToString();
                        Cancelacion.fec_movto = row["fec_movto"].ToString();
                        Cancelacion.estatusRma = row["estatusRma"].ToString();
                        Cancelacion.ProcesoAut = row["ProcesoAut"].ToString();
                        Cancelacion.idProceso = row["idProceso"].ToString();
                        Cancelacion.cancellationReason = row["cancellationReason"].ToString();
                        break;
                    }
                }

                DataSet dsDev = Reportes.DA_CancelDevolucion(OrderReferenceNumber, "retorno");

                foreach (DataTable dt in dsDev.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        Devolucion.OrderId = row["OrderId"].ToString();
                        Devolucion.clientEmail = row["clientEmail"].ToString();
                        Devolucion.accion = row["accion"].ToString();
                        Devolucion.fec_movto = row["fec_movto"].ToString();
                        Devolucion.estatusRma = row["estatusRma"].ToString();
                        Devolucion.ProcesoAut = row["ProcesoAut"].ToString();
                        Devolucion.idProceso = row["idProceso"].ToString();
                        Devolucion.cancellationReason = row["cancellationReason"].ToString();
                        break;
                    }
                }

                Response.Cancelacion = Cancelacion;
                Response.Devolucion = Devolucion;

                return Response;
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

        #region Canal Compra
        public List<ProcesadorPagosBase> BL_CanalCompra()
        {
            List<ProcesadorPagosBase> lstPPS = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstOmonel = new List<ProcesadorPagosBase>();
            List<ProcesadorPagosBase> lstAPP = new List<ProcesadorPagosBase>();

            string FechaOrden = string.Empty;
            string HoraOrden = string.Empty;

            try
            {
                DataSet ds = Reportes.DA_ReporteBase();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        ProcesadorPagosBase ppsBase = new ProcesadorPagosBase();

                        #region Mapping

                        #region Order
                        //ppsBase.OrderSaleChannel = row["OrderSaleChannel"].ToString();                  //Canal Compra
                        ppsBase.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();          //Orden ID
                        ppsBase.PaymentTransactionID = row["PaymentTransactionID"].ToString();          //Transacción
                        ppsBase.OrderAmount = row["OrderAmount"].ToString();                            //Monto Total Orden
                        #endregion

                        #region Tokenizacion
                        ppsBase.Bank = row["Bank"].ToString();                                          //Banco
                        ppsBase.BinCode = row["BinCode"].ToString();
                        ppsBase.MaskCard = row["MaskCard"].ToString().Substring(15);                    // Sufijo hay q recortar
                        ppsBase.TypeOfCard = row["TypeOfCard"].ToString();                              //Tipo Tarjeta
                        ppsBase.PaymentMethod = row["PaymentMethod"].ToString();                        //Marca
                        #endregion

                        #region shipping
                        ppsBase.ShippingStoreId = row["shippingStoreId"].ToString();                    //Tipo ALmacen
                        ppsBase.ShippingReferenceNumber = row["ShippingReferenceNumber"].ToString();    //Consignación ID

                        if (row["ShippingReferenceNumber"].ToString() == "001-1")
                        {
                            ppsBase.ShippingDeliveryDesc = "SETC";
                            ppsBase.Catalogo = "SETC";
                            ppsBase.AffiliationType = "8655759";
                            ppsBase.Adquirente = "GETNET";
                        }

                        else
                        {
                            ppsBase.ShippingDeliveryDesc = "MG";
                            ppsBase.Catalogo = "MG";
                            ppsBase.AffiliationType = "1045441";
                            ppsBase.Adquirente = "EVO Payment";
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

                                    //FechaOrden = DatosExtra.orderDateTime.Substring(0, 8);
                                    //HoraOrden = DatosExtra.orderDateTime.Substring(9, 5);
                                }

                                ppsBase.OrderDate = FechaOrden;
                                ppsBase.OrderHour = HoraOrden;

                                decimal monto = decimal.Parse(DatosExtra.orderAmount);

                                ppsBase.orderAmount = "$ " + monto.ToString("0.00");                           
                                ppsBase.TransactionReferenceID = DatosExtra.TransactionReferenceID;
                                ppsBase.IsAuthenticated = DatosExtra.IsAuthenticated;
                                ppsBase.IsAuthorized = DatosExtra.IsAuthorized;
                                ppsBase.Apply3DS = DatosExtra.Apply3DS;
                                ppsBase.MerchandiseType = DatosExtra.MerchandiseType;
                                ppsBase.clientEmail = DatosExtra.customerEmail;
                                ppsBase.TransactionStatus = DatosExtra.TransactionStatus;

                                if (DatosExtra.paymentType == "WALLET")
                                {
                                    ppsBase.paymentTypeJson = "PAYPAL";
                                }
                                else
                                {
                                    bool flagOmonel = DatosExtra.paymentToken.Contains("OMONEL");

                                    if (DatosExtra.paymentToken.Contains("OMONEL"))
                                        ppsBase.paymentTypeJson = "OMONEL";
                                    else
                                        ppsBase.paymentTypeJson = DatosExtra.paymentType;
                                }

                                if (DatosExtra.shipments.Count > 0)
                                {
                                    ppsBase.shippingDeliveryDesc = DatosExtra.shipments[0].shippingDeliveryDesc;
                                    ppsBase.shippingPaymentImport = DatosExtra.shipments[0].shippingPaymentImport;
                                    ppsBase.shippingPaymentInstallments = DatosExtra.shipments[0].shippingPaymentInstallments;
                                    ppsBase.ShippingFirstName = DatosExtra.shipments[0].shippingFirstName;
                                }
                            }
                        }
                        #endregion
                     
                        #region Estatus Shipment
                        var estatusShipment = BL_EstatusShipment(row["OrderReferenceNumber"].ToString());

                        ppsBase.TipoAlmacen = estatusShipment.CarrierName;
                        ppsBase.EstatusEnvio = estatusShipment.status;
                        #endregion
                        #endregion

                        lstPPS.Add(ppsBase);
                    }
                }

                #region Omonel
                lstOmonel = BL_Omonel("CanalCompra");

                foreach (var omonel in lstOmonel)
                {
                    lstPPS.Add(omonel);
                }
                #endregion
            
                return lstPPS;
            }
            catch (Exception ex)
            {
                throw ex;
            }
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

                if(OrderNo != "")
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
            catch(Exception ex)
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
        #endregion

        #region Ingresos TC
        public void CreateReportTC()
        {
            #region Definiciones
            string dsS = string.Empty;
            int lastColIndex = 0;
            int i = 0;
            CsvExport eftExport = new CsvExport();
            string nombreArchivo = "ReportePagoTC";
            #endregion

            var dsOrdersCCIncome = GetOrderIncomes();
            lastColIndex = dsOrdersCCIncome.Tables[0].Columns.Count - 1;

            foreach (DataTable dtOrders in dsOrdersCCIncome.Tables)
            {
                foreach (DataRow rowOrder in dtOrders.Rows)
                {
                    int strOrderLenght = rowOrder["Orden"].ToString().Length;
                    var Order = rowOrder["Orden"].ToString().Substring(0, strOrderLenght - 2); //int.Parse(rowOrder["Orden"].ToString()) / 100;

                    DataSet dsJsonResponse = GetJsonResponsebyOrder_CCIncomes(Order.ToString());

                    foreach (DataTable dt in dsJsonResponse.Tables)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            dsS = row["ResponseJson"].ToString();
                        }
                    }

                    if (dsS != null)
                    {
                        var approval = JsonConvert.DeserializeObject<Soriana.PPS.Common.DTO.ClosureOrder.ApprovalCodeModel>(dsS);

                        if (approval.ResponseObject.processorInformation.approvalCode != "")
                            dsOrdersCCIncome.Tables[0].Rows[i][lastColIndex] = approval.ResponseObject.processorInformation.approvalCode;
                    }

                    i++;
                }
            }

            #region Mapping
            foreach (DataTable dt in dsOrdersCCIncome.Tables)
            {
                foreach (DataRow row in dt.Rows)
                {
                    eftExport.AddRow();

                    eftExport["Fecha"] = row["Fecha"].ToString();
                    eftExport["Sucursal"] = row["Sucursal"].ToString();
                    eftExport["Cajero"] = row["Cajero"].ToString();
                    eftExport["Caja"] = row["Caja"].ToString();
                    eftExport["Transaccion"] = row["Transaccion"].ToString();
                    eftExport["Total"] = row["Total"].ToString();
                    eftExport["Orden"] = row["Orden"].ToString();
                    eftExport["FechaOrden"] = row["FechaOrden"].ToString();
                    eftExport["MetodoPago"] = row["MetodoPago"].ToString();
                    eftExport["Tarjeta"] = row["Tarjeta"].ToString();
                    eftExport["Autorizacion"] = row["Autorizacion"].ToString();
                }
            }
            #endregion

            byte[] buffer = eftExport.ExportToBytes();

            FtpUpload(nombreArchivo, buffer, ".csv", true);
        }
        #endregion
        #endregion

        #region Excel
        public void CreateExel_AutorizacionBancaria(List<ProcesadorPagosBase> lstAutorizacionesBancarias)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "AutorizacionesBancarias";              
                #endregion

                #region Mapping         
                foreach (var item in lstAutorizacionesBancarias)
                {
                    eftExport.AddRow();

                    eftExport["Orden ID"] = item.OrderReferenceNumber;
                    eftExport["ID Transaccion"] = item.PaymentTransactionID;
                    eftExport["Fecha Creacion"] = item.OrderDate;
                    eftExport["Hora Creacion"] = item.OrderHour;
                    eftExport["Monto Total Orden"] = item.orderAmount;
                    eftExport["Banco"] = item.Bank;
                    eftExport["BinCode"] = item.BinCode;
                    eftExport["Sufijo"] = item.MaskCard;
                    eftExport["ipo Tarjeta"] = item.TypeOfCard;
                    eftExport["Marca"] = item.PaymentMethod;
                    eftExport["Numero Autorizacion"] = item.TransactionAuthorizationId.ToString();
                    eftExport["MSI"] = item.shippingPaymentInstallments;
                    eftExport["Decision Emisor"] = item.DecisionEmisor;
                    eftExport["Codigo Respuesta emisor"] = item.CveReespuestaEmisor;
                    eftExport["Descripcion Respuesta emisor"] = item.DescReespuestaEmisor;
                    eftExport["Catalogo"] = item.Catalogo;
                    eftExport["Canal Compra"] = item.OrderSaleChannel;
                    eftExport["Forma Pago"] = item.paymentTypeJson;
                    eftExport["3D/Safe Key"] = item.Apply3DS;
                    eftExport["3D/Safe Key"] = item.Apply3DS;
                    eftExport["Estatus Orden"] = item.PaymentTransactionService;
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

        public void CreateExcel_Liquidaciones(List<ProcesadorPagosBase> lstPpsBase)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "Liquidaciones";
                #endregion

                #region Mapping         
                foreach (var item in lstPpsBase)
                {                    
                    eftExport.AddRow();     
                    
                    eftExport["Orden ID"] = item.OrderReferenceNumber;
                    eftExport["ID transaccion"] = item.PaymentTransactionID;
                    eftExport["Fecha de creaciOn"] = item.OrderDate;
                    eftExport["Hora de creación"] = item.OrderHour;
                    eftExport["No. Afiliacion"] = item.AffiliationType;
                    eftExport["Adquirente"] = item.Adquirente;
                    eftExport["Catalogo"] = item.Catalogo;
                    eftExport["Tipo entrega"] = item.DeliveryType;
                    eftExport["Canal de compra"] = item.OrderSaleChannel;
                    eftExport["Forma de pago"] = item.paymentTypeJson;
                    eftExport["Estatus de pago"] = item.TransactionStatus;
                    eftExport["Tipo de almacen"] = item.TipoAlmacen;
                    eftExport["Estatus del envio"] = item.EstatusEnvio;
                    eftExport["Costo de envio"] = item.ShippingItemTotal;
                    eftExport["Monto total de la orden"] = item.orderAmount;
                    eftExport["Banco"] = item.Bank;
                    eftExport["BinCode"] = item.BinCode;
                    eftExport["Sufijo"] = item.MaskCard;
                    eftExport["Tipo de tarjeta"] = item.PaymentMethod;
                    eftExport["Marca"] = item.TypeOfCard;
                    eftExport["No. Autorizacion"] = item.TransactionAuthorizationId;
                    eftExport["MSI"] = item.shippingPaymentInstallments;
                    eftExport["Correo"] = item.clientEmail;
                    eftExport["Nombre de persona a recibir"] = item.ShippingFirstName;
                    eftExport["Apellido_P"] = item.ShippingLastName;
                    eftExport["Apellido M"] = "";
                    eftExport["Fecha liquidacion"] = item.FechaLiquidacion;
                    eftExport["Hora de liquidacion"] = item.HoraLiquidacion;
                    eftExport["Monto de liquidacion"] = item.MontoLiquidacion;
                    eftExport["Liquidacion Manual"] = item.LiquidacionManual;
                    eftExport["Liquidacion Automatica"] = item.LiquidacionAutomatica;
                    eftExport["ID Transaccion Liquidacion"] = item.IDTransaccionLiquidacion;
                }
                #endregion

                #region Autosize
                //int lastColumNum = sheet1.GetRow(0).LastCellNum;
                //for (int inc = 0; inc <= lastColumNum; inc++)
                //{
                //    sheet1.AutoSizeColumn(inc);
                //    GC.Collect();
                //}
                #endregion

                byte[] buffer = eftExport.ExportToBytes();

                FtpUpload(nombreArchivo, buffer, ".xls", true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CreateExcel_Reverso(List<ProcesadorPagosBase> lstPpsBase)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "Reversos";
                #endregion

                #region Mapping         
                foreach (var item in lstPpsBase)
                {
                    eftExport.AddRow();

                    eftExport["Orden ID"] = item.OrderReferenceNumber;
                    eftExport["ID transaccion"] = item.PaymentTransactionID;
                    eftExport["Fecha de creación"] = item.OrderDate;
                    eftExport["Hora de creación"] = item.OrderHour;
                    eftExport["No. Afiliacion"] = item.AffiliationType;
                    eftExport["Adquirente"] = item.Adquirente;
                    eftExport["Catalogo"] = item.Catalogo;
                    eftExport["Tipo entrega"] = item.DeliveryType;
                    eftExport["Canal de compra"] = item.OrderSaleChannel;
                    eftExport["Forma de pago"] = item.paymentTypeJson;
                    eftExport["Estatus de pago"] = item.TransactionStatus;
                    eftExport["Tipo de almacen"] = item.TipoAlmacen;
                    eftExport["Estatus del envío"] = item.EstatusEnvio;
                    eftExport["Costo del envío"] = item.ShippingItemTotal;
                    eftExport["Monto total de la orden"] = item.orderAmount;
                    eftExport["Banco"] = item.Bank;
                    eftExport["BinCode"] = item.BinCode;
                    eftExport["Sufijo"] = item.MaskCard;
                    eftExport["Tipo de tarjeta"] = item.TypeOfCard;
                    eftExport["Marca"] = item.PaymentMethod;
                    eftExport["No. Autorizacion"] = item.TransactionAuthorizationId;
                    eftExport["MSI"] = item.shippingPaymentInstallments;
                    eftExport["Correo"] = item.clientEmail;
                    eftExport["Nombre de persona a recibir"] = item.ShippingFirstName;
                    eftExport["Apellido_P"] = item.ShippingLastName;
                    eftExport["Apellido_M"] = "";
                    eftExport["Fecha reversión de autorización"] = item.FechaReversoAutorizacion;
                    eftExport["Hora  de reversión de  autorización"] = item.HoraReversoAutorizacion;
                    eftExport["Monto Reverso"] = item.MontoReverso;
                    eftExport["Id Transaccion Reverso de la Autorizacion"] = item.IDTransaccionReverso;
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

        public void CreateExcel_CanalCompra(List<ProcesadorPagosBase> lstPpsBase)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "CanalCompra";
                #endregion

                #region Mapping         
                foreach (var item in lstPpsBase)
                {
                    eftExport.AddRow();

                    eftExport["Canal de compra"] = item.OrderSaleChannel;
                    eftExport["Tipo de mobile"] = item.TipoMobile;
                    eftExport["Orden ID"] = item.OrderReferenceNumber;
                    eftExport["ID transaccion"] = item.PaymentTransactionID;
                    eftExport["Fecha de creacion"] = item.OrderDate;
                    eftExport["Hora de creacion"] = item.OrderHour;
                    eftExport["No. Afiliacion"] = item.AffiliationType;
                    eftExport["Adquirente"] = item.Adquirente;
                    eftExport["Catalogo"] = item.Catalogo;
                    eftExport["Tipo entrega"] = item.DeliveryType;
                    eftExport["Tipo Almacen"] = item.TipoAlmacen;
                    eftExport["Forma de pago"] = item.paymentTypeJson;
                    eftExport["3Ds/safe key"] = item.Apply3DS;
                    eftExport["Estatus de pago de envio"] = item.EstatusEnvio;
                    eftExport["Monto total de la orden"] = item.orderAmount;
                    eftExport["Banco"] = item.Bank;
                    eftExport["BinCode"] = item.BinCode;
                    eftExport["Sufijo"] = item.MaskCard;
                    eftExport["Tipo de tarjeta"] = item.TypeOfCard;
                    eftExport["Marca"] = item.PaymentMethod;
                    eftExport["Codigo Aprovacion"] = item.TransactionAuthorizationId;
                    eftExport["MSI"] = item.shippingPaymentInstallments;
                    eftExport["Puntos Aplicados"] = item.CustomerLoyaltyRedeemPoints;
                    eftExport["Promociones aplicadas"] = "";
                    eftExport["Nombre persona registrada"] = item.ShippingFirstName;
                    eftExport["Correo"] = item.clientEmail;
                    eftExport["Estatus de la orden"] = item.TransactionStatus;
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
        private DataSet GetOrderIncomes()
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = DatabaseSchemaConstants.PROCEDURE_NAME_CC_INCOMES_REPORT_ORDERS;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

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
            catch(Exception ex)
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
