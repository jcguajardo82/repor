using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;

using EN;
using Soriana.PPS.Reportes.PaymentStoreRpt.Models;
using Jitbit.Utils;
using Soriana.PPS.Reportes.ReportesMasivos.Services;
using Newtonsoft.Json;
using Renci.SshNet;
using SorianaCCIncomesReportFunction.Models;

namespace Soriana.PPS.Reportes.PaymentStoreRpt.Services
{
    public class PaymentStoreRptService
    {
        DA_Reportes Reportes = new DA_Reportes();

        public void GenerarReportes()
        {
            var PaymentStore = GetPaymentStore();
            CreateExcel_PaymentStore(PaymentStore);

            var AprobacionesMarcas = GetAprobacionesMarcas();
            CreateExcel_AprobacionesMarcas(AprobacionesMarcas);
        }

        #region PaymentStore
        private List<PaymentStoreModelResponse> GetPaymentStore()
        {
            DataSet ds = new DataSet();
            List<ProcesadorPagosBase> LstppsBase = new List<ProcesadorPagosBase>();
            List<PaymentStoreModel> lstPaymentStore = new List<PaymentStoreModel>();
            List<PaymentStoreModelResponse> lstPaymentStoreResponse = new List<PaymentStoreModelResponse>();

            string spName = string.Empty;
            string JsonResponse = string.Empty;
            string paymentTypeJson = string.Empty;
            string ShippingDeliveryDesc = string.Empty;
            string Adquirente = string.Empty;
            string Catalogo = string.Empty;
            string AffiliationType = string.Empty;
            string CostoEnvtio = string.Empty;
            string Banco = string.Empty;
            string BIN = string.Empty;
            string Sufijo = string.Empty;
            string TipoTarjeta = string.Empty;
            string Marca = string.Empty;
            string paymentToken = string.Empty;
          
            try
            {
                ds = Reportes.DA_PaymentStore_Orders();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        PaymentStoreModel payment = new PaymentStoreModel();

                        payment.OrderReferenceNumber = row["OrderReferenceNumber"].ToString();
                        payment.OrderAmount = row["OrderAmount"].ToString();
                        payment.LineaCaptura = row["LineaCaptura"].ToString();
                        payment.Estatus = row["Estatus"].ToString();
                        payment.CreatedDate = row["CreatedDate"].ToString();

                        lstPaymentStore.Add(payment);
                    }
                }
             
                foreach (var payment in lstPaymentStore)
                {
                    ds = new DataSet();

                    ds = Reportes.DA_PaymenStore_Pagadas(payment.LineaCaptura);
                    
                    foreach (DataTable dt in ds.Tables)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            JsonResponse = row["RequestJson"].ToString();

                            if (JsonResponse != "")
                            {
                                var RequestPaymentStore = JsonConvert.DeserializeObject<JsonRespoonseModel>(JsonResponse);

                                if (RequestPaymentStore.paymentType == "INSTORE")
                                {
                                    #region Mapeo
                                    PaymentStoreModelResponse PaymentStore = new PaymentStoreModelResponse();
                                    paymentToken = RequestPaymentStore.paymentToken;
                                    string FechaOrden = RequestPaymentStore.orderDateTime.Substring(0, 10);
                                    string HoraOrden = RequestPaymentStore.orderDateTime.Substring(11, 5);

                                    if (RequestPaymentStore.paymentType == "WALLET")
                                    {
                                        paymentTypeJson = "PAYPAL";
                                    }
                                    else
                                    {
                                        bool flagOmonel = RequestPaymentStore.paymentToken.Contains("OMONEL");

                                        if (RequestPaymentStore.paymentToken.Contains("OMONEL"))
                                            paymentTypeJson = "OMONEL";
                                        else
                                            paymentTypeJson = RequestPaymentStore.paymentType;
                                    }

                                    if (RequestPaymentStore.shipments[0].shippingReferenceNumber == "001-1")
                                    {
                                        ShippingDeliveryDesc = "SETC";
                                        Catalogo = "SETC";
                                        AffiliationType = "8655759";
                                        Adquirente = "GETNET";

                                    }

                                    else
                                    {
                                        ShippingDeliveryDesc = "MG";
                                        Catalogo = "MG";
                                        AffiliationType = "1045441";
                                        Adquirente = "EVO Payment";
                                    }


                                    DataSet ds1 = new DataSet();

                                    if (paymentToken != "")
                                    {
                                        ds1 = Reportes.DA_CardDetail_PaymentStore(paymentTypeJson, paymentToken);
                                      
                                        foreach (DataTable dt1 in ds1.Tables)
                                        {
                                            foreach (DataRow row1 in dt.Rows)
                                            {
                                                Banco = row["Bank"].ToString();
                                                BIN = row["BinCode"].ToString();
                                                Sufijo = row["MaskCard"].ToString();
                                                TipoTarjeta = row["TypeOfCard"].ToString();
                                                Marca = row["PaymentMethod"].ToString();
                                            }
                                        }
                                    }

                                    #region Mapeo

                                    decimal Monto = decimal.Parse(payment.OrderAmount);

                                    PaymentStore.OrdenID = payment.OrderReferenceNumber;
                                    PaymentStore.IDtransaccion = RequestPaymentStore.TransactionReferenceID;
                                    PaymentStore.FechaCreacion = FechaOrden;
                                    PaymentStore.HoraCreacion = HoraOrden;
                                    PaymentStore.NoAfiliacion = AffiliationType;
                                    PaymentStore.Adquirente = Adquirente;
                                    PaymentStore.Catalogo = Catalogo;
                                    PaymentStore.TipoEntrega = ShippingDeliveryDesc;
                                    PaymentStore.CanalCompra = RequestPaymentStore.orderSaleChannel;
                                    //PaymentStore.FormaPago = paymentTypeJson;
                                    PaymentStore.NoTienda = RequestPaymentStore.shipments[0].shippingStoreId;
                                    PaymentStore.NombreTienda = "";
                                    PaymentStore.noCajero = "";

                                    PaymentStore.montoPagado = "$ " + Monto.ToString("0.00");
                                    PaymentStore.precioTotalOrden = "$ " + Monto.ToString("0.00");
                                    PaymentStore.fechaPago = payment.CreatedDate.Substring(0, 10);
                                    //PaymentStore.formaPago = "";
                                    PaymentStore.Banco = Banco;
                                    PaymentStore.NoAutorizacion = RequestPaymentStore.TransactionAuthorizationId;
                                    PaymentStore.BIN = BIN;
                                    PaymentStore.Sufijo = Sufijo;
                                    PaymentStore.TipoTarjeta = TipoTarjeta;
                                    PaymentStore.Marca = Marca;

                                    foreach (var sh in RequestPaymentStore.shipments)
                                    {
                                        //shipment.shippingDeliveryDesc = row.shippingDeliveryDesc;
                                        //shipment.shippingPaymentImport = row.shippingPaymentImport;
                                        //shipment.shippingPaymentInstallments = row.shippingPaymentInstallments;
                                        //shipment.shippingFirstName = row.shippingFirstName;
                                        //shipment.shippingLastName = row.shippingLastName;
                                        foreach (var row2 in sh.Items)
                                        {
                                            //item.shippingItemCategory = row2.shippingItemCategory;
                                            if (row2.shippingItemCategory == "Costo de envio")
                                            {
                                                //item.shippingItemId = row2.shippingItemId;
                                                //item.shippingItemName = row2.shippingItemName;
                                                PaymentStore.CostoEnvio = row2.ShippingItemTotal;
                                                break;
                                            }
                                            else
                                            {
                                                PaymentStore.CostoEnvio = row2.ShippingItemTotal = "";
                                            }
                                        }

                                    }

                                    PaymentStore.formato = "";
                                    PaymentStore.ciudadEstatusPago = RequestPaymentStore.customerCity;
                                    PaymentStore.MSI = RequestPaymentStore.shipments[0].shippingPaymentInstallments;
                                    PaymentStore.PuntosAplicados = RequestPaymentStore.customerLoyaltyRedeemPoints;
                                    PaymentStore.PromocionesAplicadas = "";
                                    PaymentStore.NombrePersonaRegistrada = RequestPaymentStore.shipments[0].shippingFirstName;
                                    PaymentStore.Apellido_P = RequestPaymentStore.shipments[0].shippingLastName; ;
                                    PaymentStore.Apellido_M = "";
                                    PaymentStore.NoTarjetalealtad = RequestPaymentStore.customerLoyaltyCardId;
                                    PaymentStore.Correo = RequestPaymentStore.customerEmail;
                                    PaymentStore.EstatusOrden = payment.Estatus;
                                    PaymentStore.EstatusEnvío = "";
                                    PaymentStore.almacenSurtio = "";
                                    PaymentStore.loyalty = RequestPaymentStore.customerLoyaltyCardId;

                                    PaymentStore.CreteOrderStore = payment.CreatedDate.Substring(0, 10);
                                    PaymentStore.HoraOrderStore = payment.CreatedDate.Substring(11, 5);

                                    if (RequestPaymentStore.paymentType == "WALLET")
                                    {
                                        PaymentStore.FormaPago = "PAYPAL";
                                    }
                                    else
                                    {
                                        bool flagOmonel = RequestPaymentStore.paymentToken.Contains("OMONEL");

                                        if (RequestPaymentStore.paymentToken.Contains("OMONEL"))
                                        {
                                            PaymentStore.FormaPago = "OMONEL";
                                            PaymentStore.Adquirente = "OMONEL";
                                        }
                                        else
                                        {
                                            PaymentStore.FormaPago = RequestPaymentStore.paymentType;
                                            PaymentStore.Adquirente = "EVO PAYMENT";
                                        }

                                    }
                                    #endregion

                                    lstPaymentStoreResponse.Add(PaymentStore);
                                    #endregion
                                }
                            }
                        }
                    }
                }

                return lstPaymentStoreResponse;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Aprobaciones Marcas
        private List<AprobacionesMarcas> GetAprobacionesMarcas()
        {
            #region Definiciones
            DataSet ds = new DataSet();
            List<AprobacionesMarcas> LstaprobacionesMarcas = new List<AprobacionesMarcas>();

            string spName = string.Empty;
            string CanalCompra = string.Empty;
            string TipoMobile = string.Empty;
            string Catalogo = string.Empty;

            decimal montoWEB = 0;
            decimal montoAPP = 0;
            decimal PromedioAuthWeb = 0;
            decimal PromedioDecWeb = 0;
            decimal PromedioAuthAPP = 0;
            decimal PromedioDecAPP = 0;

            int AutorizadasWeb = 0;
            int DeclinadasWeb = 0;
            int Totalweb = 0;
            int AutorizadasAPP = 0;
            int DeclinadasAPP = 0;
            int TotalAPP = 0;
            #endregion

            try
            {
                ds = Reportes.DA_ReporteBase();

                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        #region shipping
                        var ShippingReferenceNumber = row["ShippingReferenceNumber"].ToString();    //Consignación ID

                        if (row["ShippingReferenceNumber"].ToString() == "001-1")
                        {
                            Catalogo = "SETC";
                        }

                        else
                        {
                            Catalogo = "MG";
                        }
                        #endregion

                        var DatosExtra = BL_TracePayment(row["OrderReferenceNumber"].ToString(), Catalogo);

                        #region OUE

                        var oue = BL_Ordenes_APP(row["OrderReferenceNumber"].ToString());

                        if (oue.Id_Num_Apl == "22" || oue.Id_Num_Apl == "")
                        {
                            if (row["OrderSaleChannel"].ToString() == "1")
                                CanalCompra = "SFWEB";
                            else
                                CanalCompra = row["OrderSaleChannel"].ToString();                  //Canal Compra
                        }
                        else
                        {
                            CanalCompra = "APP";                 //Canal Compra
                            TipoMobile = oue.CreatedBy;
                        }
                        #endregion

                        if (CanalCompra == "SFWEB")
                        {
                            Totalweb = Totalweb + 1;

                            if (DatosExtra.TransactionStatus == "AUTHORIZED")
                            {
                                AutorizadasWeb = AutorizadasWeb + 1;

                                montoWEB = montoWEB + decimal.Parse(DatosExtra.orderAmount);

                            }
                            else
                            {
                                DeclinadasWeb = DeclinadasWeb + 1;
                            }
                        }
                        else
                        {
                            TotalAPP = TotalAPP + 1;

                            if (DatosExtra.TransactionStatus == "AUTHORIZED")
                            {
                                AutorizadasAPP = AutorizadasAPP + 1;

                                montoAPP = montoAPP + decimal.Parse(DatosExtra.orderAmount);

                            }
                            else
                            {
                                DeclinadasAPP = DeclinadasAPP + 1;
                            }

                        }
                    }
                }

                PromedioAuthWeb = (AutorizadasWeb * 100) / Totalweb;
                PromedioDecWeb = (DeclinadasWeb * 100) / Totalweb;

                AprobacionesMarcas aprobacionesWEB = new AprobacionesMarcas
                {
                    canalCompra = "WEB",
                    marca = "",
                    totalOrdenes = Totalweb.ToString(),
                    ordenesAprobadas = AutorizadasWeb.ToString(),
                    monto = "$ " + montoWEB.ToString("0.00"),
                    ordenesRechazadas = DeclinadasWeb.ToString(),
                    porcentajeAprobacion = PromedioAuthWeb.ToString() + " %",
                    porcentajeRechazo = PromedioDecWeb.ToString() + " %"
                };

                PromedioAuthAPP = (AutorizadasAPP * 100) / TotalAPP;
                PromedioDecAPP = (DeclinadasAPP * 100) / TotalAPP;

                AprobacionesMarcas aprobacionesAPP = new AprobacionesMarcas
                {
                    canalCompra = "APP",
                    marca = "",
                    totalOrdenes = TotalAPP.ToString(),
                    ordenesAprobadas = AutorizadasAPP.ToString(),
                    monto = "$ " + montoAPP.ToString("0.00"),
                    ordenesRechazadas = DeclinadasAPP.ToString(),
                    porcentajeAprobacion = PromedioAuthAPP.ToString() + " %",
                    porcentajeRechazo = PromedioDecAPP.ToString() + " %"
                };

                LstaprobacionesMarcas.Add(aprobacionesWEB);
                LstaprobacionesMarcas.Add(aprobacionesAPP);

                return LstaprobacionesMarcas;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Create Excel
        public void CreateExcel_PaymentStore(List<PaymentStoreModelResponse> lstPpsBase)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "PagoTienda";
                #endregion

                #region Mapping         
                foreach (var item in lstPpsBase)
                {
                    eftExport.AddRow();

                    eftExport["Orden ID"] = item.OrdenID;
                    eftExport["ID transaccion"] = item.IDtransaccion;
                    eftExport["Fecha de creacion"] = item.FechaCreacion;
                    eftExport["Hora de creacion"] = item.HoraCreacion;
                    eftExport["No. Afiliacion"] = item.NoAfiliacion;
                    eftExport["Adquirente"] = item.Adquirente;
                    eftExport["Catalogo"] = item.Catalogo;
                    eftExport["Tipo entrega"] = item.TipoEntrega;
                    eftExport["Canal de compra"] = item.CanalCompra;
                    eftExport["Forma de pago"] = item.formaPago;
                    eftExport["No. Tienda"] = item.NoTienda;
                    eftExport["Nombre tienda"] = item.NombreTienda;
                    eftExport["no. Cajero"] = item.noCajero;
                    eftExport["fecha creacion de la orden"] = item.CreteOrderStore;
                    eftExport["hora de creacion de la orden"] = item.HoraOrderStore;
                    eftExport["monto pagado"] = item.montoPagado;
                    eftExport["precio total de la orden"] = item.precioTotalOrden;
                    eftExport["fecha del pago"] = item.fechaPago;
                    eftExport["forma de pago"] = item.formaPago;
                    eftExport["Banco"] = item.Banco;
                    eftExport["No. Autorizacion"] = item.NoAutorizacion;
                    eftExport["BIN"] = item.BIN;
                    eftExport["Sufijo"] = item.Sufijo;
                    eftExport["Tipo de tarjeta"] = item.TipoTarjeta;
                    eftExport["Marca"] = item.Marca;
                    eftExport["formato"] = item.formato;
                    eftExport["ciudad"] = item.ciudadEstatusPago;
                    eftExport["Estatus de pago"] = item.EstatusOrden;
                    eftExport["Costo de envío"] = item.CostoEnvio;
                    eftExport["MSI"] = item.MSI;
                    eftExport["Puntos Aplicados"] = item.PuntosAplicados;
                    eftExport["Promociones aplicadas"] = item.PromocionesAplicadas;
                    eftExport["Nombre persona registrada"] = item.NombrePersonaRegistrada;
                    eftExport["Apellido_P"] = item.Apellido_P;
                    eftExport["Apellido_M"] = item.Apellido_M;
                    eftExport["No tarjeta de lealtad"] = item.loyalty;
                    eftExport["Correo"] = item.Correo;
                    eftExport["Estatus de la orden"] = item.EstatusOrden;
                    eftExport["Estatus del envío"] = item.EstatusEnvío;
                    eftExport["almacen que surtio"] = item.almacenSurtio;

                }
                #endregion
        
                byte[] buffer = eftExport.ExportToBytes();

                FtpUpload(nombreArchivo, buffer, ".xls", true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CreateExcel_AprobacionesMarcas(List<AprobacionesMarcas> lstAprobaciones)
        {
            try
            {
                #region Definiciones
                CsvExport eftExport = new CsvExport();
                string nombreArchivo = "AprobacionesMarcas";
                #endregion

                #region Mapping         
                foreach (var item in lstAprobaciones)
                {
                    eftExport.AddRow();

                    eftExport["Canal Compra"] = item.canalCompra;
                    eftExport["Tipo Tarjeta"] = item.tipoTarjeta;
                    eftExport["Total Ordenes"] = item.totalOrdenes;
                    eftExport["Monto"] = item.monto;
                    eftExport["Ordenes Aprobadas"] = item.ordenesAprobadas;
                    eftExport["Porcentaje Aprobacion"] = item.porcentajeAprobacion;
                    eftExport["Ordenes Rechazadas"] = item.ordenesRechazadas;
                    eftExport["Porcentaje Rechazo"] = item.porcentajeRechazo;
                }
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

                    EN.shipments shipment = new EN.shipments();
                    EN.items item = new EN.items();
                    List<EN.items> lstItems = new List<EN.items>();

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
    }
}
