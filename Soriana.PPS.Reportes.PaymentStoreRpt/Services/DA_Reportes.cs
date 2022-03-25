using System;
using System.Collections.Generic;
using System.Text;
using System.Data;


namespace Soriana.PPS.Reportes.ReportesMasivos.Services
{
    public class DA_Reportes
    {
        public DataSet DA_ReporteBase()
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_PaymentTransactionRpt";

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        //System.Data.SqlClient.SqlParameter param;
                        //param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.Int);
                        //param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_TracePayment(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_Trace_payments_Prov";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
                        param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_DatosEmisor(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_Sel_DecisionEmisor";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
                        param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_ReporteBaseOmonel()
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_PaymentTransactionOmonelRpt";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        //System.Data.SqlClient.SqlParameter param;
                        //param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.Int);
                        //param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_TracePaymentOmonel(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_Trace_payments_Prov_Omonel";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
                        param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_ReversoBalance(string OrderReferenceNumber, string method)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_Sel_Refound_Rev";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
                        param.Value = int.Parse("00325015");

                        System.Data.SqlClient.SqlParameter param1;
                        param1 = cmd.Parameters.Add("@method", SqlDbType.NVarChar);
                        param1.Value = method;

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

        public DataSet DA_APP_by_Order(string OrderRefrenceNumber)
        {
            DataSet ds = new DataSet();
            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_PaymentTransactionRpt_byOrder";

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;


                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.VarChar);
                        param.Value = OrderRefrenceNumber;

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

        public DataSet DA_Omonel_Autorizacion(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();
            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_OmonelAuth";

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;


                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.VarChar);
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

        public DataSet DA_PaymentStore_Orders()
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_PaymentStoreRpt";

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
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public DataSet DA_PaymenStore_Pagadas(string LineaCaptura)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_PaymentStorePagadaRpt";

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@barcode", SqlDbType.VarChar);
                        param.Value = LineaCaptura;

                        using (System.Data.SqlClient.SqlDataAdapter dataAdapter = new System.Data.SqlClient.SqlDataAdapter(cmd))
                            dataAdapter.Fill(ds);
                    }
                }

                return ds;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public DataSet DA_CardDetail_PaymentStore(string paymentTypeJson, string paymentToken)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurioPPS");
                string spName = "up_PPS_sel_CardDetailRpt";

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@TypeCard", SqlDbType.VarChar);
                        param.Value = paymentTypeJson;

                        System.Data.SqlClient.SqlParameter param2;
                        param2 = cmd.Parameters.Add("@ClientToken", SqlDbType.VarChar);
                        param2.Value = paymentToken;

                        using (System.Data.SqlClient.SqlDataAdapter dataAdapter = new System.Data.SqlClient.SqlDataAdapter(cmd))
                            dataAdapter.Fill(ds);
                    }
                }

                return ds;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        //MERCURIODB
        public DataSet DA_CancelDevolucion(string OrderReferenceNumber, string accion)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = "up_PPS_Sel_Cancel_Devolucion";  //DatabaseSchemaConstants.PROCEDU RE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
                        param.Value = OrderReferenceNumber;

                        System.Data.SqlClient.SqlParameter param1;
                        param1 = cmd.Parameters.Add("@accion", SqlDbType.NVarChar);
                        param1.Value = accion;

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

        public DataSet DA_SliptOrder(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = "spDatosSplitOrder_rpt_sUP";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar);
                        param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_ArticulosbyOrder(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = "spDatosArticulosbyOrderId_rpt_sUP";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderId", SqlDbType.NVarChar);
                        param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_EstatusShipment(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = "up_PPS_sel_EstatusShipment";  //DatabaseSchemaConstants.PROCEDURE_NAME_GET_APPROVAL_CODE;

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
                        param.Value = int.Parse(OrderReferenceNumber);

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

        public DataSet DA_OrdenesAPP(string OrderReferenceNumber)
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = "up_PPS_Sel_OrdersApp";

                using (System.Data.SqlClient.SqlConnection cnn = new System.Data.SqlClient.SqlConnection(conn))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(spName, cnn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        System.Data.SqlClient.SqlParameter param;
                        param = cmd.Parameters.Add("@OrderReferenceNumber", SqlDbType.NVarChar);
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

        public DataSet DA_OrdenesAPP()
        {
            DataSet ds = new DataSet();

            try
            {
                string conn = Environment.GetEnvironmentVariable("AmbienteMercurio");
                string spName = "up_PPS_Sel_OrdersApp";

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
    }
}
