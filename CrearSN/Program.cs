using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using SAPbobsCOM;

namespace CrearSN
{
    class Program
    {
        public static SAPbobsCOM.Company myCompany = null;
        static void Main(string[] args)
        {
            ActualizarTipoDeCambio();
        }

        public static void ActualizarTipoDeCambio()
        {
            try
            {
                SAPbobsCOM.SBObob oSBObob;
                SAPbobsCOM.Recordset oRecordSet;
                if (ConexionSAP())
                {
                    string url = "https://api.apis.net.pe/v1/tipo-cambio-sunat";
                    string jsonResult;

                    HttpClient httpClient = new HttpClient();
                    HttpResponseMessage response = httpClient.GetAsync(url).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        jsonResult = response.Content.ReadAsStringAsync().Result;

                        // Analiza la respuesta JSON para obtener el valor de la venta
                        JObject data = JObject.Parse(jsonResult);
                        double tipoCambioVenta = (double)data["venta"];

                        Console.WriteLine("Tipo de cambio obtenido: " + tipoCambioVenta);

                        if (myCompany.Connected)
                        {
                            oSBObob = myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                            oRecordSet = myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRecordSet = oSBObob.GetLocalCurrency();
                            oRecordSet = oSBObob.GetSystemCurrency();

                            // Formatear el valor del tipo de cambio manualmente
                            string tipoCambioFormatted = tipoCambioVenta.ToString("0.0000", CultureInfo.InvariantCulture);

                            oSBObob.SetCurrencyRate("USD", DateTime.Now, tipoCambioVenta, true);

                            if (myCompany.Connected)
                            {
                                myCompany.Disconnect();
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Error al obtener el tipo de cambio de la API. Código de estado: " + response.StatusCode);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Environment.Exit(0);
            }
        }


        public static bool ConexionSAP()
        {
            bool respuesta = false;
            try
            {
                myCompany = new SAPbobsCOM.Company();
                myCompany.Server = ""; //IP o Nombre del dominio del servidor de base de datos
                myCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017; //Tipo de base de datos
                // Todas las conexiones que necesitemos               
                myCompany.CompanyDB = ""; //Nombre base de datos SAP
                myCompany.UserName = ""; //Nombre usuario SAP
                myCompany.Password = ""; //Contraseña usuario SAP
                myCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La;

                int iRet = myCompany.Connect();

                if (iRet == 0)
                {
                    Console.WriteLine("Conexión exitosa a SAP");
                    respuesta = true;
                }
                else
                {
                    Console.WriteLine(myCompany.GetLastErrorDescription().ToString());
                }

                return respuesta;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());

                return respuesta;
            }
        }
    }
}
