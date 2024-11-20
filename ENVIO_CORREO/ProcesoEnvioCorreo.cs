using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using DATA.AccesoSistema;
using DATA.ModeloSistema;
using UT.DocumentoPDF;

namespace SERV_ENVIO_CORREO
{
    public class ProcesoEnvioCorreo
    {
        EstadoProceso estadoProceso;

        public List<EstadoProceso> ProcesoEnvio()
        {
            bool enviado = false;
            List<EstadoProceso> listaEstadosProcesos = new List<EstadoProceso>();

            //lista de documentos a enviar
            List<DocumentoVentaCabCE> listaENVIOS = (new dbDocumentoVentaCabCD()).F_Obtener_Documento_CorreoPendinteEnvio();

            if (listaENVIOS.Count > 0)
            {
                try
                {
                    foreach (DocumentoVentaCabCE doc in listaENVIOS)
                    {
                        estadoProceso = new EstadoProceso();
                        estadoProceso.Documento = doc;
                        //si los xml han llegado al onedrive, hace el envio //deben existir ambos XML y CDR
                        if (File.Exists(doc.archivoXML)) //& File.Exists(doc.archivoCDR))
                        {
                            estadoProceso.Documento = doc; estadoProceso.Linea = 0; estadoProceso.mensaje = "";
                            UT.DocumentoPDF.DocumentoVentaPDF documentoPDF = new UT.DocumentoPDF.DocumentoVentaPDF();
                            try
                            {
                                bool pdfGenerado = documentoPDF.GenerarReporte(doc.ID_SegDoc_FacElec, doc.archivoPDF);
                                if (pdfGenerado)
                                {
                                    doc.emisor = (new dbDocumentoVentaCabCD()).ObtenerEmisor("0008", doc.RucEmpresa);
                                    //obtiene la estructura del correo
                                    if (doc.emisor != null)
                                    {
                                        string formatoCorreo = (new dbDocumentoVentaCabCD()).F_FORMATO_CORREO(1, doc);

                                        if (formatoCorreo != "")
                                        {
                                            try
                                            {
                                                MailMessage mensaje = (new dbDocumentoVentaCabCD()).F_Mensare(doc, formatoCorreo);
                                                SmtpClient cliente = (new dbDocumentoVentaCabCD()).F_Cliente(doc);
                                                cliente.Send(mensaje);
                                                enviado = (new dbDocumentoVentaCabCD()).F_Actualiza_Seguimiento_Correo(doc.ID_SegDoc_FacElec);
                                            }
                                            catch (Exception ex)
                                            {
                                                string method = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;
                                                estadoProceso.Linea = 100;
                                                estadoProceso.Proceso = method;
                                                estadoProceso.mensaje = ex.Message;
                                                listaEstadosProcesos.Add(estadoProceso);
                                                string docString = UT.Utilidades.Serializable.sSerialize(estadoProceso);
                                                (new dbDocumentoVentaCabCD()).F_CorreoEnvioLogs(doc.ID_SegDoc_FacElec, docString);
                                            }
                                        }
                                        else
                                        {
                                            string method = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;
                                            estadoProceso.Linea = 200;
                                            estadoProceso.Proceso = method;
                                            estadoProceso.mensaje = "no tiene formato de correo";
                                            listaEstadosProcesos.Add(estadoProceso);
                                            string docString = UT.Utilidades.Serializable.sSerialize(estadoProceso);
                                            (new dbDocumentoVentaCabCD()).F_CorreoEnvioLogs(doc.ID_SegDoc_FacElec, docString);

                                        }
                                    }
                                    else
                                    {
                                        string method = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;
                                        estadoProceso.Linea = 300;
                                        estadoProceso.Proceso = method;
                                        estadoProceso.mensaje = "no tiene emisor de correo";
                                        listaEstadosProcesos.Add(estadoProceso);
                                        string docString = UT.Utilidades.Serializable.sSerialize(estadoProceso);
                                        (new dbDocumentoVentaCabCD()).F_CorreoEnvioLogs(doc.ID_SegDoc_FacElec, docString);

                                    }
                                }
                                else
                                {
                                    string method = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;
                                    estadoProceso.Linea = 400;
                                    estadoProceso.Proceso = method;
                                    estadoProceso.mensaje = "no se pudo generar el PDF";
                                    listaEstadosProcesos.Add(estadoProceso);
                                    string docString = UT.Utilidades.Serializable.sSerialize(estadoProceso);
                                    (new dbDocumentoVentaCabCD()).F_CorreoEnvioLogs(doc.ID_SegDoc_FacElec, docString);
                                }
                            }
                            catch (Exception ex)
                            {
                                string method = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;
                                estadoProceso.Linea = 500;
                                estadoProceso.Proceso = method;
                                estadoProceso.mensaje = ex.Message;
                                documentoPDF = null;
                                string docString = UT.Utilidades.Serializable.sSerialize(estadoProceso);
                                (new dbDocumentoVentaCabCD()).F_CorreoEnvioLogs(doc.ID_SegDoc_FacElec, docString);
                            }
                            finally
                            {
                                documentoPDF = null;
                            }
                        }
                        else {
                            string method = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;
                            estadoProceso.Linea = 600;
                            estadoProceso.Proceso = method;
                            estadoProceso.mensaje = "aun no se encuentran los archivos en ENVIO";
                            listaEstadosProcesos.Add(estadoProceso);
                            string docString = UT.Utilidades.Serializable.sSerialize(estadoProceso);
                            (new dbDocumentoVentaCabCD()).F_CorreoEnvioLogs(doc.ID_SegDoc_FacElec, docString);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }
                finally
                {
                    listaENVIOS = null;
                }

            }

            return listaEstadosProcesos;
        }




    }

    public class EstadoProceso
    {
        public int Linea { get; set; }
        public string Proceso { get; set; }
        public string mensaje { get; set; }
        public bool Enviado { get; set; }
        public DocumentoVentaCabCE Documento { get; set; }
    }

}
