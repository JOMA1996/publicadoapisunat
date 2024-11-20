using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SERV_ENVIO_CORREO
{
    class Program
    {
        static void Main(string[] args)
        {
            List<EstadoProceso> envio = (new ProcesoEnvioCorreo()).ProcesoEnvio();

            Console.WriteLine("FIN");

        }
    }
}
