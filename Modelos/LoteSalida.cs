using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lotes_Hornitos.Classes
{
    public class LoteSalida
    {

        public string ItemCode { get; set; }
        public string Almacen { get; set; }
        public string AlmacenDestino { get; set; }
        public string NumeroDocumento { get; set; }
        public string TipoDocumento { get; set; }
        public string NumeroLote { get; set; }
        public double Cantidad { get; set; }
        public string FechaProceso { get; set; }
        public int NumLinea { get; set; }
        public string NumeroDocumentoEntrada { get; set; }
        public string TipoDocumentoEntrada { get; set; }

    }
}
