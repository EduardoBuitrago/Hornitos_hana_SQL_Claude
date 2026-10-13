using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Lotes_Hornitos.Classes
{
    public class DataDocument
    {
        [Serializable()]
        public class BOM
        {
            public BO BO = new BO();
        }

        [Serializable()]
        public class BO
        {
            public DRF1 INV1 = new DRF1();
        }

        [Serializable()]
        public class DRF1
        {
            [XmlElement("row")]
            public List<row> row { get; set; }
        }

        [Serializable()]
        public class row
        {
            public string ItemCode;
            public double LineTotal;
            public double Price;
            public double VatSum;
            public string TaxOnly;
            public string WtLiable;
            public override string ToString()
            {
                //return "SELECT '" + ItemCode + "' AS 'Articulo', " + Price + " AS 'TotalLinea'," + VatSum + " AS 'Impuestos', '" + TaxOnly + "' AS 'SoloImpuesto', '" + WtLiable + "' AS 'AplicaRet'";
                return "SELECT '" + ItemCode + "' AS 'Articulo', CAST(REPLACE('" + Price + "',',','.') AS NUMERIC(19,6)) AS 'TotalLinea' ,  CAST(REPLACE('" + VatSum + "', ',', '.') AS NUMERIC(19, 6))  AS 'Impuestos', '" + TaxOnly + "' AS 'SoloImpuesto', '" + WtLiable + "' AS 'AplicaRet'";

            }
        }
    }
}
