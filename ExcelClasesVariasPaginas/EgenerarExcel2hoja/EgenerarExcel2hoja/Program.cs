using System;
using System.Collections.Generic;

namespace EgenerarExcel2hoja
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            IList<Prueba> lista =
            new List<Prueba>();
            Prueba pp;

            for (int a = 1; a <= 100; a++)
            {
                pp = new Prueba()
                {
                    PLANUM = "980150263312",
                    PERIODO = "201909",
                    NIT = "04130101470019",
                    RAZON_SOCIAL= "LANDAVERDE FLORES GREGORIO DE JESUS",
                    ID_SUCURSAL= "001",
                    EMPLEADOS_DECLARADOS= "1",
                    MONTO_TOTAL= "60",
                    ARCHIVO= "MAX010325057510200012020010001.ZIP",
                    NPE= "05450000600000041781012020001202002140",
                    FECHA_ADICION= "2020-02-07-13.50.08.037772",
                    FECHA_PRESENTACIÓN= "7/2/2020",
                    CATEGORIA= "CJU"
                };
                lista.Add(pp);
            }

            string nombreat=$"Sepp_dnp_{DateTime.Now.ToString("ddMMyyyy")}";
            Excel excel = new Excel("convdeuda");
            excel.ArchivoRuta(null, nombreat);
            string ho = "hoja1";

            excel.NewContent(lista,ho);
            ho = "Hoja2";
            excel.NewContent(lista, ho);
            Console.WriteLine(excel.Guardar());
      
            Console.ReadLine();
            //excel.Delete();

        }
    }
}
