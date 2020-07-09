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
                    Name = $"a00000000000000: {a}",
                    Edad = 25123232,
                    Apellido = "123",
                    Anio = 201912323,
                    Fecha = 20191223,
                    Hora = 1451
                };
                lista.Add(pp);
            }
            

            string nombreat=$"Sepp_dnp { DateTime.Now.ToString("dd-MM-yyyy hh-mm-s")}";
            Excel excel = new Excel("convdeuda");
            excel.ArchivoRuta(null, nombreat);
            string ho = "hoja1";

            excel.NewContent(lista,ho);
            ho = "Hoja2";
            excel.NewContent(lista, ho);
            ho = "Hoka3";
            excel.NewContent(lista, ho);
            ho = "Hoka4";
            excel.NewContent(lista, ho);

            Console.WriteLine(excel.Guardar());
      
            Console.ReadLine();
            //excel.Delete();

        }
    }
}
