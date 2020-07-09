using ExcelNugget02;
using ExcelNugget02.Dtos;
using ExcelNugget02.Interfaces;
using System;
using System.Collections.Generic;

namespace PruebaExcelFormat
{
    class Program
    {
        /*utiliza epplus.core para usar el nugget*/
        static void Main(string[] args)
        {
            List<Prueba> lista =  new List<Prueba>();
            Prueba pp;
            for (int a = 1; a <= 50; a++)
            {
                pp = new Prueba()
                {
                    Name = $"a00000000000000: {a}",
                    Edad = "980150159200",
                    Apellido = "1231111111111111",
                    Anio = "02071106470015",
                    Fecha = "KALIL FUERTES CARLOS ANTONIO",
                    Hora = "JAVIER EDUARDO AZUCAR CARRILLO"
                };
                lista.Add(pp);
            }

        
            Console.WriteLine("Hello World!");
            IExcel excel = new Excel();
            Console.WriteLine(lista.Count);
            excel.ArchivoRuta(null, "PruebaExcel6");
            excel.NewContent(lista,"Hoja1");
            excel.NewContent(lista, "Hoja2");
            excel.NewContent(lista, "Hoja3");
            excel.NewContent(lista, "Hoja4");
            Console.WriteLine(excel.Guardar());

            Console.ReadLine();
            //lista.Clear();
            //excel.Delete();

        }
    }
}
