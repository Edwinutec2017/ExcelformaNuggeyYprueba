
using ExcelNugget02;
using ExcelNugget02.Interfaces;
using PruebaExcelFormat;
using System;
using System.Collections.Generic;

namespace PrueebasNugguet2
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Prueba> lista = new List<Prueba>();
            Prueba pp;
            for (int a = 1; a <= 50; a++)
            {
                pp = new Prueba()
                {
                    PLANUM = "980150263312",
                    PERIODO = "201909",
                    NIT = "04130101470019",
                    RAZON_SOCIAL = "LANDAVERDE FLORES GREGORIO DE JESUS",
                    ID_SUCURSAL = "001",
                    EMPLEADOS_DECLARADOS = "1",
                    MONTO_TOTAL = "60",
                    ARCHIVO = "MAX010325057510200012020010001.ZIP",
                    NPE = "05450000600000041781012020001202002140",
                    FECHA_ADICION = "2020-02-07-13.50.08.037772",
                    FECHA_PRESENTACIÓN = "7/2/2020",
                    CATEGORIA = "05450000456300065347032020001202004177",
                    CC = "123"
                };
                lista.Add(pp);
            }

            List<Prueba2> lista2 = new List<Prueba2>();
            Prueba2 p2;
            for (int a = 1; a <= 50; a++)
            {
                p2 = new Prueba2()
                {
                    PLANUM = "980150263312",
                    PERIODO = "201909",
                    NIT = "04130101470019",
                    RAZON_SOCIAL = "LANDAVERDE FLORES GREGORIO DE JESUS",
                    ID_SUCURSAL = "001",
                    EMPLEADOS_DECLARADOS = "1",
                    MONTO_TOTAL = "60",
                    ARCHIVO = "MAX010325057510200012020010001.ZIP",
                    NPE = "05450000600000041781012020001202002140",
                    FECHA_ADICION = "2020-02-07-13.50.08.037772",
                    FECHA_PRESENTACIÓN = "7/2/2020",
                    CATEGORIA = "05450000456300065347032020001202004177",
                    CC = "123",
                    Name = "edwin nolasco",
                    Tele = "123456789"
                };
                lista2.Add(p2);
            }

            //deuda


            int c = 0;
            //IExcel excel = new Excel();
            //for (int s=1;s<=5;s++) {
             
            //    for (int a = 1; a <= 5; a++)
            //    {
            //        c++;
            //        Console.WriteLine(c);
            //        excel.NewContent(lista, $"hoja {c}");
            //        c++;
            //        Console.WriteLine(c);
            //        excel.NewContent(lista2, $"Hoja {c}");
            //        Console.WriteLine(c);
            //    }
            //    var resp = excel.Guardar($"Planilla ").Result;

            //    if (resp.FileName != null)
            //        Console.WriteLine(resp.FileName);
            //}

            IExcel excel = new Excel("deuda");
            for (int s = 1; s <= 5; s++)
            {
                for (int a = 1; a <= 2; a++)
                {
                    c++;
                    Console.WriteLine(c);
                    excel.NewContent(lista2, $"hoja {c}");
                    c++;
                    Console.WriteLine(c);
                    excel.NewContent(lista2, $"Hoja {c}");
                    Console.WriteLine(c);
                }
                var resp = excel.Guardar($"Deuda").Result;

                if (resp.FileName != null)
                    Console.WriteLine(resp.FileName);
            }


            Console.ReadLine();

        }


    }
}
