﻿using ExcelNugget02;
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
                    CC="123"
                };
                lista.Add(pp);
            }

        
            Console.WriteLine("Hello World!");
            IExcel excel = new Excel();
            Console.WriteLine(lista.Count);

            excel.ArchivoRuta(@"C:\Users\alex\Desktop\Prueba", $"SEPP_DN2P_{DateTime.Now.ToString("ddmmyyyy")}");
            excel.NewContent(lista,"Hoja1");
            excel.NewContent(lista, "Hoja2");
            Console.WriteLine(excel.Guardar());

            //lista.Clear();
            //excel.Delete();

        }
    }
}
