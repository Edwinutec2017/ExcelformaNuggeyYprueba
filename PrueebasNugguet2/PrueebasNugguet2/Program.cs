using PrueebasNugguet2.Interfaces;
using System;
using System.Collections.Generic;

namespace PrueebasNugguet2
{
    class Program
    {
        static void Main(string[] args)
        {
            IList<Prueba> lista =
            new List<Prueba>();
            Prueba pp;
            for (int a = 1; a <= 1; a++)
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

            List <Content> encabezado = new List<Content>() {
                 new Content(){
                    Conten="NOTIFICACIÓN DE PLANILLA ENVIADA POR EL SEPP QUE GENERARÁ DEUDA",
                    Celda='A',
                    PositionCelda=5
                },
                  new Content(){
                    Conten="Nit del Empleador:",
                    Celda='A',
                    PositionCelda=7
                },
                  new Content(){
                    Conten="Razón Social:",
                    Celda='A',
                    PositionCelda=8
                },
                  new Content(){
                    Conten="Ubicación físca:",
                    Celda='B',
                    PositionCelda=1
                },
                  new Content(){
                    Conten="Folio:",
                    Celda='B',
                    PositionCelda=2
                },
                    new Content(){
                    Conten="Usuario:",
                    Celda='E',
                    PositionCelda=1
                },
                    new Content(){
                    Conten="Fecha Emisión:",
                    Celda='E',
                    PositionCelda=2
                },
                    new Content(){
                    Conten="Hora Emisión:",
                    Celda='E',
                    PositionCelda=3
                },
                       new Content(){
                    Conten="Supervisor:",
                    Celda='E',
                    PositionCelda=4
                },
                          new Content(){
                    Conten="Asesor:",
                    Celda='E',
                    PositionCelda=6
                },
                             new Content(){
                    Conten="Gestor Servicio a Empresas:",
                    Celda='E',
                    PositionCelda=7
                },

            };
            List<Content> pie = new List<Content>() {
                new Content(){
                    Conten="FECHA",
                    Celda='D'
                },
                   new Content(){
                    Conten="FIRMA Y SELLO",
                    Celda='E'
                },
                    new Content(){
                    Conten="Nota: Con base a correspondencia SAPEN-ISP-014290 de " +
                    "fecha 02 de Junio 2016 de la Superintendencia " +
                    "del Sistema Financiero que dice: Toda planilla " +
                    "enviada por el empleador por medio del Sistema SEPP es " +
                    "una planilla declarada y presentada, misma que estará bajo el estado de " +
                    "Declaración y No Pago (DNP) mientras no reciba un depósito bancario.",
                    Celda='A'
               
                },
                    new Content(){
                    Conten="Señor Empleador: Después de recibida esta notificación " +
                    "contará con un plazo máximo de 10 días hábiles para brindar una respuesta, caso contrario " +
                    "se entenderá que todas las planillas declaradas quedan vigentes y pasará a generarse la deuda " +
                    "de acuerdo a lo indicado en el Art. 19-A de la Ley SAP.",
                    Celda='A'
                }
            };

            List<Content> cod = new List<Content>() {
                new Content(){
                    Conten="Código de justificación.",
                    Celda='A'
                },
                  new Content(){
                    Conten="01 - Planilla Duplicada.",
                    Celda='A'
                },
                                    new Content(){
                    Conten="02 - Planilla con Error.",
                    Celda='A'
                },
                      new Content(){
                    Conten="03 - Planilla Complementaria.",
                    Celda='A'
                },
                      new Content(){
                    Conten="04 - Planilla Pagada.",
                    Celda='A'
                },
            };

            Console.WriteLine("Hello World!");
            IExcel excel = new Excel("convdeuda");
            excel.GuardarArchivo(null, "pruebaImagen5");
         
            excel.NewContent(lista);
              Console.WriteLine(excel.Ubicacion());
            Console.ReadLine();
             //lista.Clear();
            excel.Delete();
        }
    }
}
