using EgenerarExcel2hoja.Dto;
using System;
using System.Collections.Generic;
using System.Text;

namespace EgenerarExcel2hoja
{
    public class Prueba
    {
        [DescripcionExcel(Name = "FECHA DE RECEPCION", Ignore = false)]
        public string Name { get; set; }
        [DescripcionExcel(Name = "CODIGO DE RECEPCION", Ignore = false)]
        public int Edad { get; set; }
        [DescripcionExcel(Name = "CENTRO DE TRABAJO", Ignore = false)]
        public string Apellido { get; set; }
        [DescripcionExcel(Name = "PERIODO DEVENGUE", Ignore = false)]
        public int Anio { get; set; }
        [DescripcionExcel(Name = "MONTO DE LA PLANILLA", Ignore = false)]
        public int Fecha { get; set; }
        [DescripcionExcel(Name = "NUM. DE AFILIADOS DECLARADOS", Ignore = false)]
        public int Hora { get; set; }

    }
}
