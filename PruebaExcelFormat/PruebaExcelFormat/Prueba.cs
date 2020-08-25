using ExcelNugget02;
using System;
using System.Collections.Generic;
using System.Text;

namespace PruebaExcelFormat
{
    public class Prueba
    {
        [DescripcionExcel(Name = "PLANUM", Ignore = false)]
        public string PLANUM { get; set; }
        [DescripcionExcel(Name = "PERIODO", Ignore = false)]
        public string PERIODO { get; set; }
        [DescripcionExcel(Name = "NIT", Ignore = false)]
        public string NIT { get; set; }
        [DescripcionExcel(Name = "RAZON_SOCIAL", Ignore = false)]
        public string RAZON_SOCIAL { get; set; }
        [DescripcionExcel(Name = "ID_SUCURSAL", Ignore = false)]
        public string ID_SUCURSAL { get; set; }
        [DescripcionExcel(Name = "EMPLEADOS_DECLARADOS", Ignore = false)]
        public string EMPLEADOS_DECLARADOS { get; set; }
        [DescripcionExcel(Name = "MONTO_TOTAL", Ignore = false)]
        public string MONTO_TOTAL { get; set; }
        [DescripcionExcel(Name = "ARCHIVO", Ignore = false)]
        public string ARCHIVO { get; set; }
        [DescripcionExcel(Name = "NPE", Ignore = false)]
        public string NPE { get; set; }
        [DescripcionExcel(Name = "FECHA_ADICION", Ignore = false)]
        public string FECHA_ADICION { get; set; }
        [DescripcionExcel(Name = "FECHA_PRESENTACION", Ignore = false)]
        public string FECHA_PRESENTACIÓN { get; set; }
        [DescripcionExcel(Name = "CATEGORIA", Ignore = false)]
        public string CATEGORIA { get; set; }
        [DescripcionExcel(Name = "CATEGORIA", Ignore = false)]
        public string CC { get; set; }
    }
    public class Prueba2
    {
        [DescripcionExcel(Name = "PLANUM", Ignore = false)]
        public string PLANUM { get; set; }
        [DescripcionExcel(Name = "PERIODO", Ignore = false)]
        public string PERIODO { get; set; }
        [DescripcionExcel(Name = "NIT", Ignore = false)]
        public string NIT { get; set; }
        [DescripcionExcel(Name = "RAZON_SOCIAL", Ignore = false)]
        public string RAZON_SOCIAL { get; set; }
        [DescripcionExcel(Name = "ID_SUCURSAL", Ignore = false)]
        public string ID_SUCURSAL { get; set; }
        [DescripcionExcel(Name = "EMPLEADOS_DECLARADOS", Ignore = false)]
        public string EMPLEADOS_DECLARADOS { get; set; }
        [DescripcionExcel(Name = "MONTO_TOTAL", Ignore = false)]
        public string MONTO_TOTAL { get; set; }
        [DescripcionExcel(Name = "ARCHIVO", Ignore = false)]
        public string ARCHIVO { get; set; }
        [DescripcionExcel(Name = "NPE", Ignore = false)]
        public string NPE { get; set; }
        [DescripcionExcel(Name = "FECHA_ADICION", Ignore = false)]
        public string FECHA_ADICION { get; set; }
        [DescripcionExcel(Name = "FECHA_PRESENTACION", Ignore = false)]
        public string FECHA_PRESENTACIÓN { get; set; }
        [DescripcionExcel(Name = "CATEGORIA", Ignore = false)]
        public string CATEGORIA { get; set; }
        [DescripcionExcel(Name = "CATEGORIA", Ignore = false)]
        public string CC { get; set; }
        [DescripcionExcel(Name = "Nombre", Ignore = false)]
        public string Name { get; set; }
        [DescripcionExcel(Name = "Telefono", Ignore = false)]
        public string Tele { get; set; }
    }
}
