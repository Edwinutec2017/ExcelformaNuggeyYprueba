﻿
using log4net;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PrueebasNugguet2.Dto;
using PrueebasNugguet2.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Threading.Tasks;
namespace PrueebasNugguet2
{
   public class Excel:IExcel,IDisposable
    {
        #region CONSTRUCTOR
        public Excel()
        {
                celdaInicio = 'A';
                positionInicion = 2;
        }
        public Excel(string proceso)
        {
            if (proceso.Equals("convdeuda")) {
                _proceso = proceso;
                celdaInicio = 'A';
                positionInicion = 1;
            }

        }
        #endregion
        #region Atributos
        private readonly string _proceso = null; 
        private ExtraerConten extra = new ExtraerConten();
        private char celdaInicio, celdaFinal;
        private int positionInicion;
        private PropertyInfo[] properties = null;
        private DescripcionExcel myAttribute;
        private object[] attributes = null;
        private List<string[]> headerRow = new List<string[]>();
        private List<string[]> data = new List<string[]>();
        private string[] dataconte = null;
        private ExcelWorksheet worksheet;
        private ExcelPackage excel = new ExcelPackage();
        private bool resp;
        private string nombre_archivo = "";
        private string UbicacionDoc;
        private IList<Content> encabezados;
        private IList<Content> piePagina;
        private IList<Content> Cod;
        private string RutaImagen;
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        #endregion
        #region METODOS
        #region ENCABEZADO Y PIE 
        public void Encabezado(IList<Content> encabezadoExcel)
        {

            if (encabezadoExcel != null)
            {
                if (encabezadoExcel.Count <= 3)
                    positionInicion = positionInicion + 2;
                else if (encabezadoExcel.Count > 4)
                    positionInicion = positionInicion + 10;
            }
            this.encabezados = encabezadoExcel;
            Dispos(true);
        }
        public void PiePagina(IList<Content> PieExcel)
        {
            this.piePagina = PieExcel;
        }
        #endregion
        #region LOGO
        public void NombreLogo(string nombreImagen)
        {
            var ruta = $@"../../../Img/{nombreImagen}";
            this.RutaImagen = ruta ;
        }
        #endregion
        #region CONTENIDO
        public Task<bool> NewContent<T>(IList<T> datos)
        {
            try
            {
                if (datos.Count > 0)
                {
                    var cantidad = extra.GetHeader(datos.FirstOrDefault());
                    headerRow = extra.Data();
                    dataconte = null;
                    foreach (object obj in datos)
                    {
                        dataconte = new string[cantidad];
                        var indice = 0;
                        properties = obj.GetType().GetProperties();
                        foreach (PropertyInfo property in properties)
                        {
                            attributes = property.GetCustomAttributes(typeof(DescripcionExcel), true);
                            if (attributes.Length > 0)
                            {
                                myAttribute = (DescripcionExcel)attributes[0];
                                if (!myAttribute.Ignore)
                                {
                                    if (property.GetValue(obj) != null)
                                    {
                                        dataconte[indice] = property.GetValue(obj).ToString();
                                    }
                                    else
                                    dataconte[indice] = "";
                                }
                                else
                                    indice--;
                            }
                            indice++;
                        }
                        data.Add(dataconte);
                    }
                    Header();
                    Content();
                    Save();
                }
                Dispos(true);
                return Task.FromResult(true);
            }
            catch (IOException ex)
            {
                  
                _log.Error($"Error en el contenido {ex.StackTrace}");
                throw ex;
            }
        }
        private bool Main()
        {
            try
            {
                if (Header() && Content() && Save())
                {
                    Limpiar();
                    resp = true;
                }
                else
                {
                    resp = false;
                }
               
                return resp;
            }
            catch (IOException ex)
            {
                _log.Error($"Error al crear el archico excel {ex.StackTrace}");
                throw ex;
            }
        }
        #endregion
        #region GUARDAR ARCHIVO
        public void GuardarArchivo(string ubicacion, string nombre_archivo)
        {
            try
            {
                this.nombre_archivo = nombre_archivo;
                Console.WriteLine($"Nombre {nombre_archivo}");
                if (string.IsNullOrEmpty(ubicacion))
                    this.UbicacionDoc = Directory.GetCurrentDirectory();
                else
                    this.UbicacionDoc = ubicacion;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error al guardar el archivo {ex}");
                throw ex;
            }
            Dispos(true);
        }
        private bool Save()
        {
            try
            {
                
                var execel = $@"{UbicacionDoc}/Excel/{nombre_archivo} {DateTime.Now.ToString("dd-MM-yyyy hh-mm-s")}.xlsx";
                Console.WriteLine($"Ruta del Archivo {execel}");
                FileInfo excelFile = new FileInfo(execel);
                excelFile.Directory.Create();
                UbicacionDoc = excelFile.ToString();
                excel.SaveAs(excelFile);
                Dispos(true);
                return true;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error al guardar el archivo. {ex}");
                throw ex;
            }
        }
        #endregion
        #region LOGO PIE CONTEN
        private void Imagen()
        {
            if (RutaImagen != null)
            {
                using (Image image = Image.FromFile(RutaImagen))
                {
                    var excelImage = worksheet.Drawings.AddPicture("Logo", image);
                    excelImage.SetPosition(1, 0, 0, 0);
                }
            }
            Dispos(true);
        }
        private void PiePagina(int finalConte)
        {
            positionInicion = positionInicion + 7;
            CodigosDescripcion(positionInicion);
            if (piePagina != null)
            {
                if (piePagina.Count > 0)
                {
                    foreach (var obj in piePagina)
                    {
                        if (obj.Conten.Length == 13)
                        {
                            Texto($"{obj.Celda}{positionInicion}", obj.Conten);
                            Combinacion($"{obj.Celda}{positionInicion}:{celdaFinal}{positionInicion}");
                            Border(1, $"{obj.Celda}{positionInicion}:{celdaFinal}{positionInicion}");
                            positionInicion = positionInicion + 3;
                        }
                        else if (obj.Conten.Length > 20)
                        {
                          Texto($"{obj.Celda}{positionInicion}", obj.Conten);
                            int aumento = positionInicion;
                            aumento = aumento + 3; ;
                            Combinacion($"{obj.Celda}{positionInicion}:{celdaFinal}{aumento}");
                            ColorCelda($"{obj.Celda}{positionInicion}:{celdaFinal}{aumento}", Color.WhiteSmoke);
                            Border(1, $"{obj.Celda}{positionInicion}:{celdaFinal}{positionInicion}");
                            positionInicion = positionInicion + 4;
                        }
                        else
                        {
                           Texto($"{obj.Celda}{positionInicion}", obj.Conten);
                          Border(1, $"{obj.Celda}{positionInicion}:{celdaFinal}{positionInicion}");
                        }
                    }
                }
            }
            Dispos(true);
        }
        private bool Content()
        {
            try
            {
                string range = Convertir32(data);
                CargarData(range, data);
                GenerarCeldaFinal();
                PiePagina(GenerarBorder());
                Dispos(true);
                return true;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error en el conten {ex}");
                throw ex;
            }
        }
        #endregion
        #region UNICACION
        public Task<bool> Delete()
        {
            bool resp;
            try
            {
                File.Delete(UbicacionDoc.ToString());
                if (File.Exists(UbicacionDoc.ToString()))
                    resp = false;
                else
                    resp = true;
                Dispos(true);
                return Task.FromResult(resp);
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error en el delete {ex.StackTrace}");
                return Task.FromResult(false);
            }
        }
        public string Ubicacion()
        {
            try
            {
                Console.WriteLine($"ubicacion.. {UbicacionDoc}");
                return UbicacionDoc;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error de ubicacion {ex.StackTrace}");
                throw ex;
            }
        }
        #endregion
        #region CARGAR DATA Y FILTRO 
        private void Filtro(string range)
        {
            if (encabezados != null)
            {
                if (encabezados.Count <= 2)
                {
                    worksheet.Cells[range].AutoFilter = true;
                }
            }
            else
            {
                worksheet.Cells[range].AutoFilter = true;
            }
        }
        private void CargarData(string celda, List<string[]> datos)
        {
            
            try
            {
                worksheet.Cells[celda].LoadFromArrays(datos);
            }
            catch (Exception ex)
            {
                _log.Error($"Error al cargar  la data {ex.StackTrace}");

            }
            
        }
        #endregion
        #region HEADER Y ENCABEZADO
        private bool Header()
        {
            try
            {
                if (_proceso.Equals("convdeuda"))
                {
                    excel.Workbook.Worksheets.Add($"HOJA1");
                }
                else {
                    excel.Workbook.Worksheets.Add($"HOJA1 {DateTime.Now.ToString("dd-MM-yyyy")}");
                }
                string range = Convertir32(headerRow);
                if (_proceso.Equals("convdeuda"))
                {
                    worksheet = excel.Workbook.Worksheets[$"HOJA1"];
                }
                else
                {
                    worksheet = excel.Workbook.Worksheets[$"HOJA1 {DateTime.Now.ToString("dd-MM-yyyy")}"];
                }
               
                Filtro(range);
                Encabezado();
                CargarData(range, headerRow);
              AlineacionTexto(range, ExcelVerticalAlignment.Bottom, ExcelHorizontalAlignment.Left);
                ColorTexto(range, Color.WhiteSmoke, Color.Black, 12);
                TextoAjuste(1, range);
                positionInicion++;
                Dispos(true);
                return true;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error en el header {ex}");
                throw ex;
            }
        }
        private void Encabezado()
        {
            if (encabezados != null)
            {
                Imagen();
                if (encabezados.Count > 0)
                {
                    foreach (var obj in encabezados)
                    {
                        if (obj.Conten.Length > 20)
                        {
                            var cantidad = obj.Conten.Length / 15;
                            var celdaInicio = obj.Celda;
                            string combinacionceldas = $"{obj.Celda}{obj.PositionCelda}:";
                            obj.Celda += (char)(cantidad - 1);
                            combinacionceldas = $"{combinacionceldas}{obj.Celda}{obj.PositionCelda}";
                            Texto($"{celdaInicio}{obj.PositionCelda}", obj.Conten);
                            Combinacion(combinacionceldas);
                        }
                        else
                        {
                            Texto($"{obj.Celda}{obj.PositionCelda}", obj.Conten);

                        }
                    }
                }
            }
            else
            {
                if (celdaInicio.Equals('A') && positionInicion.Equals(2))
                {
                    Texto("A1", $"FECHA : {DateTime.Now.ToString("dd-MM-yyyy")}");
                    ColorTexto($"A1", Color.WhiteSmoke, Color.Black, 12);
                }
            }
            Dispos(true);
        }
        #endregion
        #region CONVERT 32
        private string Convertir32(List<string[]> datos)
        {
            return $"{celdaInicio}{positionInicion}:{char.ConvertFromUtf32(data[0].Length + 64)}{positionInicion}";
        }
        #endregion

        #region LIBERACION MEMORIA
        private void Limpiar()
        {
            using (MemoryStream me = new MemoryStream())
            {
                headerRow.Clear();
                data.Clear();
                worksheet = null;
                dataconte = null;
                attributes = null;
                properties = null;
                encabezados = null;
                piePagina = null;
                me.Dispose();
            }
            Dispos(true);
        }
        public void Dispos(bool reps)
        {
            if (resp) {
                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
            }
        }
        public void Dispose()
        {
            GC.Collect();
        }
        #endregion
        #region CODIGO DESCRIPCION
        private void CodigosDescripcion(int numeroCelda) {
            try {
                if (Cod != null)
                {
                    numeroCelda = numeroCelda - 5;
                    foreach (var cd in Cod)
                    {
                        Texto($"{cd.Celda}{numeroCelda}", cd.Conten);
                        Combinacion($"{cd.Celda}{numeroCelda}:{cd.Celda}{numeroCelda}");
                        Border(0, $"{cd.Celda}{numeroCelda}:{cd.Celda}{numeroCelda}");
                        numeroCelda = numeroCelda +1 ;
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Error de Inertar Codigos {ex}");
            }
        }
        public void CodigoDescrip(IList<Content> codigo)
        {
            this.Cod = codigo;
        }
        #endregion
        #endregion
        #region METODOS DISEÑO
        private void GenerarCeldaFinal()
        {
            celdaFinal += (char)(celdaInicio + data[0].Length - 1);
        }
        private int GenerarBorder()
        {
            for (int a = 0; a < data.Count(); a++)
            {
                Border(0, $"{celdaInicio}{positionInicion}:{celdaFinal}{positionInicion}");
                positionInicion++;
            }
            Dispos(true);
            return positionInicion;
        }
        private void ColorCelda(string celda, Color color)
        {
            try
            {
                worksheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[celda].Style.Fill.BackgroundColor.SetColor(color);
            }
            catch (Exception ex)
            {

                _log.Error($"Errror al asignar Color celda {ex.StackTrace}");
            }


        }
        private void Border(int position, string celda)
        {
            try
            {
                switch (position)
                {
                    case 0:
                        worksheet.Cells[celda].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[celda].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[celda].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[celda].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        break;
                    case 1:
                        worksheet.Cells[celda].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        break;
                }
                worksheet.Cells[celda].Style.Font.Bold = false;
            }
            catch (Exception ex)
            {
                _log.Error($"Error al asignar el borde a la celda  {ex.StackTrace}");
            }

        }
        private void TextoAjuste(int opcion, string celda)
        {
            try
            {
                switch (opcion)
                {
                    case 1:
                        worksheet.Cells[celda].AutoFitColumns();
                        break;
                }
            }
            catch (Exception ex)
            {
                _log.Error($"Error a ajustar el texto {ex.StackTrace}");

            }

        }
        private void ColorTexto(string celda, Color fondo, Color colorTexto, int size)
        {
            try
            {
                worksheet.Cells[celda].Style.Font.Bold = true;
                worksheet.Cells[celda].Style.Font.Size = size;
                worksheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[celda].Style.Fill.BackgroundColor.SetColor(fondo);
                worksheet.Cells[celda].Style.Font.Color.SetColor(colorTexto);
            }
            catch (Exception ex)
            {
                _log.Error($"Error a asignarle el color a la celda {ex.StackTrace}");

            }

        }
        private void AlineacionTexto(string celda, ExcelVerticalAlignment vertical, ExcelHorizontalAlignment horizontal)
        {
            try
            {
                worksheet.Cells[celda].Style.VerticalAlignment = vertical;
                worksheet.Cells[celda].Style.HorizontalAlignment = horizontal;
            }
            catch (Exception ex)
            {
                _log.Error($"Error a aliniar el texto {ex.StackTrace}");

            }

        }
        private void Texto(string celda, string texto)
        {
            try
            {
                worksheet.Cells[celda].Value = texto;
            }
            catch (Exception ex)
            {

                _log.Error($"Error del texto{ex.StackTrace}");
            }

        }
        private void Combinacion(string celda)
        {
            try
            {
                worksheet.Cells[celda].Merge = true;
                worksheet.Cells[celda].Style.WrapText = true;
            }
            catch (Exception ex)
            {
                _log.Error($"Error al conbinacion de texto{ex.StackTrace}");
            }

        }


        #endregion


    }
}
