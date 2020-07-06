﻿using ExcelNugget02.Dtos;
using ExcelNugget02.Interfaces;
using log4net;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNugget02
{
    public class Excel : IExcel, IDisposable
    {
        #region CONSTRUCTOR
        public Excel()
        {
            this.celdaInicio = 'A';
            this.positionInicion = 2;
        }
        public Excel(string proceso)
        {
            if (proceso.Equals("convdeuda"))
            {
                celdaInicio = 'A';
                positionInicion = 1;
            }

        }
        #endregion
        #region Atributos
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
        private List<Content> encabezados;
        private List<Content> piePagina;
        private List<Content> Cod;
        private string RutaImagen;
        private ExtraerContent _extraerContent = new ExtraerContent();
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        #endregion
        #region METODOS
        #region PIE,ENCABEZADO,LOGO
        public void Encabezado(List<Content> encabezadoExcel)
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
        public void PiePagina(List<Content> PieExcel)
        {
            this.piePagina = PieExcel;
        }
        public void NombreLogo(string nombreImagen)
        {
            var ruta = $@"../../../Img/{nombreImagen}";
            this.RutaImagen = ruta;
        }
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
        #region CONTENIDO
        public Task<bool> NewContent<T>(List<T> datos)
        {
            try
            {
                if (datos.Count > 0)
                {
                    var cantidad = _extraerContent.GetHeader(datos.FirstOrDefault());
                    headerRow = _extraerContent.Data();
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
                    Main();
                }
                Dispos(true);
                return Task.FromResult(true);
            }
            catch (IOException ex)
            {
                _log.Error($"Error al cargar Data. {ex.StackTrace}"); 
                throw ex;
            }
        }
        private bool Content()
        {
            try
            {
                string range = Convertir32(data);
                CargarData(range, data);
                TextoAjuste(1, range);
                GenerarCeldaFinal();
                PiePagina(GenerarBorder());
                Dispos(true);
                return true;
            }
            catch (IOException ex)
            {
                throw ex;
            }
        }
        #endregion
        #region LIMPIAR
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
        #endregion
        #region LLAMADA METODOS 
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
                _log.InfoFormat("Archivo creado con exito");
                return resp;
            }
            catch (IOException ex)
            {
                _log.Warn($"Error al crear el excel {ex.StackTrace}");
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
                if (string.IsNullOrEmpty(ubicacion))
                    this.UbicacionDoc = Directory.GetCurrentDirectory();
                else
                    this.UbicacionDoc = ubicacion;
            }
            catch (IOException ex)
            {
                _log.Warn($"Error en la ruta del archivo {ex.StackTrace}");
                throw ex;
            }
            Dispos(true);
        }
        private bool Save()
        {
            try
            {
                var execel = $@"{UbicacionDoc}/Excel/{nombre_archivo} {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx";
                FileInfo excelFile = new FileInfo(execel);
                excelFile.Directory.Create();
                UbicacionDoc = excelFile.ToString();
                excel.SaveAs(excelFile);
                Dispos(true);
                return true;
            }
            catch (IOException ex)
            {
                _log.Error($"Error al guardar el archivo  {ex.StackTrace}");
                throw ex;
            }
        }
        #endregion
        #region CELDAS FINAL
        private void GenerarCeldaFinal()
        {
            celdaFinal += (char)(celdaInicio + data[0].Length - 1);
        }
        #endregion
        #region DISEÑO
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
            worksheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[celda].Style.Fill.BackgroundColor.SetColor(color);
        }
        private void Border(int position, string celda)
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
            Dispos(true);
        }
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
        private void TextoAjuste(int opcion, string celda)
        {
            switch (opcion)
            {
                case 1:
                    worksheet.Cells[celda].AutoFitColumns();
                    break;
            }
        }
        private void ColorTexto(string celda, Color fondo, Color colorTexto, int size)
        {
            worksheet.Cells[celda].Style.Font.Bold = true;
            worksheet.Cells[celda].Style.Font.Size = size;
            worksheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[celda].Style.Fill.BackgroundColor.SetColor(fondo);
            worksheet.Cells[celda].Style.Font.Color.SetColor(colorTexto);
        }
        private void AlineacionTexto(string celda, ExcelVerticalAlignment vertical, ExcelHorizontalAlignment horizontal)
        {
            worksheet.Cells[celda].Style.VerticalAlignment = vertical;
            worksheet.Cells[celda].Style.HorizontalAlignment = horizontal;
        }
        private void Texto(string celda, string texto)
        {
            worksheet.Cells[celda].Value = texto;
        }
        private void Combinacion(string celda)
        {
            worksheet.Cells[celda].Merge = true;
            worksheet.Cells[celda].Style.WrapText = true;
        }
        #endregion
        #region ELIMINAR ARCHIVO
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
                _log.Error($"Error al eliminar el archivo  {ex.StackTrace}");
                return Task.FromResult(false);
            }
        }
        #endregion
        #region UBICACION
        public string Ubicacion()
        {
            try
            {
                return UbicacionDoc;
            }
            catch (IOException ex)
            {
                _log.Error($"Error al guardar la ubicacion {ex.StackTrace}");
                throw ex;
            }
        }
        #endregion
        #region HEADER
        private bool Header()
        {
            try
            {
                excel.Workbook.Worksheets.Add($"HOJA1 {DateTime.Now.ToString("dd-MM-yyyy")}");
                string range = Convertir32(headerRow);
                worksheet = excel.Workbook.Worksheets[$"HOJA1 {DateTime.Now.ToString("dd-MM-yyyy")}"];
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
                throw ex;
            }
        }
        #endregion
        #region CONVERTIR DATA Y CARGAR
        private void CargarData(string celda, List<string[]> datos)
        {
            worksheet.Cells[celda].LoadFromArrays(datos);
        }
        private string Convertir32(List<string[]> datos)
        {
            return $"{celdaInicio}{positionInicion}:{char.ConvertFromUtf32(data[0].Length + 64)}{positionInicion}";
        }
        #endregion
        #region LIBERACION MEMORIA
        public void Dispos(bool reps)
        {
            if (resp)
            {
                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
            }
        }
        public void Dispose()
        {
            GC.Collect();
        }
        #endregion
        #region DESCRIPCION
        private void CodigosDescripcion(int numeroCelda)
        {
            try
            {
                if (Cod != null)
                {
                    numeroCelda = numeroCelda - 5;
                    foreach (var cd in Cod)
                    {
                        Texto($"{cd.Celda}{numeroCelda}", cd.Conten);
                        Combinacion($"{cd.Celda}{numeroCelda}:{cd.Celda}{numeroCelda}");
                        Border(0, $"{cd.Celda}{numeroCelda}:{cd.Celda}{numeroCelda}");
                        numeroCelda = numeroCelda + 1;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Warn($"Error al insertar codigo {ex.StackTrace}");
            }
        }
        public void CodigoDescrip(List<Content> codigo)
        {
            this.Cod = codigo;
        }
        #endregion
        #endregion
    }
}

