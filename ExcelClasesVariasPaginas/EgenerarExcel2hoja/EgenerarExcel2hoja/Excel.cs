using EgenerarExcel2hoja.Dto;
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

namespace EgenerarExcel2hoja
{
   public class Excel:IDisposable
    {
        public Excel()
        {
            celdaInicio = 'A';
            positionInicion = 2;
        }
        public Excel(string proceso)
        {
            if (proceso.Equals("convdeuda"))
            {
                _proceso = proceso;
                celdaInicio = 'A';
                positionInicion = 1;
            }

        }

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
        private string nombre_archivo = "";
        private string UbicacionDoc;
       
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        #endregion
        private void GenerarCeldaFinal()
        {
           
            celdaFinal = (char)(celdaInicio + data[0].Length - 1);
        }
        private void Encabezado()
        {
                if (celdaInicio.Equals('A') && positionInicion.Equals(2))
                {
                    Texto("A1", $"FECHA : {DateTime.Now.ToString("dd-MM-yyyy")}");
                    ColorTexto($"A1", Color.WhiteSmoke, Color.Black, 12);
                }
            
            Dispos(true);
        }
        public Task<bool> NewContent<T>(IList<T> datos,string hoja )
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
                    Header(hoja);
                    Content();
                    Limpiar();
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
        private bool Content()
        {
            try
            {
                string range = Convertir32(data);
                worksheet.Cells[range].LoadFromArrays(data);
                GenerarCeldaFinal();
                if(_proceso==null)
                GenerarBorder();
                positionInicion--;
                Dispos(true);
                return true;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error en el conten {ex}");
                throw ex;
            }
        }
        public void ArchivoRuta(string ubicacion, string nombre_archivo)
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
        public string Guardar()
        {
            string RutaUbicacion = null;
            try
            {

                var execel = $@"{UbicacionDoc}/Excel/{nombre_archivo}.xlsx";
                FileInfo excelFile = new FileInfo(execel);
                excelFile.Directory.Create();
                UbicacionDoc = excelFile.ToString();
                excel.SaveAs(excelFile);
                Dispos(true);
                RutaUbicacion = UbicacionDoc;
                
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error al guardar el archivo. {ex}");
                throw ex;
            }
            return RutaUbicacion;
        }
        private bool Header(string nombrehoja)
        {
            try
            {
                excel.Workbook.Worksheets.Add(nombrehoja);
                string range = Convertir32(headerRow);
                worksheet = excel.Workbook.Worksheets[nombrehoja];
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
                me.Dispose();
            }
            Dispos(true);
        }
        public void Dispos(bool reps)
        {
            if (reps)
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

        #region METODOS DISEÑO
        private int GenerarBorder()
        {
            int position = positionInicion;
            for (int a = 0; a < data.Count(); a++)
            {
                Border(0, $"{celdaInicio}{position}:{celdaFinal}{position}");
                position++;
            }
            position = 0;
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
        #endregion
        #region CARGAR DATA Y FILTRO 
        private void Filtro(string range)
        {

                worksheet.Cells[range].AutoFilter = true;
            
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
    }
}
