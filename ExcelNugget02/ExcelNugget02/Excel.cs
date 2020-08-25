using ExcelNugget02.Class;
using ExcelNugget02.Dtos;
using ExcelNugget02.Interfaces;
using log4net;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Polly;
using Polly.Retry;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNugget02
{
    public class Excel : IExcel
    {
        #region Atributos
        private readonly string _proceso = null;
        private ExtraerContent extra = new ExtraerContent();
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
        private string UbicacionDoc;
        private readonly Fecha _fecha = new Fecha();
        private readonly string _ubicacion = null;
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private readonly FileBase64 _fileBase64 = new FileBase64();
        #endregion

        #region constructor
        public Excel()
        {
            celdaInicio = 'A';
            positionInicion = 2;
        }

        #region PROCESO GENERAR DEUDA
        public Excel(string proceso)
        {
            if (proceso.Equals("deuda"))
            {
                _proceso = proceso;
                celdaInicio = 'A';
                positionInicion = 1;
            }

        }
        #endregion
        #endregion

        #region CELDA FINAL Y ENCABEZADO
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

            Dispose(true);
        }
        #endregion

        #region CONTENIDO

        #region GENERA EXCEL
        public Task<bool> NewContent<T>(List<T> datos, string hoja)
        {
            bool _resp = false;
            try
            {
                var policyExcel = RetryPolicy.Handle<Exception>().Or<NullReferenceException>().
                   WaitAndRetry(4, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)), (ex, time) =>
                   {
                       _log.Warn($"Intent Para crear el excel {time.Seconds}, {_fecha.FechaNow().Result}");
                   });

                policyExcel.Execute(() => {
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
                        bool resp = Header(hoja).Result;
                        if (resp)
                            resp = Content().Result;
                        _log.Info($"Archivo excel creado con exito {_fecha.FechaNow().Result}");
                    }
                });
            }
            catch (Exception ex)
            {
                _log.Error($"Excel {ex.StackTrace}");
            }
            Limpiar();
            Dispose(true);
            return Task.FromResult(_resp);
        }
        #endregion

        #region CONTENIDO
        private Task<bool> Content()
        {
            bool _resp = false;

            try
            {
                var policyConten = RetryPolicy.Handle<Exception>().Or<NullReferenceException>().
                    WaitAndRetry(4, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)), (ex, time) =>
                    {
                        _log.Fatal($"Intento para crear el contenido del excel, {time.Seconds} {_fecha.FechaNow().Result}");
                    });

                policyConten.Execute(() =>
                {
                    string range = Convertir32(data);
                    worksheet.Cells[range].LoadFromArrays(data);
                    GenerarCeldaFinal();
                    if (_proceso == null)
                        GenerarBorder();
                    positionInicion--;
                    _resp = true;
                });
            }
            catch (Exception ex)
            {
                _log.Warn($"Excepcion {ex.StackTrace}");
            }
            return Task.FromResult(_resp);
        }
        #endregion

        #region HEADER
        private Task<bool> Header(string nombrehoja)
        {
            bool _resp = false;
            try
            {
                var policyHeader = RetryPolicy.Handle<Exception>().Or<NullReferenceException>().
                    WaitAndRetry(4, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)), (excel, time) => {
                        _log.Warn($"Intento para crear el Header, {time.Seconds}  {_fecha.FechaNow().Result}");
                    });

                policyHeader.Execute(() => {
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
                    _log.Info($"Creacion con exito de los Headers de las columnas.");
                    _resp = true;
                });
            }
            catch (Exception ex)
            {
                _log.Error($"Excepcion {ex.StackTrace}");
            }
            return Task.FromResult(_resp);
        }
        #endregion

        #endregion

        #region RUTA ARCHIVO
        #region CREAR DIRECTORIO
        private void Directorio()
        {
            try
            {
                var policyDirectorio = RetryPolicy.Handle<Exception>().Or<NullReferenceException>().
                    WaitAndRetry(4, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)), (ex, time) =>
                    {
                        _log.Warn($"Intento para crear el directorio,{time.Seconds} !!  {_fecha.FechaNow().Result}");
                    });

                policyDirectorio.Execute(() => {
                    if (string.IsNullOrEmpty(_ubicacion))
                        this.UbicacionDoc = Directory.GetCurrentDirectory();
                    else
                        this.UbicacionDoc = _ubicacion;
                    _log.Info($"Directorio creado con exito {_fecha.FechaNow().Result}");
                });
            }
            catch (Exception ex)
            {
                _log.Error($"Exception {ex.StackTrace}");
            }
        }
        #endregion

        #region Guardar el archivo
        public Task<FileBase64> Guardar(string FileName)
        {
            FileBase64 _fileBase64 = new FileBase64();
            try
            {
                var policySave = RetryPolicy.Handle<Exception>().Or<NullReferenceException>().
                       WaitAndRetry(4, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)), (ex, time) =>
                       {
                           _log.Warn($"Intenton  Para guardar el archivo excel {time.Seconds}, {_fecha.FechaNow().Result}");
                       });

                policySave.Execute(() => {
                    Directorio();
                    var excelUbicacion = $@"{UbicacionDoc}/Excel/{FileName}.xlsx";


                    FileInfo excelFile = new FileInfo(excelUbicacion);
                    excelFile.Directory.Create();
                    excel.SaveAs(excelFile);
                    _log.Info($"Archivo guardado con exito {_fecha.FechaNow().Result}");
                    _fileBase64 = new FileBase64()
                    {
                        FileName = Path.GetFileName(excelUbicacion),
                        Base64Data = Convert.ToBase64String(File.ReadAllBytes(excelUbicacion))
                    };
                    _log.Info($"Proceso de conversion Base64 {_fecha.FechaNow().Result}");
                    // File.Delete(excelUbicacion);
                    // _log.Info($"Archivo elminado con exito {_fecha.FechaNow().Result}");

                });
            }
            catch (Exception ex)
            {
                _log.Warn($"Exception {ex.StackTrace}");
            }
            Dispose(true);
            return Task.FromResult(_fileBase64);
        }
        #endregion
        #endregion

        #region CONVERT 32
        private string Convertir32(List<string[]> datos)
        {
            return $"{celdaInicio}{positionInicion}:{char.ConvertFromUtf32(data[0].Length + 64)}{positionInicion}";
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
            Dispose(true);
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
                _log.Info("Cargo la Data con exito..");
            }
            catch (Exception ex)
            {
                _log.Error($"Error al cargar  la data {ex.StackTrace}");

            }

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
            Dispose(true);
        }

        public void Dispose(bool reps)
        {
            if (reps)
            {
                Dispose();
            }
        }
        public void Dispose()
        {
            GC.Collect(2, GCCollectionMode.Forced);
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
        }


        #endregion

    }
}

