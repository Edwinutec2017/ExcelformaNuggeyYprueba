using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Text;

namespace ExcelNugget02.Dtos
{
   public class ExtraerContent:IDisposable
    {
        #region ATRIBUTOS
        private PropertyInfo[] properties = null;
        private object[] attributes = null;
        private List<string[]> headerRow = new List<string[]>();
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        #endregion
        #region LIBERACION
        public void Dispose()
        {
            GC.Collect(2, GCCollectionMode.Forced);
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
        }
        #endregion
        #region EXTRAER CONTENIDO
        private int CantidadMostrar(PropertyInfo[] properties)
        {
            try
            {
                var cantidad = properties.Select(property => ConvertObject(property).Length > 0
                ? !((DescripcionExcel)ConvertObject(property).FirstOrDefault()).Ignore : true)
                .Where(z => z).Count();
                Dispose();
                return cantidad;
            }
            catch (IOException ex)
            {
                _log.Warn($"Error al extraer la informacion {ex.StackTrace}");
                throw ex;
            }
        }
        private object[] ConvertObject(PropertyInfo property)
        {
            return property.GetCustomAttributes(typeof(DescripcionExcel), true);
        }

        public List<string[]> Data()
        {
            return headerRow;
        }
        public int GetHeader(object obj)
        {

            try
            {
                properties = obj.GetType().GetProperties();
                string[] header = new string[CantidadMostrar(properties)];
                var indice = 0;
                foreach (PropertyInfo property in properties)
                {
                    attributes = property.GetCustomAttributes(typeof(DescripcionExcel), true);
                    DescripcionExcel myAttribute = (DescripcionExcel)attributes[0];
                    if (!myAttribute.Ignore)
                    {
                        header[indice] = (!string.IsNullOrEmpty(myAttribute.Name)) ? myAttribute.Name : property.Name.ToUpper();
                    }
                    else
                    {
                        indice--;
                    }
                    indice++;
                }
                headerRow.Add(header);
                Dispose();
                return header.Length;
            }
            catch (IOException ex)
            {
                _log.Error($"Error en al estraer  GetHeader  {ex.StackTrace}");
                throw ex;
            }
        }
        #endregion

    }
}
