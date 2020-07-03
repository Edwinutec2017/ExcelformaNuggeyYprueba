using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace PrueebasNugguet2.Dto
{
   public  class ExtraerConten:IDisposable
    {

        private PropertyInfo[] properties = null;
        private object[] attributes = null;
        private List<string[]> headerRow = new List<string[]>();


        public void Dispose()
        {
            throw new NotImplementedException();
        }

        private int CantidadMostrar(PropertyInfo[] properties)
        {
            try
            {
                var cantidad = properties.Select(property => ConvertObject(property).Length > 0
                ? !((DescripcionExcel)ConvertObject(property).FirstOrDefault()).Ignore : true)
                .Where(z => z).Count();
                return cantidad;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error en la cantidad a mostrar {ex}");
                throw ex;
            }
        }
        private object[] ConvertObject(PropertyInfo property)
        {
            return property.GetCustomAttributes(typeof(DescripcionExcel), true);
        }

        public List<string[]> Data() {
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
                //Dispos(true);
                return header.Length;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error en el GeHeader {ex}");
                throw ex;
            }
        }
    }
}
