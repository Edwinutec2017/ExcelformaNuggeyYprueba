using ExcelNugget02.Dtos;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNugget02.Interfaces
{
  public interface IExcel
    {
        Task<bool> NewContent<T>(List<T> datos, string hoja);
        Task<bool> Delete();
        string Guardar();
        void ArchivoRuta(string ubicacion, string nombre_archivo);
    }
}
