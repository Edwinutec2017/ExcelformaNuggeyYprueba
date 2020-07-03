using ExcelNugget02.Dtos;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNugget02.Interfaces
{
  public interface IExcel
    {
        Task<bool> NewContent<T>(List<T> datos);
        Task<bool> Delete();
        string Ubicacion();
        void GuardarArchivo(string ubicacion, string nombre_archivo);
        void Encabezado(List<Content> encabezadoExcel);
        void PiePagina(List<Content> PieExcel);
        void NombreLogo(string nombreImagen);
        void CodigoDescrip(List<Content> codigo);

    }
}
