using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace PrueebasNugguet2.Interfaces
{
   public interface IExcel
    {
        Task<bool> NewContent<T>(IList<T> datos);
        Task<bool> Delete();
        string Ubicacion();
        void GuardarArchivo(string ubicacion, string nombre_archivo);
        void Encabezado(IList<Content> encabezadoExcel);
        void PiePagina(IList<Content> PieExcel);
        void NombreLogo(string nombreImagen);
        void CodigoDescrip(IList<Content> codigo);

    }
}
