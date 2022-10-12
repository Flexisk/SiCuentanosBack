using Dominio.Administrador;
using System.Net;

namespace WebApi.Responses
{
    public class ActividadResponse : GenericResponse
    {
        public Actividad actividad { get; set; }



        
        public ActividadResponse(HttpStatusCode codigo, string titulo, string mensaje, Actividad actividad):base(codigo, titulo, mensaje)
        {
            this.actividad = actividad;
        }



    }
}
