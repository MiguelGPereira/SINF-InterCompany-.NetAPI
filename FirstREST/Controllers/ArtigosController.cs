using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using FirstREST.Lib_Primavera.Model;


namespace FirstREST.Controllers
{
    public class ArtigosController : ApiController
    {
        //
        // GET: /Artigos/

        public IEnumerable<Lib_Primavera.Model.Artigo> Get(string codEmpresa)
        {
            return Lib_Primavera.Comercial.ListaArtigos(codEmpresa);
        }


        // GET api/artigo/5    
        public Artigo Get(string codEmpresa, string idArtigo)
        {
            Lib_Primavera.Model.Artigo artigo = Lib_Primavera.Comercial.GetArtigo(codEmpresa, idArtigo);
            if (artigo == null)
            {
                throw new HttpResponseException(
                  Request.CreateResponse(HttpStatusCode.NotFound));
            }
            else
            {
                return artigo;
            }
        }

    }
}

