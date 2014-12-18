using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using FirstREST.Lib_Primavera.Model;

namespace FirstREST.Controllers
{
    public class FornecedoresController : ApiController
    {
        //
        // GET: /Fornecedores/

        public IEnumerable<Lib_Primavera.Model.Fornecedor> Get(string codEmpresa)
        {
            return Lib_Primavera.Comercial.ListaFornecedores(codEmpresa);
        }

        public Fornecedor Get(string codEmpresa, string idFornecedor)
        {
            Lib_Primavera.Model.Fornecedor fornecedor = Lib_Primavera.Comercial.GetFornecedor(codEmpresa, idFornecedor);
            if (fornecedor == null)
            {
                throw new HttpResponseException(
                        Request.CreateResponse(HttpStatusCode.NotFound));

            }
            else
            {
                return fornecedor;
            }
        }


        public HttpResponseMessage Post(string id, Lib_Primavera.Model.Fornecedor fornecedor)
        {
            Lib_Primavera.Model.RespostaErro erro = new Lib_Primavera.Model.RespostaErro();
            erro = Lib_Primavera.Comercial.InsereFornecedor(id, fornecedor);

            if (erro.Erro == 0)
            {
                var response = Request.CreateResponse(
                   HttpStatusCode.Created, fornecedor);
                string uri = Url.Link("DefaultApi", new { CodFornecedor = fornecedor.CodFornecedor });
                response.Headers.Location = new Uri(uri);
                return response;
            }

            else
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest);
            }

        }

    }
}
