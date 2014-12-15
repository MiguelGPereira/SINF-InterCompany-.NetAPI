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

    }
}
