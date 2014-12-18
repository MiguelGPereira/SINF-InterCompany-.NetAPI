using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FirstREST.Lib_Primavera.Model
{
    public class Fornecedor
    {
        //
        // GET: /Fornecedor/
        public string CodFornecedor
        {
            get;
            set;
        }

        public string NomeFornecedor
        {
            get;
            set;
        }

        public string NumContribuinte
        {
            get;
            set;
        }

        public string NomeFiscal
        {
            get;
            set;
        }

        public string Moeda
        {
            get;
            set;
        }
    }
}
