using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Interop.ErpBS800;
using Interop.StdPlatBS800;
using Interop.StdBE800;
using Interop.GcpBE800;
using ADODB;
using Interop.IGcpBS800;
using System.IO;
//using Interop.StdBESql800;
//using Interop.StdBSSql800;

/*
 * GET  /api/empresa
GET  /api/empresa/:id
GET  /api/empresa/:id/artigo
GET  /api/empresa/:id/artigo/:id
GET  /api/empresa/:id/fornecedor
GET  /api/empresa/:id/fornecedor/:id
GET  /api/empresa/:id/cliente
GET  /api/empresa/:id/cliente/:id
POST /api/empresa/:id/encomendaCliente
GET  /api/empresa/:id/encomendaCliente
GET  /api/empresa/:id/encomendaCliente/:id
POST /api/empresa/:id/encomendaFornecedor
GET  /api/empresa/:id/encomendaFornecedor
GET  /api/empresa/:id/encomendaFornecedor/:id
POST /api/empresa/:id/faturaVenda
GET  /api/empresa/:id/faturaVenda
GET  /api/empresa/:id/faturaVenda/:id
POST /api/empresa/:id/faturaCompra
GET  /api/empresa/:id/faturaCompra
GET  /api/empresa/:id/faturaVenda/:id
 */

namespace FirstREST.Lib_Primavera
{

    public class Comercial
    {

        static string filePath = "~/";

       /* string pathEmpresasDB = @"C:\ADMIN\empresas.txt";
        string pathArtigosDB;*/

        //--------------EMPRESAS

        #region Empresas
        //OK
        public static List<Model.Empresa> ListaEmpresas()
        {
            List<Model.Empresa> listEmpresas = new List<Model.Empresa>();
            
            Model.Empresa emp = new Model.Empresa();

            //Microsoft ftw
            System.IO.StreamReader file = new System.IO.StreamReader((System.Web.HttpContext.Current.Server.MapPath(filePath) + "empresas.txt").Replace("/","//"));
            string line;

            while ((line = file.ReadLine()) != null)
            {
                emp = new Model.Empresa();
                emp.codEmpresa = line.Split(';')[0];
                emp.nomeEmpresa = line.Split(';')[1];

                //ABRIR EMPRESA: RETORNAR O SEU NOME
              /*  if (PriEngine.InitializeCompany(line, "", "") == true)
                {
                    StdBELista dadosEmpresa = new StdBELista();
                    dadosEmpresa = PriEngine.Engine.Consulta("SELECT");
                    emp.nomeEmpresa = dadosEmpresa.Valor("nome");
                }*/


                listEmpresas.Add(emp);
            }

            file.Close();

            return listEmpresas;
        }

        #endregion Empresa; 
        //--------------------------

        # region Cliente
        //OK!
        public static List<Model.Cliente> ListaClientes(string codEmpresa)
        {
            ErpBS objMotor = new ErpBS();
             
            StdBELista objList;

            Model.Cliente cli = new Model.Cliente();
            List<Model.Cliente> listClientes = new List<Model.Cliente>();


            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                objList = PriEngine.Engine.Consulta("SELECT Cliente, Nome, Moeda, NumContrib as NumContribuinte FROM  CLIENTES");

                while (!objList.NoFim())
                {
                    cli = new Model.Cliente();
                    cli.CodCliente = objList.Valor("Cliente");
                    cli.NomeCliente = objList.Valor("Nome");
                    cli.Moeda = objList.Valor("Moeda");
                    cli.NumContribuinte = objList.Valor("NumContribuinte");

                    listClientes.Add(cli);
                    objList.Seguinte();

                }

                return listClientes;
            }
            else
                return null;
        }

        //OK!
        public static Lib_Primavera.Model.Cliente GetCliente(string codEmpresa, string idCliente)
        {
            ErpBS objMotor = new ErpBS();
            StdBELista objCli = new StdBELista();
            Model.Cliente myCli = new Model.Cliente();

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                if (PriEngine.Engine.Comercial.Clientes.Existe(idCliente) == true)
                {
                    objCli = PriEngine.Engine.Consulta("SELECT Cliente, Nome, Moeda, NumContrib as NumContribuinte FROM CLIENTES WHERE Cliente = '" + idCliente + "'");
                    myCli.CodCliente = objCli.Valor("Cliente");
                    myCli.NomeCliente = objCli.Valor("Nome");
                    myCli.Moeda = objCli.Valor("Moeda");
                    myCli.NumContribuinte = objCli.Valor("NumContribuinte");
                    return myCli;
                }
                else
                {
                    return null;
                }
            }
            else
                return null;
        }

        //NOT NECESSARY FOR NOW
        public static Lib_Primavera.Model.RespostaErro UpdCliente(Lib_Primavera.Model.Cliente cliente)
        {



            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();
            ErpBS objMotor = new ErpBS();

            GcpBECliente objCli = new GcpBECliente();

            try
            {

                if (PriEngine.InitializeCompany("EMP1", "", "") == true)
                {

                    if (PriEngine.Engine.Comercial.Clientes.Existe(cliente.CodCliente) == false)
                    {
                        erro.Erro = 1;
                        erro.Descricao = "O cliente não existe";
                        return erro;
                    }
                    else
                    {

                        objCli = PriEngine.Engine.Comercial.Clientes.Edita(cliente.CodCliente);
                        objCli.set_EmModoEdicao(true);

                        objCli.set_Nome(cliente.NomeCliente);
                        objCli.set_NumContribuinte(cliente.NumContribuinte);
                        objCli.set_Moeda(cliente.Moeda);

                        PriEngine.Engine.Comercial.Clientes.Actualiza(objCli);

                        erro.Erro = 0;
                        erro.Descricao = "Sucesso";
                        return erro;
                    }
                }
                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir a empresa";
                    return erro;

                }

            }

            catch (Exception ex)
            {
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }

        }
         

        //NOT NECESSARY FOR NOW
        public static Lib_Primavera.Model.RespostaErro DelCliente(string idCliente)
        {

            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();
            GcpBECliente objCli = new GcpBECliente();


            try
            {

                if (PriEngine.InitializeCompany("EMP1", "", "") == true)
                {
                    if (PriEngine.Engine.Comercial.Clientes.Existe(idCliente) == false)
                    {
                        erro.Erro = 1;
                        erro.Descricao = "O cliente não existe";
                        return erro;
                    }
                    else
                    {

                        PriEngine.Engine.Comercial.Clientes.Remove(idCliente);
                        erro.Erro = 0;
                        erro.Descricao = "Sucesso";
                        return erro;
                    }
                }

                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir a empresa";
                    return erro;
                }
            }

            catch (Exception ex)
            {
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }

        }
        

        public static Lib_Primavera.Model.RespostaErro InsereClienteObj(Model.Cliente cli, string codEmpresa)
        {

            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();
         
            GcpBECliente myCli = new GcpBECliente();

            try
            {
                if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
                {

                    myCli.set_Cliente(cli.CodCliente);
                    myCli.set_Nome(cli.NomeCliente);
                    myCli.set_NumContribuinte(cli.NumContribuinte);
                    myCli.set_Moeda(cli.Moeda);

                    PriEngine.Engine.Comercial.Clientes.Actualiza(myCli);

                    erro.Erro = 0;
                    erro.Descricao = "Sucesso";
                    return erro;
                }
                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir empresa";
                    return erro;
                }
            }

            catch (Exception ex)
            {
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }


        }

        /*
        public static void InsereCliente(string codCliente, string nomeCliente, string numContribuinte, string moeda)
        {
            ErpBS objMotor = new ErpBS();
            MotorPrimavera mp = new MotorPrimavera();

            GcpBECliente myCli = new GcpBECliente();

            objMotor = mp.AbreEmpresa("DEMO", "", "", "Default");

            myCli.set_Cliente(codCliente);
            myCli.set_Nome(nomeCliente);
            myCli.set_NumContribuinte(numContribuinte);
            myCli.set_Moeda(moeda);

            objMotor.Comercial.Clientes.Actualiza(myCli);

        }


        */


        #endregion Cliente;   // -----------------------------  END   CLIENTE    -----------------------

        # region Fornecedor 
        // ----------------------- FORNECEDOR

        //OK!
        public static List<Model.Fornecedor> ListaFornecedores(string codEmpresa)
        {
            ErpBS objMotor = new ErpBS();

            StdBELista objList;

            Model.Fornecedor forn = new Model.Fornecedor();
            List<Model.Fornecedor> listFornecedores = new List<Model.Fornecedor>();


            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                objList = PriEngine.Engine.Consulta("SELECT Fornecedor, Nome, NumContrib, NomeFiscal FROM Fornecedores");

                while (!objList.NoFim())
                {
                    forn = new Model.Fornecedor();
                    forn.CodFornecedor = objList.Valor("Fornecedor");
                    forn.NomeFornecedor = objList.Valor("Nome");
                    forn.NomeFiscal = objList.Valor("NomeFiscal");
                    forn.NumContribuinte = objList.Valor("NumContrib");

                    listFornecedores.Add(forn);
                    objList.Seguinte();

                }

                return listFornecedores;
            }
            else
                return null;
        }

        //OK!
        public static Model.Fornecedor GetFornecedor(string codEmpresa, string idFornecedor)
        {

            ErpBS objMotor = new ErpBS();

            StdBELista objForn = new StdBELista();

            Model.Fornecedor myForn = new Model.Fornecedor();

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                if (PriEngine.Engine.Comercial.Fornecedores.Existe(idFornecedor) == true)
                {

                    objForn = PriEngine.Engine.Consulta("SELECT Fornecedor, Nome, NumContrib, NomeFiscal FROM Fornecedores WHERE Fornecedor = '" + idFornecedor + "'");
                    myForn.CodFornecedor = objForn.Valor("Fornecedor");
                    myForn.NomeFornecedor = objForn.Valor("Nome");
                    myForn.NomeFiscal = objForn.Valor("NomeFiscal");
                    myForn.NumContribuinte = objForn.Valor("NumContrib");
                    return myForn;
                }
                else
                {
                    return null;
                }
            }
            else
                return null;

        }

        public static Lib_Primavera.Model.RespostaErro InsereFornecedorObj(Model.Fornecedor forn, string codEmpresa)
        {

            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();

            GcpBEFornecedor myForn = new GcpBEFornecedor();

            try
            {
                if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
                {

                    myForn.set_Fornecedor(forn.CodFornecedor);
                    myForn.set_Nome(forn.NomeFornecedor);
                    myForn.set_NumContribuinte(forn.NumContribuinte);
                    myForn.set_NomeFiscal(forn.NomeFiscal);

                    PriEngine.Engine.Comercial.Fornecedores.Actualiza(myForn);

                    erro.Erro = 0;
                    erro.Descricao = "Sucesso";
                    return erro;
                }
                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir empresa";
                    return erro;
                }
            }

            catch (Exception ex)
            {
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }


        }

        #endregion Fornecedor; // ----------------------------- END FORNECEDOR -----------------------

        #region Artigo

        //OK!
        public static Lib_Primavera.Model.Artigo GetArtigo(string codEmpresa, string idArtigo)
        {

            StdBELista objArtigo = new StdBELista();
            Model.Artigo myArt = new Model.Artigo();

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                if (PriEngine.Engine.Comercial.Artigos.Existe(idArtigo) == true)
                {

                    objArtigo = PriEngine.Engine.Consulta("SELECT Artigo.Artigo, Artigo.Descricao, Artigo.UnidadeVenda, Artigo.Iva, ArtigoMoeda.Moeda AS Moeda, ArtigoMoeda.PVP1 AS PVP1 FROM Artigo INNER JOIN ArtigoMoeda ON Artigo.Artigo = ArtigoMoeda.Artigo WHERE Artigo.Artigo = '" + idArtigo + "'");
                    myArt.CodArtigo = objArtigo.Valor("Artigo");
                    myArt.DescArtigo = objArtigo.Valor("Descricao");
                    myArt.UnidadeArtigo = objArtigo.Valor("UnidadeVenda");
                    myArt.IVAArtigo = objArtigo.Valor("Iva");
                    myArt.PrecoArtigo = objArtigo.Valor("PVP1");
                    myArt.MoedaArtigo = objArtigo.Valor("Moeda");
                    return myArt;
                }
                else
                {
                    return null;
                }
            }
            else
                return null;
        }

        //OK!
        public static List<Model.Artigo> ListaArtigos(string codEmpresa)
        {

            ErpBS objMotor = new ErpBS();

            StdBELista objList;

            Model.Artigo artigo = new Model.Artigo();
            List<Model.Artigo> listArtigos = new List<Model.Artigo>();


            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                objList = PriEngine.Engine.Consulta("SELECT Artigo.Artigo, Artigo.Descricao, Artigo.UnidadeVenda, Artigo.Iva, ArtigoMoeda.Moeda AS Moeda, ArtigoMoeda.PVP1 AS PVP1 FROM Artigo INNER JOIN ArtigoMoeda ON Artigo.Artigo = ArtigoMoeda.Artigo");

                while (!objList.NoFim())
                {
                    artigo = new Model.Artigo();
                    artigo.CodArtigo = objList.Valor("Artigo");
                    artigo.DescArtigo = objList.Valor("Descricao");
                    artigo.UnidadeArtigo = objList.Valor("UnidadeVenda");
                    artigo.IVAArtigo = objList.Valor("Iva");
                    artigo.PrecoArtigo = objList.Valor("PVP1");
                    artigo.MoedaArtigo = objList.Valor("Moeda");

                    listArtigos.Add(artigo);
                    objList.Seguinte();

                }

                return listArtigos;
            }
            else
                return null;
        }



        #endregion Artigo;

        //------------------------------------ ENCOMENDA ---------------------
        #region Encomenda

        /*
        public static Model.RespostaErro TransformaDoc(Model.DocCompra dc)
        {

            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();
            GcpBEDocumentoCompra objEnc = new GcpBEDocumentoCompra();
            GcpBEDocumentoCompra objGR = new GcpBEDocumentoCompra();
            GcpBELinhasDocumentoCompra objLinEnc = new GcpBELinhasDocumentoCompra();
            PreencheRelacaoCompras rl = new PreencheRelacaoCompras();

            List<Model.LinhaDocCompra> lstlindc = new List<Model.LinhaDocCompra>();

            try
            {
                if (PriEngine.InitializeCompany("EMP1", "", "") == true)
                {
                

                    objEnc = PriEngine.Engine.Comercial.Compras.Edita("000", "ECF", "2013", 3);

                    // --- Criar os cabeçalhos da GR
                    objGR.set_Entidade(objEnc.get_Entidade());
                    objEnc.set_Serie("2013");
                    objEnc.set_Tipodoc("ECF");
                    objEnc.set_TipoEntidade("F");

                    objGR = PriEngine.Engine.Comercial.Compras.PreencheDadosRelacionados(objGR,rl);
 

                    // façam p.f. o ciclo para percorrer as linhas da encomenda que pretendem copiar
                     
                        double QdeaCopiar;
                        PriEngine.Engine.Comercial.Internos.CopiaLinha("C", objEnc, "C", objGR, lin.NumLinha, QdeaCopiar);
                       
                        // precisamos aqui de um metodo que permita actualizar a Qde Satisfeita da linha de encomenda.  Existe em VB mas ainda não sei qual é em c#
                       
                    PriEngine.Engine.IniciaTransaccao();
                    PriEngine.Engine.Comercial.Compras.Actualiza(objEnc, "");
                    PriEngine.Engine.Comercial.Compras.Actualiza(objGR, "");

                    PriEngine.Engine.TerminaTransaccao();

                    erro.Erro = 0;
                    erro.Descricao = "Sucesso";
                    return erro;
                }
                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir empresa";
                    return erro;

                }

            }
            catch (Exception ex)
            {
                PriEngine.Engine.DesfazTransaccao();
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }
        
        
        }

        */

        #endregion Encomenda;


        // ------------------------ Documentos de Compra --------------------------//
        #region DocCompra

        public static List<Model.DocCompra> ListaDocumentosCompra(string codEmpresa, string tipoDeDocumento)
        {
            ErpBS objMotor = new ErpBS();
            
            StdBELista objListCab;
            StdBELista objListLin;
            Model.DocCompra dc = new Model.DocCompra();
            List<Model.DocCompra> listdc = new List<Model.DocCompra>();
            Model.LinhaDocCompra lindc = new Model.LinhaDocCompra();
            List<Model.LinhaDocCompra> listlindc = new List<Model.LinhaDocCompra>();

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {
                objListCab = PriEngine.Engine.Consulta("SELECT id, NumDocExterno, Entidade, DataDoc, NumDoc, TotalMerc, Serie From CabecCompras where TipoDoc='"+ tipoDeDocumento +"'");
                while (!objListCab.NoFim())
                {
                    dc = new Model.DocCompra();
                    dc.id = objListCab.Valor("id");
                    dc.NumDocExterno = objListCab.Valor("NumDocExterno");
                    dc.Entidade = objListCab.Valor("Entidade");
                    dc.NumDoc = objListCab.Valor("NumDoc");
                    dc.Data = objListCab.Valor("DataDoc");
                    dc.TotalMerc = objListCab.Valor("TotalMerc");
                    dc.Serie = objListCab.Valor("Serie");
                    objListLin = PriEngine.Engine.Consulta("SELECT idCabecCompras, Artigo, Descricao, Quantidade, Unidade, PrecUnit, Desconto1, TotalILiquido, PrecoLiquido, Armazem, Lote from LinhasCompras where IdCabecCompras='" + dc.id + "' order By NumLinha");
                    listlindc = new List<Model.LinhaDocCompra>();

                    while (!objListLin.NoFim())
                    {
                        lindc = new Model.LinhaDocCompra();
                        lindc.IdCabecDoc = objListLin.Valor("idCabecCompras");
                        lindc.CodArtigo = objListLin.Valor("Artigo");
                        lindc.DescArtigo = objListLin.Valor("Descricao");
                        lindc.Quantidade = objListLin.Valor("Quantidade");
                        lindc.Unidade = objListLin.Valor("Unidade");
                        lindc.Desconto = objListLin.Valor("Desconto1");
                        lindc.PrecoUnitario = objListLin.Valor("PrecUnit");
                        lindc.TotalILiquido = objListLin.Valor("TotalILiquido");
                        lindc.TotalLiquido = objListLin.Valor("PrecoLiquido");
                        lindc.Armazem = objListLin.Valor("Armazem");
                        lindc.Lote = objListLin.Valor("Lote");

                        listlindc.Add(lindc);
                        objListLin.Seguinte();
                    }

                    dc.LinhasDoc = listlindc;
                    
                    listdc.Add(dc);
                    objListCab.Seguinte();
                }
            }
            return listdc;
        }



        public static Model.RespostaErro VGR_New(Model.DocCompra dc)
        {
            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();
            

            GcpBEDocumentoCompra myGR = new GcpBEDocumentoCompra();
            GcpBELinhaDocumentoCompra myLin = new GcpBELinhaDocumentoCompra();
            GcpBELinhasDocumentoCompra myLinhas = new GcpBELinhasDocumentoCompra();

            PreencheRelacaoCompras rl = new PreencheRelacaoCompras();
            List<Model.LinhaDocCompra> lstlindv = new List<Model.LinhaDocCompra>();

            try
            {
                if (PriEngine.InitializeCompany("EMP1", "", "") == true)
                {
                    // Atribui valores ao cabecalho do doc
                    //myEnc.set_DataDoc(dv.Data);
                    myGR.set_Entidade(dc.Entidade);
                    myGR.set_NumDocExterno(dc.NumDocExterno);
                    myGR.set_Serie(dc.Serie);
                    myGR.set_Tipodoc("VGR");
                    myGR.set_TipoEntidade("F");
                    // Linhas do documento para a lista de linhas
                    lstlindv = dc.LinhasDoc;
                    PriEngine.Engine.Comercial.Compras.PreencheDadosRelacionados(myGR, rl);
                    foreach (Model.LinhaDocCompra lin in lstlindv)
                    {
                        PriEngine.Engine.Comercial.Compras.AdicionaLinha(myGR, lin.CodArtigo, lin.Quantidade, lin.Armazem, "", lin.PrecoUnitario, lin.Desconto);
                    }


                    PriEngine.Engine.IniciaTransaccao();
                    PriEngine.Engine.Comercial.Compras.Actualiza(myGR, "Teste");
                    PriEngine.Engine.TerminaTransaccao();
                    erro.Erro = 0;
                    erro.Descricao = "Sucesso";
                    return erro;
                }
                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir empresa";
                    return erro;

                }

            }
            catch (Exception ex)
            {
                PriEngine.Engine.DesfazTransaccao();
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }
        }


        public static Model.DocCompra GetDocumentoCompra(string codEmpresa, string tipoDeDocumento, string numDoc)
        {

            ErpBS objMotor = new ErpBS();

            Model.DocCompra dc = new Model.DocCompra();
            Model.LinhaDocCompra lindc = new Model.LinhaDocCompra();
            List<Model.LinhaDocCompra> listlindc = new List<Model.LinhaDocCompra>();
            StdBELista objListCab;
            StdBELista objListLin;

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {

                objListCab = PriEngine.Engine.Consulta("SELECT id, NumDocExterno, Entidade, DataDoc, NumDoc, TotalMerc, Serie From CabecCompras where TipoDoc='" + tipoDeDocumento + "' and NumDoc='" + numDoc + "'");
                    dc.id = objListCab.Valor("id");
                    dc.NumDocExterno = objListCab.Valor("NumDocExterno");
                    dc.Entidade = objListCab.Valor("Entidade");
                    dc.NumDoc = objListCab.Valor("NumDoc");
                    dc.Data = objListCab.Valor("DataDoc");
                    dc.TotalMerc = objListCab.Valor("TotalMerc");
                    dc.Serie = objListCab.Valor("Serie");
                    objListLin = PriEngine.Engine.Consulta("SELECT idCabecCompras, Artigo, Descricao, Quantidade, Unidade, PrecUnit, Desconto1, TotalILiquido, PrecoLiquido, Armazem, Lote from LinhasCompras where IdCabecCompras='" + dc.id + "' order By NumLinha");
                    listlindc = new List<Model.LinhaDocCompra>();

                    while (!objListLin.NoFim())
                    {
                        lindc = new Model.LinhaDocCompra();
                        lindc.IdCabecDoc = objListLin.Valor("idCabecCompras");
                        lindc.CodArtigo = objListLin.Valor("Artigo");
                        lindc.DescArtigo = objListLin.Valor("Descricao");
                        lindc.Quantidade = objListLin.Valor("Quantidade");
                        lindc.Unidade = objListLin.Valor("Unidade");
                        lindc.Desconto = objListLin.Valor("Desconto1");
                        lindc.PrecoUnitario = objListLin.Valor("PrecUnit");
                        lindc.TotalILiquido = objListLin.Valor("TotalILiquido");
                        lindc.TotalLiquido = objListLin.Valor("PrecoLiquido");
                        lindc.Armazem = objListLin.Valor("Armazem");
                        lindc.Lote = objListLin.Valor("Lote");
                        listlindc.Add(lindc);
                        objListLin.Seguinte();
                    }

                    dc.LinhasDoc = listlindc;
                    return dc;                
            }
            return null;
        } 

        #endregion DocCompra;

        // ------ Documentos de venda ----------------------

        #region DocVenda

        public static Model.RespostaErro Encomendas_New(Model.DocVenda dv)
        {
            Lib_Primavera.Model.RespostaErro erro = new Model.RespostaErro();
            GcpBEDocumentoVenda myEnc = new GcpBEDocumentoVenda();
             
            GcpBELinhaDocumentoVenda myLin = new GcpBELinhaDocumentoVenda();

            GcpBELinhasDocumentoVenda myLinhas = new GcpBELinhasDocumentoVenda();
             
            PreencheRelacaoVendas rl = new PreencheRelacaoVendas();
            List<Model.LinhaDocVenda> lstlindv = new List<Model.LinhaDocVenda>();
            
            try
            {
                if (PriEngine.InitializeCompany("EMP1", "", "") == true)
                {
                    // Atribui valores ao cabecalho do doc
                    //myEnc.set_DataDoc(dv.Data);
                    myEnc.set_Entidade(dv.Entidade);
                    myEnc.set_Serie(dv.Serie);
                    myEnc.set_Tipodoc("ECL");
                    myEnc.set_TipoEntidade("C");
                    // Linhas do documento para a lista de linhas
                    lstlindv = dv.LinhasDoc;
                    PriEngine.Engine.Comercial.Vendas.PreencheDadosRelacionados(myEnc, rl);
                    foreach (Model.LinhaDocVenda lin in lstlindv)
                    {
                        PriEngine.Engine.Comercial.Vendas.AdicionaLinha(myEnc, lin.CodArtigo, lin.Quantidade, "", "", lin.PrecoUnitario, lin.Desconto);
                    }


                   // PriEngine.Engine.Comercial.Compras.TransformaDocumento(

                    PriEngine.Engine.IniciaTransaccao();
                    PriEngine.Engine.Comercial.Vendas.Actualiza(myEnc, "Teste");
                    PriEngine.Engine.TerminaTransaccao();
                    erro.Erro = 0;
                    erro.Descricao = "Sucesso";
                    return erro;
                }
                else
                {
                    erro.Erro = 1;
                    erro.Descricao = "Erro ao abrir empresa";
                    return erro;

                }

            }
            catch (Exception ex)
            {
                PriEngine.Engine.DesfazTransaccao();
                erro.Erro = 1;
                erro.Descricao = ex.Message;
                return erro;
            }
        }


        public static List<Model.DocVenda> ListaDocumentosVenda(string codEmpresa, string tipoDeDocumento)
        {
            ErpBS objMotor = new ErpBS();
            
            StdBELista objListCab;
            StdBELista objListLin;
            Model.DocVenda dv = new Model.DocVenda();
            List<Model.DocVenda> listdv = new List<Model.DocVenda>();
            Model.LinhaDocVenda lindv = new Model.LinhaDocVenda();
            List<Model.LinhaDocVenda> listlindv = new
            List<Model.LinhaDocVenda>();

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {
                objListCab = PriEngine.Engine.Consulta("SELECT id, Entidade, Data, NumDoc, TotalMerc, Serie From CabecDoc where TipoDoc='" + tipoDeDocumento + "'");
                while (!objListCab.NoFim())
                {
                    dv = new Model.DocVenda();
                    dv.id = objListCab.Valor("id");
                    dv.Entidade = objListCab.Valor("Entidade");
                    dv.NumDoc = objListCab.Valor("NumDoc");
                    dv.Data = objListCab.Valor("Data");
                    dv.TotalMerc = objListCab.Valor("TotalMerc");
                    dv.Serie = objListCab.Valor("Serie");
                    objListLin = PriEngine.Engine.Consulta("SELECT idCabecDoc, Artigo, Descricao, Quantidade, Unidade, PrecUnit, Desconto1, TotalILiquido, PrecoLiquido from LinhasDoc where IdCabecDoc='" + dv.id + "' order By NumLinha");
                    listlindv = new List<Model.LinhaDocVenda>();

                    while (!objListLin.NoFim())
                    {
                        lindv = new Model.LinhaDocVenda();
                        lindv.IdCabecDoc = objListLin.Valor("idCabecDoc");
                        lindv.CodArtigo = objListLin.Valor("Artigo");
                        lindv.DescArtigo = objListLin.Valor("Descricao");
                        lindv.Quantidade = objListLin.Valor("Quantidade");
                        lindv.Unidade = objListLin.Valor("Unidade");
                        lindv.Desconto = objListLin.Valor("Desconto1");
                        lindv.PrecoUnitario = objListLin.Valor("PrecUnit");
                        lindv.TotalILiquido = objListLin.Valor("TotalILiquido");
                        lindv.TotalLiquido = objListLin.Valor("PrecoLiquido");

                        listlindv.Add(lindv);
                        objListLin.Seguinte();
                    }

                    dv.LinhasDoc = listlindv;
                    listdv.Add(dv);
                    objListCab.Seguinte();
                }
            }
            return listdv;
        }


        public static Model.DocVenda GetDocumentoVenda(string codEmpresa, string tipoDeDocumento, string numDoc)
        {
            ErpBS objMotor = new ErpBS();
             
            StdBELista objListCab;
            StdBELista objListLin;
            Model.DocVenda dv = new Model.DocVenda();
            Model.LinhaDocVenda lindv = new Model.LinhaDocVenda();
            List<Model.LinhaDocVenda> listlindv = new List<Model.LinhaDocVenda>();

            if (PriEngine.InitializeCompany(codEmpresa, "", "") == true)
            {
                 
                string st = "SELECT id, Entidade, Data, NumDoc, TotalMerc, Serie From CabecDoc where TipoDoc='"+ tipoDeDocumento +"' and NumDoc='" + numDoc + "'";
                objListCab = PriEngine.Engine.Consulta(st);
                dv = new Model.DocVenda();
                dv.id = objListCab.Valor("id");
                dv.Entidade = objListCab.Valor("Entidade");
                dv.NumDoc = objListCab.Valor("NumDoc");
                dv.Data = objListCab.Valor("Data");
                dv.TotalMerc = objListCab.Valor("TotalMerc");
                dv.Serie = objListCab.Valor("Serie");
                objListLin = PriEngine.Engine.Consulta("SELECT idCabecDoc, Artigo, Descricao, Quantidade, Unidade, PrecUnit, Desconto1, TotalILiquido, PrecoLiquido from LinhasDoc where IdCabecDoc='" + dv.id + "' order By NumLinha");
                listlindv = new List<Model.LinhaDocVenda>();

                while (!objListLin.NoFim())
                {
                    lindv = new Model.LinhaDocVenda();
                    lindv.IdCabecDoc = objListLin.Valor("idCabecDoc");
                    lindv.CodArtigo = objListLin.Valor("Artigo");
                    lindv.DescArtigo = objListLin.Valor("Descricao");
                    lindv.Quantidade = objListLin.Valor("Quantidade");
                    lindv.Unidade = objListLin.Valor("Unidade");
                    lindv.Desconto = objListLin.Valor("Desconto1");
                    lindv.PrecoUnitario = objListLin.Valor("PrecUnit");
                    lindv.TotalILiquido = objListLin.Valor("TotalILiquido");
                    lindv.TotalLiquido = objListLin.Valor("PrecoLiquido");
                    listlindv.Add(lindv);
                    objListLin.Seguinte();
                }

                dv.LinhasDoc = listlindv;
                return dv;
            }
            return null;
        }

        #endregion DocVend;

    }
}