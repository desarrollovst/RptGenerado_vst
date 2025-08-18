using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Oracle.DataAccess.Client;

/// <summary>
/// Summary description for Parameters
/// </summary>
public class Parameters
{
    public Parameters()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    #region ORACLE_PARAMETERS

    public OracleParameter[] ParamsAgenda(string empresa, DateTime fecha, string usuario, string coord, string asesor)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("vCDGCO", OracleDbType.Varchar2),
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2),
                                               new OracleParameter("vINICIO", OracleDbType.Date),
                                               new OracleParameter("vCDGPE", OracleDbType.Varchar2)};


        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = coord;
        OracleParameters[2].Value = asesor;
        OracleParameters[3].Value = fecha;
        OracleParameters[4].Value = usuario;

        return OracleParameters;
    }

    public OracleParameter[] ParamsBandas(string empresa, DateTime fecha, int tipoCart, string usuario, string region,
                                           string coord, string asesor)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("dFECHAPAGOS", OracleDbType.Date),
                                               new OracleParameter("vREGION", OracleDbType.Varchar2),
                                               new OracleParameter("vSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2), 
                                               new OracleParameter("vUSUARIO", OracleDbType.Varchar2),
                                               new OracleParameter("vTIPOCAR", OracleDbType.Int32)}; 

        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = fecha;
        OracleParameters[2].Value = region;
        OracleParameters[3].Value = coord;
        OracleParameters[4].Value = asesor;
        OracleParameters[5].Value = usuario;
        OracleParameters[6].Value = tipoCart;

        return OracleParameters;
    }

    public OracleParameter[] ParamsCierreAcred(string empresa, string sucursal, string fecha, string usuario)
    {
        OracleParameter[] OracleParameters = {   new OracleParameter("prmCDGEM", OracleDbType.Varchar2), 
                                                 new OracleParameter("prmCDGCO", OracleDbType.Varchar2),
                                                 new OracleParameter("prmFECHA", OracleDbType.Varchar2),
                                                 new OracleParameter("prmCDGPE", OracleDbType.Varchar2)};

        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = sucursal;
        OracleParameters[2].Value = fecha;
        OracleParameters[3].Value = usuario;

        return OracleParameters;
    }

    public OracleParameter[] ParamsContPagos(string empresa, string fecha, string region, string sucursal, string asesor, string usuario)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("prmCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("prmFECHA", OracleDbType.Varchar2),
                                               new OracleParameter("prmREGION", OracleDbType.Varchar2),
                                               new OracleParameter("prmSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("prmASESOR", OracleDbType.Varchar2),
                                               new OracleParameter("prmCDGPE", OracleDbType.Varchar2),
                                               new OracleParameter("prmMENSAJE", OracleDbType.Varchar2, 200)};

        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = fecha;
        OracleParameters[2].Value = region;
        OracleParameters[3].Value = sucursal;
        OracleParameters[4].Value = asesor;
        OracleParameters[5].Value = usuario;
        OracleParameters[6].Direction = ParameterDirection.Output;

        return OracleParameters;
    }

    public OracleParameter[] ParamsMora(string empresa, DateTime fecha, int nivel, int tipoCart, string usuario, string region,
                                           string coord, string asesor)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("dFECHAHASTA", OracleDbType.Date),
                                               new OracleParameter("vNIVEL", OracleDbType.Int32),
                                               new OracleParameter("vUSUARIO", OracleDbType.Varchar2),
                                               new OracleParameter("vREGION", OracleDbType.Varchar2),
                                               new OracleParameter("vSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2), 
                                               new OracleParameter("vTIPOCAR", OracleDbType.Int32) };

        
        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = fecha;
        OracleParameters[2].Value = nivel;
        OracleParameters[3].Value = usuario;
        OracleParameters[4].Value = region;
        OracleParameters[5].Value = coord;
        OracleParameters[6].Value = asesor;
        OracleParameters[7].Value = tipoCart;

        
        return OracleParameters;
    }

    public OracleParameter[] ParamsMora(string empresa, DateTime fecha, int nivel, int cartVig, int cartVenc, int cartRest, int cartCast, string usuario,
                                           string region, string sucursal, string coord, string asesor)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("dFECHAHASTA", OracleDbType.Date),
                                               new OracleParameter("vNIVEL", OracleDbType.Int32),
                                               new OracleParameter("vUSUARIO", OracleDbType.Varchar2),
                                               new OracleParameter("vREGION", OracleDbType.Varchar2),
                                               new OracleParameter("vSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("vCOORD", OracleDbType.Varchar2), 
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2),
                                               new OracleParameter("vCARTVIG", OracleDbType.Int32),
                                               new OracleParameter("vCARTVENC", OracleDbType.Int32),
                                               new OracleParameter("vCARTREST", OracleDbType.Int32),
                                               new OracleParameter("vCARTCAST", OracleDbType.Int32)};


        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = fecha;
        OracleParameters[2].Value = nivel;
        OracleParameters[3].Value = usuario;
        OracleParameters[4].Value = region;
        OracleParameters[5].Value = sucursal;
        OracleParameters[6].Value = coord;
        OracleParameters[7].Value = asesor;
        OracleParameters[8].Value = cartVig;
        OracleParameters[9].Value = cartVenc;
        OracleParameters[10].Value = cartRest;
        OracleParameters[11].Value = cartCast;

        return OracleParameters;
    }

    public OracleParameter[] ParamsMora(string empresa, DateTime fecha, int nivel, int cartVig, int cartVenc, int cartRest, int cartCast, string usuario, 
                                           string region, string sucursal, string coord, string asesor, string tipoProd, string nivelMora)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("dFECHAHASTA", OracleDbType.Date),
                                               new OracleParameter("vNIVEL", OracleDbType.Int32),
                                               new OracleParameter("vUSUARIO", OracleDbType.Varchar2),
                                               new OracleParameter("vREGION", OracleDbType.Varchar2),
                                               new OracleParameter("vSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("vCOORD", OracleDbType.Varchar2), 
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2),
                                               new OracleParameter("vCARTVIG", OracleDbType.Int32),
                                               new OracleParameter("vCARTVENC", OracleDbType.Int32),
                                               new OracleParameter("vCARTREST", OracleDbType.Int32),
                                               new OracleParameter("vCARTCAST", OracleDbType.Int32),
                                               new OracleParameter("vTIPOPROD", OracleDbType.Varchar2), 
                                               new OracleParameter("vNIVMORA", OracleDbType.Varchar2) };


        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = fecha;
        OracleParameters[2].Value = nivel;
        OracleParameters[3].Value = usuario;
        OracleParameters[4].Value = region;
        OracleParameters[5].Value = sucursal;
        OracleParameters[6].Value = coord;
        OracleParameters[7].Value = asesor;
        OracleParameters[8].Value = cartVig;
        OracleParameters[9].Value = cartVenc;
        OracleParameters[10].Value = cartRest;
        OracleParameters[11].Value = cartCast;
        OracleParameters[12].Value = tipoProd;
        OracleParameters[13].Value = nivelMora;

        return OracleParameters;
    }

    public OracleParameter[] ParamsPagosTeoricos(string empresa, DateTime fecha, string usuario, string region,
                                           string coord, string asesor)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("dFECHAHASTA", OracleDbType.Date),
                                               new OracleParameter("vUSUARIO", OracleDbType.Varchar2),
                                               new OracleParameter("vREGION", OracleDbType.Varchar2),
                                               new OracleParameter("vSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2) };

        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = fecha;
        OracleParameters[2].Value = usuario;
        OracleParameters[3].Value = region;
        OracleParameters[4].Value = coord;
        OracleParameters[5].Value = asesor;

        return OracleParameters;
    }

    public OracleParameter[] ParamsPagosTeoricosCobranza(string empresa, DateTime fecha, int tipoCart, string usuario,
                                                         string coord, string asesor, string supervisor)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("vSUCURSAL", OracleDbType.Varchar2),
                                               new OracleParameter("vASESOR", OracleDbType.Varchar2),
                                               new OracleParameter("dFECHAHASTA", OracleDbType.Date),
                                               new OracleParameter("vUSUARIO", OracleDbType.Varchar2),                                        
                                               new OracleParameter("vSUPERVISOR", OracleDbType.Varchar2),
                                               new OracleParameter("vTIPOCART", OracleDbType.Varchar2) };
      
        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = coord;
        OracleParameters[2].Value = asesor;
        OracleParameters[3].Value = fecha;
        OracleParameters[4].Value = usuario;
        OracleParameters[5].Value = supervisor;
        OracleParameters[6].Value = tipoCart;

        return OracleParameters;
    }

    public OracleParameter[] ParamsPlanPagos(string empresa, string grupo, string ciclo, string clns, string usuario)
    {
        OracleParameter[] OracleParameters = { new OracleParameter("vCDGPE", OracleDbType.Varchar2),
                                               new OracleParameter("vCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("vCDGCLNS", OracleDbType.Varchar2),
                                               new OracleParameter("vCICLO", OracleDbType.Varchar2),
                                               new OracleParameter("vCLNS", OracleDbType.Varchar2) };

        OracleParameters[0].Value = usuario;
        OracleParameters[1].Value = empresa;
        OracleParameters[2].Value = grupo;
        OracleParameters[3].Value = ciclo;
        OracleParameters[4].Value = clns;
      
        return OracleParameters;
    }

    public OracleParameter[] ParamsRepReferencias(string empresa, string asesor, string usuario)
    {
        OracleParameter[] OracleParameters = { 
                                               new OracleParameter("pCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("pCDGOCPE", OracleDbType.Varchar2),
                                               new OracleParameter("pCDGPE", OracleDbType.Varchar2)};
        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = asesor;
        OracleParameters[2].Value = usuario;

        return OracleParameters;
    }

    public OracleParameter[] ParamsRepSol(string empresa, string grupo, string ciclo, string usuario)
    {
        OracleParameter[] OracleParameters = { 
                                               new OracleParameter("pCDGEM", OracleDbType.Varchar2),
                                               new OracleParameter("pCDGNS", OracleDbType.Varchar2),
                                               new OracleParameter("pCICLO", OracleDbType.Varchar2),
                                               new OracleParameter("pCDGPE", OracleDbType.Varchar2)};  
        OracleParameters[0].Value = empresa;
        OracleParameters[1].Value = grupo;
        OracleParameters[2].Value = ciclo;
        OracleParameters[3].Value = usuario;
        
        
        return OracleParameters;
    }
    
    #endregion
}
