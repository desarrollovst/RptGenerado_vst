using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Globalization;

using System.IO;
using Microsoft.VisualBasic;
using BarcodeLib;

// Agregar los espacios de nombres al principio. 
using CrystalDecisions.CrystalReports.Engine; 
using CrystalDecisions.Shared;

public partial class _Default : System.Web.UI.Page 
{
    public string cdgEmpresa;
    public string tipoDoc;
    ReportDocument grcReporte = new ReportDocument();

    DBService db = new DBService();
    Parameters oP = new Parameters();

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            cdgEmpresa = ConfigurationManager.AppSettings.Get("Empresa");
            string id = Request.QueryString.Get("id");
            string fecha = Request.QueryString.Get("fecha");
            string fechaFin = Request.QueryString.Get("fechaFin");
            string nivel = Request.QueryString.Get("nivel");
            string tipoCart = Request.QueryString.Get("tipoCart");
            string tipoProd = Request.QueryString.Get("tipoProd");
            string cartVig = Request.QueryString.Get("cartVig");
            string cartVenc = Request.QueryString.Get("cartVenc");
            string cartRest = Request.QueryString.Get("cartRest");
            string cartCast = Request.QueryString.Get("cartCast");
            string usuario = Request.QueryString.Get("usuario");
            string nomUsuario = Request.QueryString.Get("nomUsuario");
            string puesto = Request.QueryString.Get("puesto");
            string region = Request.QueryString.Get("region");
            string sucursal = Request.QueryString.Get("sucursal");
            string coord = Request.QueryString.Get("coord");
            string asesor = Request.QueryString.Get("asesor");
            string supervisor = Request.QueryString.Get("supervisor");
            string grupo = Request.QueryString.Get("grupo");
            string ciclo = Request.QueryString.Get("ciclo");
            string acred = Request.QueryString.Get("acred");
            string chqIni = Request.QueryString.Get("chqIni");
            string chqFin = Request.QueryString.Get("chqFin");
            string cuenta = Request.QueryString.Get("cuenta");
            string tipo = Request.QueryString.Get("tipo");
            string clns = Request.QueryString.Get("clns");
            string codigo = Request.QueryString.Get("codigo");
            string mes = Request.QueryString.Get("mes");
            string anio = Request.QueryString.Get("anio");
            string formato = Request.QueryString.Get("formato");
            string nivelMora = Request.QueryString.Get("nivelMora");
            string titulo = Request.QueryString.Get("titulo");
            string sic = Request.QueryString.Get("sic");
            tipoDoc = Request.QueryString.Get("tipoDoc");

            string closeS =  Request.QueryString.Get("close");

            string resultado = string.Empty;
            fecha = (fecha != null ? Convert.ToDateTime(fecha).ToString("dd/MM/yyyy"): fecha);
            fechaFin = (fechaFin != null ? Convert.ToDateTime(fechaFin).ToString("dd/MM/yyyy"): fechaFin);
            tipoDoc = (tipoDoc != null ? tipoDoc : "PDF");
            clns = (clns != null ? clns : "G");
            //id = "35";
            //usuario = "ADMIN";
            //nomUsuario = "ADMIN";
            //fecha = "27/02/2024";
            //fechaFin = "07/11/2014";
            //region = "000";
            //sucursal = "005";
            //coord = "003";
            //acred = "014371";
            //asesor = "000000";
            //supervisor = "000000";
            //nivel = "3";
            //tipoCart = "0";
            //tipoProd = "0";
            //cartVig = "1";
            //cartVenc = "1";
            //cartRest = "1";
            //cartCast = "0";
            //grupo = "017088";
            //ciclo = "01";
            //cuenta = "14";
            //chqIni = "0000661";
            //chqFin = "0000665";
            //tipo = "C";
            //formato = "CertificadoVST.rpt";
            //puesto = "D";
            //mes = "07";
            //anio = "2018";
            //nivelMora = "1";
            //clns = "G";
            //tipoDoc = "PDF";
            sic = "CC";
            if (id != null)
            {
                switch (id)
                {
                    case "1":
                        getRepSitCartera(id, fecha, nivel, cartVig, cartVenc, cartRest, cartCast, usuario, nomUsuario, region, sucursal, coord, asesor, tipoProd, nivelMora, titulo);
                        break;
                    case "8":
                        getSolicitudCred(grupo, ciclo, usuario, nomUsuario);
                        break;
                    case "19":
                        getConsultaRepCredito(tipo, acred, usuario, nomUsuario, sic, fecha);
                        //getConsultaBuro(acred, usuario, nomUsuario);
                        break;
                    case "21":
                        getImpresionCheques(grupo, ciclo, cuenta, coord, chqIni, chqFin);
                        break;
                    case "26":
                        getReimpCheques(tipo, grupo, ciclo, clns, cuenta, chqIni);
                        break;
                    case "27":
                        getEstadoCuentaGrupal(grupo, ciclo);
                        break;
                    case "28":
                        getPagare(grupo, ciclo, usuario);
                        break;
                    case "29":
                        getContrato(grupo, ciclo, usuario);
                        break;
                    case "35":
                        getCartaPago(grupo, acred, ciclo, usuario);
                        break;
                    case "36":
                        getCartaGarantia(grupo, ciclo, usuario);
                        break;
                    case "41":
                    case "42":
                        getDocMicroseguro(id, grupo, ciclo, clns, usuario, tipo, formato);
                        break;
                    case "43":
                        getPagareGrupal(grupo, ciclo, usuario);
                        break;
                }
            }
            else
            {
                LlenaRptError("", "Debe especificar correctamente el tipo de documento");
            }
        }
    }

    private void getCartaGarantia(string grupo, string ciclo, string usuario)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        string queryGrupo = string.Empty;
        int iRes;
        try
        {
            DataSet dsCarta = new DataSet();

            if (grupo != null)
                queryGrupo = "AND CODIGO_GPO = '" + grupo + "' ";

            query = "SELECT IM.CODIGO_GPO " +
                    ",IM.NOMBRE_GPO " +
                    ",IM.CODIGO_CTE " +
                    ",IM.NOMBRE_CTE " +
                    ",IM.NOM_SUCURSAL " +
                    ",IM.PRESIDENTE " +
                    ",IM.TESORERO " +
                    ",IM.SECRETARIO " +
                    ",IM.SUPERVISOR " +
                    ",IM.PAGO_PARCIAL_LETRA " +
                    ",IM.TOTAL_A_PAGAR_LETRA " +
                    ",IM.CUENTA_BANCARIA " +
                    ",IM.NOM_GERENTE_SUC " +
                    ",IM.NOM_ASESOR REF_TELECOMM " +
                    ",IM.DIR_SUCURSAL REF_BANCOMER " +
                    ",IM.REF_BANSEFI " +
                    ",IM.TIPO_PROD " + 
                    ",CO.CODIGO CDGCO " +
                    ",IM.NOM_PROD " +
                    ",IM.NOM_INTEG_DNI " +
                    ",IM.REF_OPENPAY " +
                    ",IM.REF_PAYCASH " +
                    "FROM IMPCONT IM, NS, CO " +
                    "WHERE IM.CODIGO_EMP = '" + empresa + "' " +
                    queryGrupo +
                    "AND IM.TIPO_DOC = 'GARANTIA' " +
                    "AND IM.CDGPE = '" + usuario + "' " +
                    "AND NS.CDGEM = IM.CODIGO_EMP " +
                    "AND NS.CODIGO = IM.CODIGO_GPO " +
                    "AND CO.CDGEM = NS.CDGEM " +
                    "AND CO.CODIGO = NS.CDGCO";

            iRes = db.ExecuteDS(ref dsCarta, query, CommandType.Text);

            if (dsCarta.Tables[0].Rows.Count > 0)
                LlenaCartaGarantia(dsCarta, grupo);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getCartaPago(string grupo, string acred, string ciclo, string usuario)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        string queryGrupo = string.Empty;
        string queryAcred = string.Empty;
        int iRes;
        try
        {
            DataSet dsCarta = new DataSet();
            DataSet dsCartaDet = new DataSet();

            if (grupo != null)
                queryGrupo = "AND IM.CODIGO_GPO = '" + grupo + "' ";
            else if (grupo == null)
            {
                queryGrupo = "AND CODIGO_GPO IS NULL ";
                queryAcred = "AND CODIGO_CTE = '" + acred + "' ";
            }

            query = "SELECT IM.CODIGO_GPO " +
                    ",IM.NOMBRE_GPO " +
                    ",IM.NO_CICLO CICLO " +
                    ",IM.NUM_INT " +
                    ",IM.TASA_CTE TASA " +
                    ",IM.PLAZO " +
                    ",TO_DATE(IM.FECHA_INICIO,'DD/MM/YYYY') INICIO " +
                    ",TO_DATE(IM.FECHA_FIN_LETRA,'DD/MM/YYYY') FIN " +
                    ",IM.TIPO_PROD " +
                    ",IM.NUM_INT " +
                    ",IM.CANTENTRE " +
                    ",IM.TOTAL_SIN_IVA TOTAL " +
                    ",IM.PARC_SIN_IVA PARCIALIDAD " +
                    ",(SELECT REF_PAYCASH FROM PAYCASH_REF WHERE CDGEM = IM.CODIGO_EMP AND CDGCLNS = IM.CODIGO_GPO AND CDGTPC = IM.TIPO_PROD AND TIPO = 1) REF_PAYCASH " +
                    ",(SELECT REF_PAYCASH FROM PAYCASH_REF WHERE CDGEM = IM.CODIGO_EMP AND CDGCLNS = IM.CODIGO_GPO AND CDGTPC = IM.TIPO_PROD AND TIPO = 0) REF_PAYCASH_GL " +
                    "FROM IMPCONT IM " +
                    "WHERE IM.CODIGO_EMP = '" + empresa + "' " +
                    queryGrupo +
                    queryAcred +
                    "AND IM.CDGPE = '" + usuario + "' " +
                    "AND IM.TIPO_DOC = 'PAGO'";

            iRes = db.ExecuteDS(ref dsCarta, query, CommandType.Text);

            query = "SELECT NOMBREC(PRC.CDGEM,PRC.CDGCL,'I','A',NULL,NULL,NULL,NULL) NOMBRE " +
                    ",ROUND(IM.CANTENTRE * (PRC.CANTENTRE / PRN.CANTENTRE),2) CANTENTRE " +
                    ",ROUND(IM.TOTAL_SIN_IVA * (PRC.CANTENTRE / PRN.CANTENTRE),2) TOTAL " +
                    ",ROUND(IM.PARC_SIN_IVA * (PRC.CANTENTRE / PRN.CANTENTRE),2) PARCIALIDAD " +
                    "FROM IMPCONT IM, PRN, PRC " +
                    "WHERE IM.CODIGO_EMP = '" + empresa + "' " +
                    queryGrupo +
                    queryAcred +
                    "AND IM.CDGPE = '" + usuario + "' " +
                    "AND IM.TIPO_DOC = 'PAGO' " +
                    "AND PRN.CDGEM = IM.CODIGO_EMP " +
                    "AND PRN.CDGNS = IM.CODIGO_GPO " +
                    "AND PRN.CICLO = IM.NO_CICLO " +
                    "AND PRN.INICIO = TO_DATE(IM.FECHA_INICIO,'DD/MM/YYYY') " +
                    "AND PRC.CDGEM = PRN.CDGEM " +
                    "AND PRC.CDGCLNS = PRN.CDGNS " +
                    "AND PRC.CICLO = PRN.CICLO " +
                    "AND PRC.SITUACION = 'E' " +
                    "ORDER BY NOMBRE";

            iRes = db.ExecuteDS(ref dsCartaDet, query, CommandType.Text);

            if (dsCarta.Tables[0].Rows.Count > 0 && dsCartaDet.Tables[0].Rows.Count > 0)
                LlenaCartaPago(dsCarta, dsCartaDet, grupo);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getDocMicroseguro(string id, string grupo, string ciclo, string clns, string usuario, string tipoProd, string formato)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        int iRes;
        try
        {
            DataSet dsCons = new DataSet();

            query = " SELECT M.CDGCL CODIGO_CTE " +
                    ",NOMBREC(CL.CDGEM,CL.CODIGO,'I','N',NULL,NULL,NULL,NULL) NOMBRE_CTE " +
                    ",CL.NACIMIENTO " +
                    ",TO_CHAR(M.INICIO,'DD/MM/YYYY') FECINI " +
                    ",TO_CHAR(M.FIN,'DD/MM/YYYY') FECFIN " +
                    ",MF.CDGPMS " +
                    "FROM IMPCONT IM " +
                    "JOIN (SELECT T1.CDGEM, T1.CDGCL, T1.CDGASE, T1.INICIO, T1.FIN, SUBSTR(XMLAGG(XMLELEMENT(E,T1.CDGPMS||',')).EXTRACT('//text()'), 1, LENGTH (XMLAGG(XMLELEMENT(E,T1.CDGPMS||',')).EXTRACT('//text()')) - 1) CDGPMS " +
                          "FROM (SELECT M.CDGEM, M.CDGASE, M.CDGCL, M.CDGPMS, M.INICIO, M.FIN " +
                                "FROM MICROSEGURO M " +
                                "WHERE M.CDGEM = '" + empresa + "' " +
                                "AND M.CDGCLNS = '" + grupo + "' " +
                                "AND M.CICLO = '" + ciclo + "' " +
                                "AND M.CLNS = '" + clns + "' " +
                                "AND M.ESTATUS IN ('R','V') " +
                                "GROUP BY M.CDGEM, M.CDGASE, M.CDGCL, M.CDGPMS, M.INICIO, M.FIN " +
                                "ORDER BY M.CDGEM, M.CDGASE, M.CDGCL, M.CDGPMS, M.INICIO, M.FIN ) T1 " +
                          "GROUP BY T1.CDGEM, T1.CDGCL, T1.CDGASE, T1.INICIO, T1.FIN ) M ON IM.CODIGO_EMP = M.CDGEM " +
                    "JOIN CL ON M.CDGEM = CL.CDGEM AND M.CDGCL = CL.CODIGO " +
                    "JOIN MICROSEGURO_FORMATOS MF ON M.CDGEM = MF.CDGEM AND M.CDGASE = MF.CDGASE AND M.CDGPMS = MF.CDGPMS " +
                    "WHERE IM.CODIGO_EMP = '" + empresa + "' " +
                    "AND IM.CODIGO_GPO = '" + grupo + "' " +
                    "AND IM.NO_CICLO = '" + ciclo + "' " +
                    "AND IM.CDGPE = '" + usuario + "' " +
                    "AND IM.TIPO_DOC = 'CONTRATO' " +
                    "AND MF.CDGPMS = '" + tipoProd + "' ";

            iRes = db.ExecuteDS(ref dsCons, query, CommandType.Text);

            if (dsCons.Tables[0].Rows.Count > 0)
                LlenaDocMicroseguro(id, dsCons, tipoProd, formato);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getConsultaBuro(string acred, string usuario, string nomUsuario)
    {
        string empresa = cdgEmpresa;
        int iRes;
        try
        {
            DataSet dsInfo = new DataSet();
            DataSet dsNom = new DataSet();
            DataSet dsDir = new DataSet();
            DataSet dsCta = new DataSet();
            DataSet dsRes = new DataSet();
            DataSet dsCons = new DataSet();

            string query = "SELECT TO_CHAR(CB.FCONSULTA,'DD/MM/YYYY') FECHACONS, " +
                           "TO_CHAR(CB.FVIGENCIA,'DD/MM/YYYY') AS FECHAVIG, " +
                           "TO_CHAR(sysdate,'DD/MM/YYYY') AS FECHAIMP, " + 
                           "TO_CHAR(sysdate,'DD/MM/YYYY') AS FECHAIMP, " +
                           "TO_CHAR(sysdate,'HH24:MI:SS') AS HORAIMP " +
                           "FROM CONSULTA_BURO CB " +
                           "WHERE CB.CDGEM = '" + empresa + "' " +
                           "AND CB.CDGCL = '" + acred + "' " +
                           "AND CB.FVIGENCIA >= SYSDATE "; 

            iRes = db.ExecuteDS(ref dsInfo, query, CommandType.Text);
            
            query = "SELECT SN.*, " +
                    "TO_CHAR(SN.FNACIMIENTO, 'DD/MM/YYYY') FECNAC, " +
                    "TO_CHAR(SN.FINFODEPEND, 'DD/MM/YYYY') FECINFODEPEND, " +
                    "TO_CHAR(SN.FDEFUNC, 'DD/MM/YYYY') FECDEFUNC " +
                    "FROM CONSULTA_BURO CB, " +
                    "SEGMENTO_NOMBRE SN " +
                    "WHERE CB.CDGEM = '" + empresa + "' " +
                    "AND CB.CDGCL = '" + acred + "' " +
                    "AND CB.FVIGENCIA >= SYSDATE " +
                    "AND SN.CDGEM = CB.CDGEM " +
                    "AND SN.CDGCL = CB.CDGCL " +
                    "AND SN.FCONSULTA = TRUNC(CB.FCONSULTA)";
                           
            iRes = db.ExecuteDS(ref dsNom, query, CommandType.Text);

            query = "SELECT SD.*, " +
                    "TO_CHAR(SD.FRESID, 'DD/MM/YYYY') FECRESID, " +
                    "TO_CHAR(SD.FREPDIR, 'DD/MM/YYYY') FECREPDIR " +
                    "FROM CONSULTA_BURO CB, " +
                    "SEGMENTO_DIRECCION SD " +
                    "WHERE CB.CDGEM = '" + empresa + "' " +
                    "AND CB.CDGCL = '" + acred + "' " +
                    "AND CB.FVIGENCIA >= SYSDATE " +
                    "AND SD.CDGEM = CB.CDGEM " +
                    "AND SD.CDGCL = CB.CDGCL " +
                    "AND SD.FCONSULTA = TRUNC(CB.FCONSULTA) " +
                    "ORDER BY SD.FREPDIR DESC";

            iRes = db.ExecuteDS(ref dsDir, query, CommandType.Text);

            query = "SELECT SC.*, " +
                    "TO_CHAR(SC.FACTUALIZA, 'DD/MM/YYYY') FECACT, " +
                    "TO_CHAR(SC.FINICTA, 'DD/MM/YYYY') FECINICTA, " +
                    "TO_CHAR(SC.FULTPAGO, 'DD/MM/YYYY') FECULTPAGO, " +
                    "TO_CHAR(SC.FULTCOMP, 'DD/MM/YYYY') FECULTCOMP, " +
                    "TO_CHAR(SC.FFINCTA, 'DD/MM/YYYY') FECFINCTA, " +
                    "TO_CHAR(SC.FREP, 'DD/MM/YYYY') FECREP, " +
                    "TO_CHAR(SC.FSINSALDO, 'DD/MM/YYYY') FECSINSALDO, " +
                    "TO_CHAR(SC.FRECHPAGOS, 'DD/MM/YYYY') FECRECHPAGOS, " +
                    "TO_CHAR(SC.FANTHPAGOS, 'DD/MM/YYYY') FECANTHPAGOS, " +
                    "TO_CHAR(SC.FIMPMOR, 'DD/MM/YYYY') FECIMPMOR, " +
                    "TO_CHAR(SC.FREST, 'DD/MM/YYYY') FECREST, " +
                    "(SELECT DESCRIPCION FROM CAT_CVE_OBSERVACION WHERE CODIGO = SC.CDGOBS) DESCCDGOBS, " +
                    "(SELECT DESCRIPCION FROM CAT_TIPO_RESPONSABILIDAD WHERE CODIGO = SC.TIPORESP) DESCTIPORESP, " +
                    "(SELECT DESCRIPCION FROM CAT_TIPO_CUENTA WHERE CODIGO = SC.TIPOCTA) DESCTIPOCTA, " +
                    "(SELECT DESCRIPCION FROM CAT_TIPO_CREDITO WHERE CODIGO = SC.TIPOCONT) DESCTIPOCRED, " +
                    "(SELECT DESCRIPCION FROM CAT_MOP WHERE CODIGO = SC.MOP) DESCMOP, " +
                    "(SELECT DESCRIPCION FROM CAT_FREC_PAGO_BC WHERE CODIGO = SC.FRECPAGOS) DESCFRECPAGOS " +
                    "FROM CONSULTA_BURO CB, " +
                    "SEGMENTO_CUENTA SC " +
                    "WHERE CB.CDGEM = '" + empresa + "' " +
                    "AND CB.CDGCL = '" + acred + "' " +
                    "AND CB.FVIGENCIA >= SYSDATE " +
                    "AND SC.CDGEM = CB.CDGEM " +
                    "AND SC.CDGCL = CB.CDGCL " +
                    "AND SC.FCONSULTA = TRUNC(CB.FCONSULTA)";

            iRes = db.ExecuteDS(ref dsCta, query, CommandType.Text);

            query = "SELECT MOP " +
                    ",NVL((SELECT COUNT(MOP) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) NUMABR " +
                    ",NVL((SELECT COUNT(MOP) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) NUMCER " +
                    ",NVL((SELECT SUM(LIMCRD) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) LIMABR " +
                    ",NVL((SELECT SUM(LIMCRD) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) LIMCER " +
                    ",NVL((SELECT SUM(CREDMAX) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) CREDMAXABR " +
                    ",NVL((SELECT SUM(CREDMAX) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) CREDMAXCER " +
                    ",NVL((SELECT SUM(SDOACT) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) SDOACTABR " +
                    ",NVL((SELECT SUM(SDOACT) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) SDOACTCER " +
                    ",NVL((SELECT SUM(SDOVENC) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) SDOVENCABR " +
                    ",NVL((SELECT SUM(MONTO) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) MONTOABR " +
                    ",NVL((SELECT SUM(MONTO) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND MOP = S.MOP),0) MONTOCER " +
                    "FROM CONSULTA_BURO CB, " +
   	                "SEGMENTO_CUENTA S " +
                    "WHERE CB.CDGEM = '" + empresa + "' " +
                    "AND CB.CDGCL = '" + acred + "' " +
                    "AND CB.FVIGENCIA >= SYSDATE " + 
                    "AND S.CDGEM = CB.CDGEM " +
                    "AND S.CDGCL = CB.CDGCL " +
                    "AND S.FCONSULTA = TRUNC(CB.FCONSULTA) " +
                    "GROUP BY MOP, S.CDGEM, S.CDGCL, S.FCONSULTA";

            iRes = db.ExecuteDS(ref dsRes, query, CommandType.Text);

            query = "SELECT SC.*, " +
                    "TO_CHAR(SC.FREGCONS, 'DD/MM/YYYY') FECREGCONS, " +
                    "(SELECT DESCRIPCION FROM CAT_TIPO_RESPONSABILIDAD WHERE CODIGO = SC.TIPORESP) DESCTIPORESP, " +
                    "(SELECT DESCRIPCION FROM CAT_TIPO_CREDITO WHERE CODIGO = SC.TIPOCONT) DESCTIPOCONT " +
                    "FROM CONSULTA_BURO CB, " +
                    "SEGMENTO_CONSULTA SC " +
                    "WHERE CB.CDGEM = '" + empresa + "' " +
                    "AND CB.CDGCL = '" + acred + "' " +
                    "AND CB.FVIGENCIA >= SYSDATE " +
                    "AND SC.CDGEM = CB.CDGEM " +
                    "AND SC.CDGCL = CB.CDGCL " +
                    "AND SC.FCONSULTA = TRUNC(CB.FCONSULTA)";

            iRes = db.ExecuteDS(ref dsCons, query, CommandType.Text);

            if (dsNom.Tables[0].Rows.Count > 0)
                LlenaConsultaBuro(dsInfo, dsNom, dsDir, dsCta, dsRes, dsCons, nomUsuario);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getConsultaRepCredito(string tipo, string acred, string usuario, string nomUsuario, string sic, string fecha)
    {
        string empresa = cdgEmpresa;
        string cadFecha = string.Empty;
        int iRes;
        try
        {
            DataSet dsInfo = new DataSet();
            DataSet dsNom = new DataSet();
            DataSet dsDir = new DataSet();
            DataSet dsCta = new DataSet();
            DataSet dsRes = new DataSet();
            DataSet dsCons = new DataSet();
            DataSet dsMens = new DataSet();
            DataSet dsScore = new DataSet();
            string strCampo = string.Empty;
            string strTablaRep = string.Empty;
            string strTablaCta = string.Empty;
            string strTablaNom = string.Empty;
            string strTablaDir = string.Empty;
            string strTablaCons = string.Empty;
            string strTablaMsj = string.Empty;
            string strTablaScore = string.Empty;

            if (fecha == null)
                cadFecha = "AND CR.FVIGENCIA >= SYSDATE ";
            else
                cadFecha = "AND TRUNC(CR.FCONSULTA) = '" + fecha + "' ";

            if (tipo == "C")
            {
                strCampo = "CDGCL";
                strTablaRep = "CONSULTA_REP_CREDITO";
                strTablaCta = "SEGMENTO_CUENTA";
                strTablaNom = "SEGMENTO_NOMBRE";
                strTablaDir = "SEGMENTO_DIRECCION";
                strTablaCons = "SEGMENTO_CONSULTA";
                strTablaMsj = "SEGMENTO_MENSAJE";
                strTablaScore = "SEGMENTO_SCORE";
            }
            else if (tipo == "P")
            {
                strCampo = "CDGPROSP";
                strTablaRep = "CONSULTA_REP_CREDITO_P";
                strTablaCta = "SEGMENTO_CUENTA_P";
                strTablaNom = "SEGMENTO_NOMBRE_P";
                strTablaDir = "SEGMENTO_DIRECCION_P";
                strTablaCons = "SEGMENTO_CONSULTA_P";
                strTablaMsj = "SEGMENTO_MENSAJE_P";
                strTablaScore = "SEGMENTO_SCORE_P";
            }

            string query = "SELECT TO_CHAR(CR.FCONSULTA,'DD/MM/YYYY') FECHACONS " +
                           ",TO_CHAR(CR.FVIGENCIA,'DD/MM/YYYY') FECHAVIG " +
                           ",TO_CHAR(SYSDATE,'DD/MM/YYYY') FECHAIMP " +
                           ",TO_CHAR(SYSDATE,'HH24:MI:SS') HORAIMP " +
                           ",DECODE(INSTCRED, 'CC','REPORTE DE CRÉDITO (CIRCULO DE CRÉDITO)','BC','REPORTE DE BURÓ DE CRÉDITO') TITULO " +
                           ",CR.FOLIOCONS NUMCONS " +
                           ",(SELECT COUNT(*) FROM " + strTablaCta + " WHERE CDGEM = CR.CDGEM AND " + strCampo + " = CR." + strCampo + " AND FCONSULTA = TRUNC(CR.FCONSULTA) AND INSTCRED = CR.INSTCRED AND TO_CHAR(FINICTA,'YYYY') = TO_CHAR(SYSDATE,'YYYY')) REGISTROS " +
                           ",(SELECT MIN(FINICTA) FROM " + strTablaCta + " WHERE CDGEM = CR.CDGEM AND " + strCampo + " = CR." + strCampo + " AND FCONSULTA = TRUNC(CR.FCONSULTA) AND INSTCRED = CR.INSTCRED) FECCRED " +
                           ",(SELECT MAX(CREDMAX) FROM " + strTablaCta + " WHERE CDGEM = CR.CDGEM AND " + strCampo + " = CR." + strCampo + " AND FCONSULTA = TRUNC(CR.FCONSULTA) AND INSTCRED = CR.INSTCRED) MONTOCRED " +
                           "FROM " + strTablaRep + " CR " +
                           "WHERE CR.CDGEM = '" + empresa + "' " +
                           "AND CR." + strCampo + " = '" + acred + "' " +
                           "AND CR.INSTCRED = '" + sic + "' " +
                           cadFecha;

            iRes = db.ExecuteDS(ref dsInfo, query, CommandType.Text);

            query = "SELECT SN.* " +
                    ",TO_CHAR(SN.FNACIMIENTO, 'DD/MM/YYYY') FECNAC " +
                    ",TO_CHAR(SN.FINFODEPEND, 'DD/MM/YYYY') FECINFODEPEND " +
                    ",TO_CHAR(SN.FDEFUNC, 'DD/MM/YYYY') FECDEFUNC " +
                    "FROM " + strTablaRep + " CR, " +
                    strTablaNom + " SN " +
                    "WHERE CR.CDGEM = '" + empresa + "' " +
                    "AND CR." + strCampo + " = '" + acred + "' " +
                    cadFecha +
                    "AND CR.INSTCRED = '" + sic + "' " +
                    "AND SN.CDGEM = CR.CDGEM " +
                    "AND SN." + strCampo + " = CR." + strCampo + " " +
                    "AND SN.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                    "AND SN.INSTCRED = CR.INSTCRED";

            iRes = db.ExecuteDS(ref dsNom, query, CommandType.Text);

            query = "SELECT SD.* " +
                    ",TO_CHAR(SD.FRESID, 'DD/MM/YYYY') FECRESID " +
                    ",TO_CHAR(SD.FREPDIR, 'DD/MM/YYYY') FECREPDIR " +
                    "FROM " + strTablaRep + " CR, " +
                    strTablaDir + " SD " +
                    "WHERE CR.CDGEM = '" + empresa + "' " +
                    "AND CR." + strCampo + " = '" + acred + "' " +
                    cadFecha +
                    "AND CR.INSTCRED = '" + sic + "' " +
                    "AND SD.CDGEM = CR.CDGEM " +
                    "AND SD." + strCampo + " = CR." + strCampo + " " +
                    "AND SD.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                    "AND SD.INSTCRED = CR.INSTCRED " +
                    "ORDER BY SD.FREPDIR DESC";

            iRes = db.ExecuteDS(ref dsDir, query, CommandType.Text);

            query = "SELECT SC.* " +
                    ",TO_CHAR(SC.FACTUALIZA, 'DD/MM/YYYY') FECACT " +
                    ",TO_CHAR(SC.FINICTA, 'DD/MM/YYYY') FECINICTA " +
                    ",TO_CHAR(SC.FULTPAGO, 'DD/MM/YYYY') FECULTPAGO " +
                    ",TO_CHAR(SC.FULTCOMP, 'DD/MM/YYYY') FECULTCOMP " +
                    ",TO_CHAR(SC.FFINCTA, 'DD/MM/YYYY') FECFINCTA " +
                    ",TO_CHAR(SC.FREP, 'DD/MM/YYYY') FECREP " +
                    ",TO_CHAR(SC.FSINSALDO, 'DD/MM/YYYY') FECSINSALDO " +
                    ",TO_CHAR(SC.FRECHPAGOS, 'DD/MM/YYYY') FECRECHPAGOS " +
                    ",TO_CHAR(SC.FANTHPAGOS, 'DD/MM/YYYY') FECANTHPAGOS " +
                    ",TO_CHAR(SC.FIMPMOR, 'DD/MM/YYYY') FECIMPMOR " +
                    ",TO_CHAR(SC.FREST, 'DD/MM/YYYY') FECREST " +
                    ",CASE WHEN FFINCTA IS NOT NULL THEN 'C' ELSE (CASE WHEN SC.SDOVENC > 0 THEN 'V' ELSE 'A' END) END ESTATUS " +
                    ",(SELECT DESCRIPCION FROM CAT_CVE_OBSERVACION WHERE CODIGO = SC.CDGOBS AND TIPO = SC.INSTCRED) DESCCDGOBS " +
                    ",(SELECT DESCRIPCION FROM CAT_TIPO_RESPONSABILIDAD WHERE CODIGO = SC.TIPORESP AND TIPO = SC.INSTCRED) DESCTIPORESP " +
                    ",(SELECT DESCRIPCION FROM CAT_TIPO_CUENTA WHERE CODIGO = SC.TIPOCTA AND TIPO = SC.INSTCRED) DESCTIPOCTA " +
                    ",(SELECT DESCRIPCION FROM CAT_TIPO_CREDITO WHERE CODIGO = SC.TIPOCONT AND TIPO = SC.INSTCRED) DESCTIPOCRED " +
                    ",(SELECT DESCRIPCION FROM CAT_MOP WHERE CODIGO = SC.MOP) DESCMOP " +
                    ",(SELECT DESCRIPCION FROM CAT_FREC_PAGO_BC WHERE CODIGO = SC.FRECPAGOS AND TIPO = SC.INSTCRED) DESCFRECPAGOS " +
                    "FROM " + strTablaRep + " CR, " +
                    strTablaCta + " SC " +
                    "WHERE CR.CDGEM = '" + empresa + "' " +
                    "AND CR." + strCampo + " = '" + acred + "' " +
                    cadFecha +
                    "AND CR.INSTCRED = '" + sic + "' " +
                    "AND SC.CDGEM = CR.CDGEM " +
                    "AND SC." + strCampo + " = CR." + strCampo + " " +
                    "AND SC.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                    "AND SC.INSTCRED = CR.INSTCRED " +
                    "ORDER BY FFINCTA DESC NULLS FIRST, FINICTA DESC";

            iRes = db.ExecuteDS(ref dsCta, query, CommandType.Text);

            if (sic == "BC")
            {
                query = "SELECT MOP " +
                        ",NVL((SELECT COUNT(MOP) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) NUMABR " +
                        ",NVL((SELECT COUNT(MOP) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) NUMCER " +
                        ",NVL((SELECT SUM(LIMCRD) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) LIMABR " +
                        ",NVL((SELECT SUM(LIMCRD) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) LIMCER " +
                        ",NVL((SELECT SUM(CREDMAX) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) CREDMAXABR " +
                        ",NVL((SELECT SUM(CREDMAX) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) CREDMAXCER " +
                        ",NVL((SELECT SUM(SDOACT) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) SDOACTABR " +
                        ",NVL((SELECT SUM(SDOACT) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) SDOACTCER " +
                        ",NVL((SELECT SUM(SDOVENC) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) SDOVENCABR " +
                        ",NVL((SELECT SUM(MONTO) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) MONTOABR " +
                        ",NVL((SELECT SUM(MONTO) FROM SEGMENTO_CUENTA WHERE FFINCTA IS NOT NULL AND CDGEM = S.CDGEM AND CDGCL = S.CDGCL AND TRUNC(FCONSULTA) = TRUNC(S.FCONSULTA) AND INSTCRED = S.INSTCRED AND MOP = S.MOP),0) MONTOCER " +
                        "FROM CONSULTA_REP_CREDITO CR, " +
                        "SEGMENTO_CUENTA S " +
                        "WHERE CR.CDGEM = '" + empresa + "' " +
                        "AND CR.CDGCL = '" + acred + "' " +
                        cadFecha +
                        "AND CR.INSTCRED = '" + sic + "' " +
                        "AND S.CDGEM = CR.CDGEM " +
                        "AND S.CDGCL = CR.CDGCL " +
                        "AND S.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                        "AND S.INSTCRED = CR.INSTCRED " +
                        "GROUP BY MOP, S.CDGEM, S.INSTCRED, S.CDGCL, S.FCONSULTA";
            }
            else if (sic == "CC")
            {
                query = "SELECT A.DESCTIPOCONT " +
                        ",DECODE(A.ESTATUS,1,'C',0,'A',2,'V') MOP " +
                        ",COUNT (A.TIPOCONT) NUMABR " +
                        ",SUM (A.LIMCRD) LIMABR " +
                        ",SUM (A.CREDMAX) CREDMAXABR " +
                        ",SUM (A.SDOACT) SDOACTABR " +
                        ",SUM (A.SDOVENC) SDOVENCABR " +
                        ",SUM (A.MONTO) MONTOABR " +
                        ",SUM (A.SEMANAL) CREDMAXCER " +
                        ",SUM (A.QUINCENAL) SDOACTCER " +
                        ",SUM (A.MENSUAL) MONTOCER " +
                        "FROM (SELECT S.TIPOCONT " +
                              ",TC.DESCRIPCION DESCTIPOCONT " +
                              ",S.LIMCRD " +
                              ",S.CREDMAX " +
                              ",S.SDOACT " +
                              ",S.SDOVENC " +
                              ",S.MONTO " +
                              ",S.FRECPAGOS " +
                              ",CASE WHEN S.FFINCTA IS NOT NULL THEN 1 " +
                              "      ELSE (CASE WHEN S.SDOVENC > 0 THEN 2 ELSE 0 END) END ESTATUS " +
                              ",CASE WHEN S.FRECPAGOS = 'S' THEN S.MONTO ELSE 0 END SEMANAL " +
                              ",CASE WHEN S.FRECPAGOS = 'Q' THEN S.MONTO ELSE 0 END QUINCENAL " +
                              ",CASE WHEN S.FRECPAGOS = 'M' THEN S.MONTO ELSE 0 END MENSUAL " +
                              "FROM " + strTablaRep + " CR, " + strTablaCta + " S, CAT_TIPO_CREDITO TC " +
                              "WHERE CR.CDGEM = '" + empresa + "' " +
                              "AND CR." + strCampo + " = '" + acred + "' " +
                              cadFecha +
                              "AND CR.INSTCRED = '" + sic + "' " +
                              "AND S.CDGEM = CR.CDGEM " +
                              "AND S." + strCampo + " = CR." + strCampo + " " +
                              "AND S.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                              "AND S.INSTCRED = CR.INSTCRED " +
                              "AND TC.CODIGO = S.TIPOCONT " +
                              "AND TC.TIPO = S.INSTCRED) A " +
                        "GROUP BY A.TIPOCONT, A.DESCTIPOCONT, A.FRECPAGOS, A.ESTATUS";
            }

            iRes = db.ExecuteDS(ref dsRes, query, CommandType.Text);

            query = "SELECT SC.* " +
                    ",TO_CHAR(SC.FREGCONS, 'DD/MM/YYYY') FECREGCONS " +
                    ",(SELECT DESCRIPCION FROM CAT_TIPO_RESPONSABILIDAD WHERE CODIGO = SC.TIPORESP AND TIPO = SC.INSTCRED) DESCTIPORESP " +
                    ",(SELECT DESCRIPCION FROM CAT_TIPO_CUENTA WHERE CODIGO = SC.TIPOCONT AND TIPO = SC.INSTCRED) DESCTIPOCONT " +
                    "FROM " + strTablaRep + " CR, " +
                    strTablaCons + " SC " +
                    "WHERE CR.CDGEM = '" + empresa + "' " +
                    "AND CR." + strCampo + " = '" + acred + "' " +
                    cadFecha +
                    "AND CR.INSTCRED = '" + sic + "' " +
                    "AND SC.CDGEM = CR.CDGEM " +
                    "AND SC." + strCampo + " = CR." + strCampo + " " +
                    "AND SC.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                    "AND SC.INSTCRED = CR.INSTCRED";

            iRes = db.ExecuteDS(ref dsCons, query, CommandType.Text);

            query = "SELECT SM.* " +
                    ",(SELECT DESCRIPCION FROM CAT_TIPO_MENSAJE WHERE CODIGO = SM.TIPOMENS AND TIPO = SM.INSTCRED) DESCTIPOMENS " +
                    ",CASE WHEN SM.TIPOMENS = '2' THEN " +
                        "(SELECT DESCRIPCION FROM CAT_LEYENDA WHERE CODIGO = SM.LEYENDA AND TIPO = SM.INSTCRED) " +
                    "ELSE " +
                        "SM.LEYENDA " +
                    "END DESCLEYENDA " +
                    "FROM " + strTablaRep + " CR, " +
                    strTablaMsj + " SM " +
                    "WHERE CR.CDGEM = '" + empresa + "' " +
                    "AND CR." + strCampo + " = '" + acred + "' " +
                    cadFecha +
                    "AND CR.INSTCRED = '" + sic + "' " +
                    "AND SM.CDGEM = CR.CDGEM " +
                    "AND SM." + strCampo + " = CR." + strCampo + " " +
                    "AND SM.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                    "AND SM.INSTCRED = CR.INSTCRED";

            iRes = db.ExecuteDS(ref dsMens, query, CommandType.Text);

            query = "SELECT * " +
                    "FROM " + strTablaScore + " SS " +
                    "JOIN " + strTablaRep + " CR ON SS." + strCampo + " = CR." + strCampo + " AND SS.FCONSULTA = TRUNC(CR.FCONSULTA) " +
                    "WHERE SS.CDGEM = '" + empresa + "' " +
                    "AND SS." + strCampo + " = '" + acred + "' " +
                    " " + cadFecha + " ";

            iRes = db.ExecuteDS(ref dsScore, query, CommandType.Text);

            if (dsNom.Tables[0].Rows.Count > 0)
                LlenaConsultaRepCredito(sic, dsInfo, dsNom, dsDir, dsCta, dsRes, dsCons, dsMens, dsScore, nomUsuario);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getContrato(string grupo, string ciclo, string usuario)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        int iRes;
        try
        {
            DataSet dsContrato = new DataSet();
            DataSet dsCliente = new DataSet();

            query = "SELECT CODIGO_GPO " +
                    ",NOMBRE_GPO " +
                    ",MUNICIPIO_CTE " +
                    ",ENTIDAD_CTE " +
                    ",FECHA_INICIO " +
                    ",TASA_CTE " +
                    ",TASA_ANUAL " +
                    ",PLAZO " +
                    ",CANTIDAD_NUMERO " +
                    ",FECHA_FIN_LETRA " +
                    ",PRESIDENTE " +
                    ",TESORERO " +
                    ",SECRETARIO " +
                    ",SUPERVISOR " +
                    ",PAGO_PARCIAL " +
                    ",TOTAL_A_PAGAR " +
                    ",PER_SINGULAR " +
                    ",PER_PLURAL " +
                    ",NOM_SUCURSAL " +
                    ",NOM_INTEG " +
                    ",NOM_GERENTE_SUC " +
                    ",NOM_ASESOR " +
                    ",DIR_SUCURSAL " +
                    ",CUENTA_BANCARIA " +
                    ",(SELECT (CF.IVA/CF.PORCENTAJE) " +
                      "FROM CF,PRN " +
                      "WHERE PRN.CDGEM = CF.CDGEM " +
                      "AND PRN.CDGFDI = CF.CDGFDI " +
                      "AND PRN.CDGEM = C.CODIGO_EMP " +
                      "AND PRN.CDGNS = C.CODIGO_GPO " +
                      "AND PRN.CICLO = C.NO_CICLO) IVA " +
                    ",TIPO_PROD " +
                    ",PARC_SIN_IVA " +
                    ",TOTAL_SIN_IVA " +
                    ",FECHA_PAGO " +
                    "FROM IMPCONT C " +
                    "WHERE C.CODIGO_EMP = '" + empresa + "' " +
                    "AND C.CODIGO_GPO = '" + grupo + "' " +
                    "AND C.NO_CICLO = '" + ciclo + "' " +
                    "AND C.CDGPE = '" + usuario + "' " +
                    "AND C.TIPO_DOC = 'CONTRATO'";

            iRes = db.ExecuteDS(ref dsContrato, query, CommandType.Text);

            query = "SELECT CL.CODIGO " +
                    ",NOMBREC(CL.CDGEM,CL.CODIGO,'I','N',NULL,NULL,NULL,NULL) NOMBRE_CL " +
                    ",DECODE(CL.EDOCIVIL,'S','SOLTERO (A)','C','CASADO (A)','V','VIUDO (A)','D','DIVORCIADO (A)','U','UNION LIBRE','','NO ESPECIFICADO') EDO_CIVIL " +
                    ",NVL(PI.NOMBRE,'NO ESPECIFICADO') OCUPACION " +
                    ",CL.CALLE " +
                    ",(SELECT NOMBRE FROM COL WHERE CDGEF = CL.CDGEF AND CDGMU = CL.CDGMU AND CDGLO = CL.CDGLO AND CODIGO = CL.CDGCOL) COLONIA " +
                    ",(SELECT NOMBRE FROM LO WHERE CDGEF = CL.CDGEF AND CDGMU = CL.CDGMU AND CODIGO = CL.CDGLO) LOCALIDAD " +
                    ",(SELECT NOMBRE FROM MU WHERE CDGEF = CL.CDGEF AND CODIGO = CL.CDGMU) MUNIC " +
                    ",(SELECT CDGPOSTAL FROM COL WHERE CDGEF = CL.CDGEF AND CDGMU = CL.CDGMU AND CDGLO = CL.CDGLO AND CODIGO = CL.CDGCOL) CODPOSTAL " +
                    ",CL.IFE " +
                    "FROM PRC, CL, SC, PI " +
                    "WHERE PRC.CDGEM = '" + empresa + "' " +
                    "AND PRC.CDGNS = '" + grupo + "' " +
                    "AND PRC.CICLO = '" + ciclo + "' " +
                    "AND PRC.SITUACION = 'E' " +
                    "AND PRC.CANTENTRE > 0 " +
                    "AND CL.CDGEM = PRC.CDGEM " +
                    "AND CL.CODIGO = PRC.CDGCL " +
                    "AND SC.CDGEM = PRC.CDGEM " +
                    "AND SC.CDGNS = PRC.CDGNS " +
                    "AND SC.CICLO = PRC.CICLO " +
                    "AND SC.CDGCL = PRC.CDGCL " +
                    "AND PI.CDGEM = SC.CDGEM " +
                    "AND PI.CDGCL = SC.CDGCL " +
                    "AND PI.PROYECTO = SC.CDGPI";

            iRes = db.ExecuteDS(ref dsCliente, query, CommandType.Text);

            if (dsContrato.Tables[0].Rows.Count > 0 && dsCliente.Tables[0].Rows.Count > 0)
                LlenaContrato(dsContrato, dsCliente);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getControlPagos(string id, string fecha, string grupo, string ciclo, string acred, string usuario, 
                                 string nomUsuario, string puesto, string region, string coord, string asesor)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        string strReg = string.Empty;
        string strCoord = string.Empty;
        string strAsesor = string.Empty;
        int iRes;
        try
        {
            DataSet ds = new DataSet();
            DataSet dsGrupo = new DataSet();
            DataSet dsAcred = new DataSet();

            if (puesto == "A")
                strAsesor = "AND PRN.CDGOCPE = '" + usuario + "' ";
            else
                strAsesor = "AND PRN.CDGCO IN (SELECT DISTINCT(CDGCO) FROM PCO WHERE CDGEM = '" + empresa + "' AND CDGPE = '" + usuario + "') ";

            if (grupo == null && ciclo == null && acred == null)
            {
                iRes = db.myExecuteNonQuery("SP_REP_SIT_CONTROL_PAGOS", CommandType.StoredProcedure,
                      oP.ParamsContPagos(empresa, fecha, region, coord, asesor, usuario));

                query = "SELECT * " +
                        "FROM REP_SIT_CONTROL_PAGOS " +
                        "WHERE CDGEM = '" + empresa + "' " +
                        "AND CDGPE = '" + usuario + "' " +
                        "ORDER BY DIFERENCIA";

                iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);
            }
            else if (grupo != null && ciclo != null && acred == null)
            {
                query = "SELECT CO.NOMBRE COORD " +
                        ",PRN.CDGNS || PRN.CICLO CONTRATO " +
                        ",NOMBREC(NULL, NULL, 'I', 'N', A.NOMBRE1, A.NOMBRE2, A.PRIMAPE, A.SEGAPE) GERENTE " +
                        ",NOMBREC(NULL, NULL, 'I', 'N', B.NOMBRE1, B.NOMBRE2, B.PRIMAPE, B.SEGAPE) ASESOR " +
                        ",PRN.CDGNS " +
                        ",NS.NOMBRE GRUPO " +
                        ",PRN.CICLO " +
                        ",TO_CHAR(PRN.INICIO,'DD/MM/YYYY') FINICIO " +
                        ",TO_CHAR(DECODE(NVL(PRN.PERIODICIDAD,''), " +
                                      "'S', PRN.INICIO + (7 * NVL(PRN.PLAZO,0)), " +
                                      "'Q', PRN.INICIO + (15 * NVL(PRN.PLAZO,0)), " +
                                      "'C', PRN.INICIO + (14 * NVL(PRN.PLAZO,0)), " +
                                      "'M', PRN.INICIO + (30 * NVL(PRN.PLAZO,0)), " +
                                      "'', ''),'DD/MM/YYYY') FFIN " +
                        "FROM PRN, NS, CO, PE A, PE B " +
                        "WHERE PRN.CDGEM = '" + empresa + "' " + 
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND SITUACION IN ('E', 'L') " +
                        strAsesor +
                        "AND NS.CDGEM = PRN.CDGEM " +
                        "AND NS.CODIGO = PRN.CDGNS " +
                        "AND CO.CDGEM = PRN.CDGEM " +
                        "AND CO.CODIGO = PRN.CDGCO " +
                        "AND A.CDGEM = CO.CDGEM " +
                        "AND A.CODIGO = CO.CDGPE " +
                        "AND B.CDGEM = PRN.CDGEM " +
                        "AND B.CODIGO = PRN.CDGOCPE";

                iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);

                query = "SELECT COORD " +
                            ",CONTRATO " +
                            ",ASESOR " +
                            ",CDGNS " +
                            ",GRUPO " +
                            ",CDGCL " +
                            ",ACRED " +
                            ",CICLO " +
                            ",FINICIO " +
                            ",FFIN " +
                            ",CANTENTRE " +
                            ",SITUA " +
                            ",PAGOCOMP " +
                            ",SALDOTOTAL " +
                            ",DIASMORA " +
                            ",PARCIALIDAD " +
                            ",PAGO_SEM " +
                            ",PAGO_EXT " +
                            ",APORTACRED " +
                            ",(PAGO_SEM + PAGO_EXT + APORTACRED) PAGOREAL " +
                            ",((PAGO_SEM + PAGO_EXT + APORTACRED) - PAGOCOMP) DIFERENCIA " +
                            ",DEVACRED " +
                            ",AHORROACRED " +
                            ",MULTAACRED " +
                            ",(TOTALPAGAR - (PAGO_SEM + PAGO_EXT + APORTACRED)) SALDO " +
                            "FROM (SELECT A.COORD " +                //DEFINE EL ORIGEN DE DATOS AGRUPADO
                            ",A.CONTRATO " +
                            ",A.ASESOR " +
                            ",A.CDGNS " +
                            ",A.GRUPO " +
                            ",A.CDGCL " +
                            ",A.ACRED " +
                            ",A.CICLO " +
                            ",A.FINICIO " +
                            ",A.FFIN " +
                            ",A.CANTENTRE " +
                            ",A.SITUA " +
                            ",A.PAGOCOMP " +
                            ",A.SALDOTOTAL " +
                            ",A.DIASMORA " +
                            ",A.PARCIALIDAD " +
                            ",A.TOTALPAGAR " +
                            ",NVL(SUM(PAGOSEM),0) PAGO_SEM " +              //SUMATORIA DE LOS PAGOS SEMANALES
                            ",NVL(SUM(PAGOEXT),0) PAGO_EXT " +              //SUMATORIA DE LOS PAGOS EXTEMPORANEOS
                            ",NVL(SUM(APORT_ACRED),0) APORTACRED " +        //SUMATORIA DE LAS APORTACIONES DE CREDITO
                            ",SUM(DEV_ACRED) DEVACRED " +
                            ",SUM(AHORRO_ACRED) AHORROACRED " +
                            ",SUM(MULTA_ACRED) MULTAACRED " +
                             "FROM (SELECT CO.NOMBRE COORD " +
                             ",PRN.CDGNS || PRN.CICLO CONTRATO " +
                             ",NOMBREC (NULL, NULL, 'I', 'N', PE.NOMBRE1, PE.NOMBRE2, PE.PRIMAPE, PE.SEGAPE) ASESOR " +
                             ",PRN.CDGNS " +
                             ",NS.NOMBRE GRUPO " +
                             ",PRC.CDGCL " +
                             ",NOMBREC(CL.CDGEM, CL.CODIGO, 'I', 'N', NULL, NULL, NULL, NULL) ACRED " +
                             ",PRN.CICLO " +
                             ",TO_CHAR(PRN.INICIO, 'DD/MM/YYYY') FINICIO " +
                             ",TO_CHAR(DECODE (NVL (PRN.PERIODICIDAD, ''), " +
                                               "'S', PRN.INICIO + (7 * NVL(PRN.PLAZO, 0)), " +
                                               "'Q', PRN.INICIO + (15 * NVL(PRN.PLAZO, 0)), " +
                                               "'C', PRN.INICIO + (14 * NVL(PRN.PLAZO, 0)), " +
                                               "'M', PRN.INICIO + (30 * NVL(PRN.PLAZO, 0)), " +
                                               "'', ''), 'DD/MM/YYYY') FFIN " +
                            ",PRC.CANTENTRE " +
                            ",CASE WHEN CSA.TIPO = 'S' THEN " +
                                "CSA.PAGO_REAL " +
                            "END PAGOSEM " +
                            ",CASE WHEN CSA.TIPO = 'P' THEN " +
                                "CSA.PAGO_REAL " +
                            "END PAGOEXT " +
                            ",CSA.PAGO_REAL " +
                            ",CSA.APORT_ACRED " +
                            ",CSA.DEV_ACRED " +
                            ",CSA.AHORRO_ACRED " +
                            ",CSA.MULTA_ACRED " +
                            ",DECODE(PRN.SITUACION,'E','ENTREGADO','L','LIQUIDADO') SITUA " +
                            ",ROUND((PRC.CANTENTRE / PRN.CANTENTRE) * (PagoVencidoCapitalPrN(PrN.CdgEm,  PRN.CdgNs, PrN.Ciclo,PrN.CantEntre, PrN.Tasa, PrN.Plazo,PrN.Periodicidad, PrN.CdgMCI, Prn.Inicio, Prn.DiaJunta,Prn.MULTPER, PrN.PeriGrCap, PrN.PeriGrInt, PrN.DesfasePago, PrN.CdgTI,PrN.ModoApliReca,'" + fecha + "',null,'S')),2) PAGOCOMP " +
                            ",ROUND((PRC.CANTENTRE / PRN.CANTENTRE) * (SALDOTOTALPRN(PrN.CdgEm, PrN.CdgNS, PrN.Ciclo, PrN.CantEntre, PrN.Tasa,    PrN.Plazo, PrN.Periodicidad, PrN.CdgMCI, PrN.Inicio, PrN.DiaJunta,    PrN.MULTPER, PrN.PeriGrCap, PrN.PeriGrInt, PrN.DesfasePago, PrN.CdgTI,    PrN.ModoApliReca, '" + fecha + "')),2) SALDOTOTAL " +
                            ",CASE WHEN PRN.SITUACION = 'E' THEN " +
                                "(SELECT DIAS_MORA FROM TBL_CIERRE_DIA WHERE CDGEM = PRN.CDGEM AND CDGCLNS = PRN.CDGNS AND CLNS = 'G' AND CICLO = PRN.CICLO AND FECHA_CALC = '" + fecha + "') " +
                            "ELSE " +
                                "0 " +
                            "END DIASMORA " +
                            ",ROUND((PRC.CANTENTRE / PRN.CANTENTRE) * PARCIALIDADPrN (PrN.CdgEm, PrN.CdgNs, PrN.Ciclo, NVL(PrN.cantentre,PrN.Cantautor), PrN.Tasa, PrN.Plazo, PrN.Periodicidad, PrN.CdgMCI, PrN.Inicio,    PrN.DiaJunta, PrN.MULTPER, PrN.PeriGrCap, PrN.PeriGrInt, PrN.DesFasePago, PrN.CdgTi, NULL),3) PARCIALIDAD " +
                            ",round((PRC.CANTENTRE / PRN.CANTENTRE) * (round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRN.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRN.cantentre,0))/(4 * 100), " +
                                           "'Q', nvl(PRN.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRN.cantentre,0) * 15)/(30 * 100), " +
                                           "'C', nvl(PRN.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRN.cantentre,0))/(2 * 100), " +
                                           "'M', nvl(PRN.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRN.cantentre,0))/(100), " +
                                           "'',  ''),2)) ,2) TOTALPAGAR " +
                            "FROM PRN, PRC, CL, CO, PE, NS, " +
                            "(SELECT CDGEM, CDGNS, CICLO, CDGCL, TIPO, SUM(PAGOREAL) PAGO_REAL, " +
                              "SUM(APORT) APORT_ACRED, SUM(DEVOLUCION) DEV_ACRED, SUM(AHORRO) AHORRO_ACRED, SUM(MULTA) MULTA_ACRED " +
                              "FROM CONTROL_PAGOS_ACRED " +
                              "WHERE CDGEM = '" + empresa + "' " +
                              "AND CDGNS = '" + grupo + "' " +
                              "AND CICLO = '" + ciclo + "' " +
                              "AND FREALPAGO <= '" + fecha + "' " +
                              "GROUP BY CDGEM, CDGNS, CICLO, CDGCL, TIPO) CSA " +
                            "WHERE PRN.CDGEM = '" + empresa + "' " +
                            "AND PRN.CDGNS = '" + grupo + "' " +
                            "AND PRN.CICLO = '" + ciclo + "' " +
                            "AND PRN.CANTENTRE > 0 " +
                            "AND PRN.SITUACION IN ('E', 'L') " +
                            strAsesor +
                            "AND PRC.CDGEM = PRN.CDGEM " +
                            "AND PRC.CDGNS = PRN.CDGNS " +
                            "AND PRC.CICLO = PRN.CICLO " +
                            "AND PRC.SITUACION IN ('E', 'L') " +
                            "AND PRC.CANTENTRE > 0 " +
                            "AND CSA.CDGEM = PRC.CDGEM " +
                            "AND CSA.CDGNS = PRC.CDGNS " +
                            "AND CSA.CICLO = PRC.CICLO " +
                            "AND CSA.CDGCL = PRC.CDGCL " +
                            "AND CL.CDGEM = PRC.CDGEM " +
                            "AND CL.CODIGO = PRC.CDGCL " +
                            "AND CO.CDGEM = PRN.CDGEM " +
                            "AND CO.CODIGO = PRN.CDGCO " +
                            "AND PE.CDGEM = PRN.CDGEM " +
                            "AND PE.CODIGO = PRN.CDGOCPE " +
                            "AND NS.CDGEM = PRN.CDGEM " +
                            "AND NS.CODIGO = PRN.CDGNS) A " +
                            "GROUP BY A.COORD " +
                            ",A.CONTRATO " +
                            ",A.ASESOR " +
                            ",A.CDGNS " +
                            ",A.GRUPO " +
                            ",A.CDGCL " +
                            ",A.ACRED " +
                            ",A.CICLO " +
                            ",A.FINICIO " +
                            ",A.FFIN " +
                            ",A.CANTENTRE " +
                            ",A.SITUA " +
                            ",A.PAGOCOMP " +
                            ",A.SALDOTOTAL " +
                            ",A.DIASMORA " +
                            ",A.PARCIALIDAD " +
                            ",A.TOTALPAGAR) " +
                            "ORDER BY DIFERENCIA";

                iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);
            }
            else if (grupo != null && ciclo != null && acred != null)
            {
                query = "SELECT CO.NOMBRE COORD " +
                        ",PRN.CDGNS || PRN.CICLO CONTRATO " +
                        ",NOMBREC(NULL, NULL, 'I', 'N', A.NOMBRE1, A.NOMBRE2, A.PRIMAPE, A.SEGAPE) GERENTE " +
                        ",NOMBREC(NULL, NULL, 'I', 'N', B.NOMBRE1, B.NOMBRE2, B.PRIMAPE, B.SEGAPE) ASESOR " +
                        ",PRN.CDGNS " +
                        ",NS.NOMBRE GRUPO " +
                        ",PRN.CICLO " +
                        ",TO_CHAR(PRN.INICIO,'DD/MM/YYYY') FINICIO " +
                        ",TO_CHAR(DECODE(NVL(PRN.PERIODICIDAD,''), " +
                                      "'S', PRN.INICIO + (7 * NVL(PRN.PLAZO,0)), " +
                                      "'Q', PRN.INICIO + (15 * NVL(PRN.PLAZO,0)), " +
                                      "'C', PRN.INICIO + (14 * NVL(PRN.PLAZO,0)), " +
                                      "'M', PRN.INICIO + (30 * NVL(PRN.PLAZO,0)), " +
                                      "'', ''),'DD/MM/YYYY') FFIN " +
                        "FROM PRN, NS, CO, PE A, PE B " +
                        "WHERE PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND PRN.SITUACION IN ('E', 'L') " +
                        strAsesor +
                        "AND NS.CDGEM = PRN.CDGEM " +
                        "AND NS.CODIGO = PRN.CDGNS " +
                        "AND CO.CDGEM = PRN.CDGEM " +
                        "AND CO.CODIGO = PRN.CDGCO " +
                        "AND A.CDGEM = CO.CDGEM " +
                        "AND A.CODIGO = CO.CDGPE " +
                        "AND B.CDGEM = PRN.CDGEM " +
                        "AND B.CODIGO = PRN.CDGOCPE";

                iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);

                query = "SELECT CSA.* " +
                       ",TO_CHAR(CSA.FREALPAGO,'DD/MM/YYYY') FPAGO " +
                       ",DECODE(CSA.TIPO,'S','SEMANAL','P','EXTEMPORANEO') TIPOREG " +
                       ",DECODE(CSA.ASISTENCIA, 'A', 'Asistencia', " +
                                          "'R', 'Retardo', " +
                                          "'F', 'Falta', " +
                                          "'P', 'Permiso', " +
                                          "'MP', 'Mandó Pago') ASIST " +
                       ",NS.NOMBRE GRUPO " +
                       ",NOMBREC(CL.CDGEM, CL.CODIGO,'I','N',NULL,NULL,NULL,NULL) NOMBRE_CL " +
                       ",TO_CHAR(PRN.INICIO,'DD/MM/YYYY') FINICIO " +
                       ",TO_CHAR(DECODE(NVL(PRN.PERIODICIDAD,''), " +
                                            "'S', PRN.INICIO + (7 * NVL(PRN.PLAZO,0)), " +
                                            "'Q', PRN.INICIO + (15 * NVL(PRN.PLAZO,0)), " +
                                            "'C', PRN.INICIO + (14 * NVL(PRN.PLAZO,0)), " +
                                            "'M', PRN.INICIO + (30 * NVL(PRN.PLAZO,0)), " +
                                            "'', ''),'DD/MM/YYYY') FFIN " +
                       ",PRC.CANTENTRE " +
                       ",CO.NOMBRE COORD " +
                       ",NOMBREC(NULL, NULL, 'I', 'N', PE.NOMBRE1, PE.NOMBRE2, PE.PRIMAPE, PE.SEGAPE) ASESOR " +
                       "FROM CONTROL_PAGOS_ACRED CSA, PRN, PRC, NS, CL, CO, PE " +
                       "WHERE CSA.CDGEM = '" + empresa + "' " +
                       "AND CSA.CDGNS = '" + grupo + "' " +
                       "AND CSA.CICLO = '" + ciclo + "' " +
                       "AND CSA.CDGCL = '" + acred + "' " +
                       "AND CSA.FREALPAGO <= '" + fecha + "' " +
                       "AND PRN.CDGEM = CSA.CDGEM " +
                       "AND PRN.CDGNS = CSA.CDGNS " +
                       "AND PRN.CICLO = CSA.CICLO " +
                       strAsesor +
                       "AND PRC.CDGEM = CSA.CDGEM " +
                       "AND PRC.CDGNS = CSA.CDGNS " +
                       "AND PRC.CICLO = CSA.CICLO " +
                       "AND PRC.CDGCL = CSA.CDGCL " +
                       "AND NS.CDGEM = CSA.CDGEM " +
                       "AND NS.CODIGO = CSA.CDGNS " +
                       "AND CL.CDGEM = CSA.CDGEM " +
                       "AND CL.CODIGO = CSA.CDGCL " +
                       "AND CO.CDGEM = PRN.CDGEM " +
                       "AND CO.CODIGO = PRN.CDGCO " +
                       "AND PE.CDGEM = PRN.CDGEM " +
                       "AND PE.CODIGO = PRN.CDGOCPE " +
                       "ORDER BY CSA.FREALPAGO, CSA.SECUENCIA";

                iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);
            }

            query = "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY') FECHAIMP " +
                    ",TO_CHAR(SYSDATE,'HH24:MI:SS') HORAIMP " +
                    "FROM DUAL";

            iRes = db.ExecuteDS(ref ds, query, CommandType.Text);

            if (dsGrupo.Tables[0].Rows.Count > 0 || dsAcred.Tables[0].Rows.Count > 0)
                LlenaControlPagos(id, ds, dsGrupo, dsAcred, Convert.ToDateTime(fecha), nomUsuario);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getControlSemanalEmp(string grupo, string ciclo)
    {
        string empresa = cdgEmpresa;
        int iRes;
        try
        {
            DataSet dsGrupo = new DataSet();
            DataSet dsAcred = new DataSet();

            string query = "SELECT PRN.* " +
                           ",DECODE(NVL(PRN.PERIODICIDAD, ''), " +
                                   "'S', PRN.INICIO + (7 * NVL(PRN.PLAZO,0)), " +
                                   "'Q', PRN.INICIO + (15 * NVL(PRN.PLAZO,0)), " +
                                   "'C', PRN.INICIO + (14 * NVL(PRN.PLAZO,0)), " +
                                   "'M', PRN.INICIO + (30 * NVL(PRN.PLAZO,0)), " +
                                   "'', '') FECHAFIN " +
                           ",NS.NOMBRE GRUPO " +
                           ",NOMBREC(NULL,NULL,NULL,'A',A.NOMBRE1,A.NOMBRE2,A.PRIMAPE,A.SEGAPE) NOMBREGERENTE " +
                           ",NOMBREC(NULL,NULL,NULL,'A',B.NOMBRE1,B.NOMBRE2,B.PRIMAPE,B.SEGAPE) NOMBRECOORDINADOR " +
                           ",NOMBREC(NULL,NULL,NULL,'A',PE.NOMBRE1,PE.NOMBRE2,PE.PRIMAPE,PE.SEGAPE) NOMBREASESOR " +
                           ",CO.NOMBRE COORDINACION " +
                           ",CS.SEMANA " +
                           ",CS.PAGOTEO " +
                           ",CS.PAGOREAL " +
                           ",CS.APORT " +
                           ",CS.DEVOLUCION " +
                           ",CS.MORA " +
                           ",CS.AHORRO " +
                           ",CS.MULTA " +
                           ",CS.OBSERVACION " +
                           ",TO_CHAR(CS.FREALPAGO, 'DD/MM/YYYY') FECHA " +
                           ",CS.SECUENCIA " +
                           ",DECODE(CS.TIPO,'S','SEMANAL','P','EXTEMPORANEO') TIPOREG " +
                           "FROM PRN, NS, CO, PE, PE A, CONTROL_PAGOS CS, PE B " +
                           "WHERE PRN.CDGEM = '" + empresa + "' " +
                           "AND PRN.CDGNS = '" + grupo + "' " +
                           "AND PRN.CICLO = '" + ciclo + "' " +
                           "AND CS.CDGEM = PRN.CDGEM " +
                           "AND CS.CDGNS = PRN.CDGNS " +
                           "AND CS.CICLO = PRN.CICLO " +
                           "AND NS.CDGEM = PRN.CDGEM " +
                           "AND NS.CODIGO = PRN.CDGNS " +
                           "AND CO.CDGEM = NS.CDGEM " +
                           "AND CO.CODIGO = NS.CDGCO " +
                           "AND PRN.CDGEM = PE.CDGEM " +
                           "AND PRN.CDGOCPE = PE.CODIGO " +
                           "AND CO.CDGEM = A.CDGEM " +
                           "AND CO.CDGPE = A.CODIGO " +
                           "AND PE.CDGEM = B.CDGEM " +
                           "AND PE.CALLE = B.TELEFONO " +
                           "ORDER BY CS.FREALPAGO, CS.SECUENCIA";

            iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);

            //Control semanal por acreditado
            query = "SELECT CASE WHEN PRC.CDGCL = PRN.PRESIDENTE THEN 'P' " +
                    "WHEN PRC.CDGCL = PRN.TESORERO THEN 'T' " +
                    "WHEN PRC.CDGCL = PRN.SECRETARIO THEN 'SE' " +
                    "END PUESTO " +
                    ",NOMBREC(NULL,NULL,NULL,'A',CL.NOMBRE1,CL.NOMBRE2,CL.PRIMAPE,CL.SEGAPE) CLIENTE " +
                    ",round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(4 * 100), " +
                                                           "'Q', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0) * 15)/(30 * 100), " +
                                                           "'C', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(2 * 100), " +
                                                           "'M', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(100), " +
                                                           "'',  ''),0) as TOTAL_A_PAGAR " +
                    ",PRC.CANTENTRE " +
                    ",round((round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(4 * 100), " +
                                                                  "'Q', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0) * 15)/(30 * 100), " +
                                                                  "'C', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(2 * 100), " +
                                                                  "'M', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(100), " +
                                                                  "'',  ''),0)) / PRN.PLAZO,2) AS PARCIALIDAD " +
                    ",CS.SEMANA " +
                    ",CS.PAGOREAL " +
                    ",CS.APORT " +
                    ",CS.DEVOLUCION " +
                    ",CS.AHORRO " +
                    ",CS.MULTA " +
                    ",DECODE(CS.ASISTENCIA, 'A', 'ASISTENCIA', " +
                                          "'R', 'RETARDO', " +
                                          "'F', 'FALTA', " +
                                          "'P', 'PERMISO', " +
                                          "'MP', 'MANDÓ PAGO') ASIST " +
                    ",TO_CHAR(CS.FREALPAGO, 'DD/MM/YYYY') FECHA " +
                    ",CS.SECUENCIA " +
                    ",CS.TIPO " +
                    "FROM PRC,PRN,CL,CONTROL_PAGOS_ACRED CS " +
                    "WHERE PRN.CDGEM = PRC.CDGEM  " +
                    "AND PRN.CDGNS = PRC.CDGNS  " +
                    "AND PRN.CICLO = PRC.CICLO  " +
                    "AND PRC.CDGEM = CL.CDGEM  " +
                    "AND PRC.CDGCL = CL.CODIGO  " +
                    "AND PRN.CDGEM = '" + empresa + "'  " +
                    "AND PRN.CDGNS = '" + grupo + "' " +
                    "AND PRN.CICLO = '" + ciclo + "' " +
                    "AND PRC.CANTENTRE > 0  " +
                    "AND CS.CDGEM = PRC.CDGEM " +
                    "AND CS.CDGNS = PRC.CDGNS " +
                    "AND CS.CICLO = PRC.CICLO " +
                    "AND CS.CDGCL = PRC.CDGCL " +
                    //"ORDER BY CL.PRIMAPE, CL.SEGAPE, CL.NOMBRE1";
                    "ORDER BY CS.FREALPAGO, CS.SECUENCIA, CL.PRIMAPE, CL.SEGAPE, CL.NOMBRE1";

            iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);

            if (dsGrupo.Tables[0].Rows.Count > 0 && dsAcred.Tables[0].Rows.Count > 0)
            {
                LlenaControlSemanalEmp(dsGrupo, dsAcred);
            }
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getControlSemanalImp(string grupo, string ciclo)
    {
        string empresa = cdgEmpresa;
        int iRes;
        try
        {
            DataSet dsGrupo = new DataSet();
            DataSet dsAcred = new DataSet();

            string query = "SELECT PRN.*, " +
                           "DECODE(nvl(PRN.periodicidad,''), " +
                                   "'S', PRN.inicio + (7 * nvl(PRN.plazo,0)), " +
                                   "'Q', PRN.inicio + (15 * nvl(PRN.plazo,0)), " +
                                   "'C', PRN.inicio + (14 * nvl(PRN.plazo,0)), " +
                                   "'M', PRN.inicio + (30 * nvl(PRN.plazo,0)), " +
                                   "'', '') AS FECHAFIN, " +
                           "NS.NOMBRE AS GRUPO, " +
                           "NVL(TRIM(A.nombre1),'') || ' ' || NVL(TRIM(A.nombre2),'') || ' ' || NVL(TRIM(A.primape),'')|| ' ' || NVL(TRIM(A.segape),'') AS NOMBREGERENTE, " +
                           "NVL(TRIM(B.nombre1),'') || ' ' || NVL(TRIM(B.nombre2),'') || ' ' || NVL(TRIM(B.primape),'')|| ' ' || NVL(TRIM(B.segape),'') AS NOMBRECOORDINADOR, " +
                           "NVL(TRIM(pe.nombre1),'') || ' ' || NVL(TRIM(pe.nombre2),'') || ' ' || NVL(TRIM(pe.primape),'')|| ' ' || NVL(TRIM(pe.segape),'') AS NOMBREASESOR, " +
                           "CO.NOMBRE AS COORDINACION " +
                           "FROM PRN, NS, CO, PE, PE A, PE B " +
                           "WHERE PRN.CDGEM = '" + empresa + "' " +
                           "AND PRN.CDGNS = '" + grupo + "' " +
                           "AND PRN.CICLO = '" + ciclo + "' " +
                           "AND NS.CDGEM = PRN.CDGEM " +
                           "AND NS.CODIGO = PRN.CDGNS " +
                           "AND CO.CDGEM = NS.CDGEM " +
                           "AND CO.CODIGO = NS.CDGCO " +
                           "AND PRN.CDGEM = PE.CDGEM " +
                           "AND PRN.CDGOCPE = PE.CODIGO " +
                           "AND CO.CDGEM = A.CDGEM " +
                           "AND CO.CDGPE = A.CODIGO " +
                           "AND PE.CDGEM = B.CDGEM " +
                           "AND PE.CALLE = B.TELEFONO";

            iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);

            //Control semanal Impreso
            query = "SELECT CASE WHEN PRC.CDGCL = PRN.PRESIDENTE THEN 'P' " +
                    "WHEN PRC.CDGCL = PRN.TESORERO THEN 'T' " +
                    "WHEN PRC.CDGCL = PRN.SECRETARIO THEN 'SE' " +
                    "END PUESTO, " +
		            "TRIM(CL.PRIMAPE) || ' ' || TRIM(CL.SEGAPE) || ' ' || TRIM(CL.NOMBRE1) || ' ' || NVL(CL.NOMBRE2,'') CLIENTE, " +
                    "round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(4 * 100), " +
                                			               "'Q', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0) * 15)/(30 * 100), " +
											               "'C', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(2 * 100), " +
                                			               "'M', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(100), " +
                                			               "'',  ''),0) as TOTAL_A_PAGAR, " +
                    "PRC.CANTENTRE, " +
                    "round((round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(4 * 100), " +
                                					              "'Q', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0) * 15)/(30 * 100), " +
                                					              "'C', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(2 * 100), " +
                                					              "'M', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(100), " +
                                					              "'',  ''),0)) / PRN.PLAZO,2) AS PARCIALIDAD " +  
		            ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,1) <= SYSDATE AND PRN.PLAZO >= 1 THEN " +
                          "(SELECT NVL(SUM(PAGO_REAL),0) " +
				          "FROM CONTROL_PAGOS_DETALLE " +
                          "WHERE CDGEM = PRN.CDGEM " +
                          "AND CDGCL = PRC.CDGCL " +
				          "AND CDGCLNS = PRN.CDGNS " +
                          "AND CICLO = PRN.CICLO " +
                          "AND FMOVIMIENTO BETWEEN PRN.INICIO AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,1) " +
                          "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,1)) < 7 AND PRN.PLAZO >= 1 THEN " +
                  	      "(SELECT NVL(SUM(PAGO_REAL),0) " +
					      "FROM CONTROL_PAGOS_DETALLE " +
                  		  "WHERE CDGEM = PRN.CDGEM " +
						  "AND CDGCL = PRC.CDGCL " +
                  		  "AND CDGCLNS = PRN.CDGNS " +
                  		  "AND CICLO = PRN.CICLO " +
                  		  "AND FMOVIMIENTO BETWEEN PRN.INICIO AND SYSDATE " +
                  		  "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_1 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,2) <= SYSDATE AND PRN.PLAZO >= 2 THEN " +
                          "(SELECT NVL(SUM(PAGO_REAL),0) " +
				          "FROM CONTROL_PAGOS_DETALLE " +
                          "WHERE CDGEM = PRN.CDGEM " +
                          "AND CDGCL = PRC.CDGCL " +
				          "AND CDGCLNS = PRN.CDGNS " +
                          "AND CICLO = PRN.CICLO " +
                          "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,1) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,2) " +
                          "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,2)) < 7 AND PRN.PLAZO >= 2 THEN " +
                          "(SELECT NVL(SUM(PAGO_REAL),0) " +
				          "FROM CONTROL_PAGOS_DETALLE " +
                          "WHERE CDGEM = PRN.CDGEM " +
                          "AND CDGCL = PRC.CDGCL " +
				          "AND CDGCLNS = PRN.CDGNS " +
                          "AND CICLO = PRN.CICLO " +
                          "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,1) + 1 AND SYSDATE " +
                          "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_2 " +
		            ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,3) <= SYSDATE AND PRN.PLAZO >= 3 THEN " +
                          "(SELECT NVL(SUM(PAGO_REAL),0) " +
				          "FROM CONTROL_PAGOS_DETALLE " +
                          "WHERE CDGEM = PRN.CDGEM " +
                          "AND CDGCL = PRC.CDGCL " +
				          "AND CDGCLNS = PRN.CDGNS " +
                          "AND CICLO = PRN.CICLO " +
                          "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,2) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,3) " +
                          "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,3)) < 7 AND PRN.PLAZO >= 3 THEN " +
              	          "(SELECT NVL(SUM(PAGO_REAL),0) " +
				          "FROM CONTROL_PAGOS_DETALLE " +
                          "WHERE CDGEM = PRN.CDGEM " +
                          "AND CDGCL = PRC.CDGCL " +
				          "AND CDGCLNS = PRN.CDGNS " +
                          "AND CICLO = PRN.CICLO " +
                          "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,2) + 1 AND SYSDATE " +
                          "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_3 " +
  		            ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,4) <= SYSDATE AND PRN.PLAZO >= 4 THEN " +
                          "(SELECT NVL(SUM(PAGO_REAL),0) " +
				          "FROM CONTROL_PAGOS_DETALLE " +
                          "WHERE CDGEM = PRN.CDGEM " +
                          "AND CDGCL = PRC.CDGCL " +
				          "AND CDGCLNS = PRN.CDGNS " +
                          "AND CICLO = PRN.CICLO " +
                          "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,3) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,4) " +
                          "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,4)) < 7 AND PRN.PLAZO >= 4 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " + 
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,3) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_4 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,5) <= SYSDATE AND PRN.PLAZO >= 5 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,4) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,5) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,5)) < 7 AND PRN.PLAZO >= 5 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,4) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_5 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,6) <= SYSDATE AND PRN.PLAZO >= 6 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,5) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,6) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,6)) < 7 AND PRN.PLAZO >= 6 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,5) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_6 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,7) <= SYSDATE AND PRN.PLAZO >= 7 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,6) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,7) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,7)) < 7 AND PRN.PLAZO >= 7 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO  BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,6) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_7 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,8) <= SYSDATE AND PRN.PLAZO >= 8 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,7) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,8) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,8)) < 7 AND PRN.PLAZO >= 8 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,7) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_8 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,9) <= SYSDATE AND PRN.PLAZO >= 9 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,8) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,9) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,9)) < 7 AND PRN.PLAZO >= 9 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,8) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_9 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,10) <= SYSDATE AND PRN.PLAZO >= 10 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,9) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,10) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,10)) < 7 AND PRN.PLAZO >= 10 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,9) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_10 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,11) <= SYSDATE AND PRN.PLAZO >= 11 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,10) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,11) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,11)) < 7 AND PRN.PLAZO >= 11 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,10) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_11 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,12) <= SYSDATE AND PRN.PLAZO >= 12 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,11) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,12) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,12)) < 7 AND PRN.PLAZO >= 12 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,11) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_12 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,13) <= SYSDATE AND PRN.PLAZO >= 13 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,12) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,13) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,13)) < 7 AND PRN.PLAZO >= 13 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,12) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_13 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,14) <= SYSDATE AND PRN.PLAZO >= 14 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,13) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,14) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                    "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,14)) < 7 AND PRN.PLAZO >= 14 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,13) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_14 " +
                    ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,15) <= SYSDATE AND PRN.PLAZO >= 15 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,14) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,15) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,15)) < 7 AND PRN.PLAZO >= 15 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,14) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_15 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,16) <= SYSDATE AND PRN.PLAZO >= 16 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,15) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,16) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,16)) < 7 AND PRN.PLAZO >= 16 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,15) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_16 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,17) <= SYSDATE AND PRN.PLAZO >= 17 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,16) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,17) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,17)) < 7 AND PRN.PLAZO >= 17 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,16) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_17 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,18) <= SYSDATE AND PRN.PLAZO >= 18 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,17) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,18) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,18)) < 7 AND PRN.PLAZO >= 18 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,17) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_18 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,19) <= SYSDATE AND PRN.PLAZO >= 19 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,18) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,19) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,19)) < 7 AND PRN.PLAZO >= 19 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,18) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_19 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,20) <= SYSDATE AND PRN.PLAZO >= 20 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,19) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,20) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,20)) < 7 AND PRN.PLAZO >= 20 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,19) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_20 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,21) <= SYSDATE AND PRN.PLAZO >= 21 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " + 
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,20) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,21) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,21)) < 7 AND PRN.PLAZO >= 21 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,20) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_21 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,22) <= SYSDATE AND PRN.PLAZO >= 22 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,21) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,22) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,22)) < 7 AND PRN.PLAZO >= 22 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,21) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_22 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,23) <= SYSDATE AND PRN.PLAZO >= 23 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,22) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,23) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,23)) < 7 AND PRN.PLAZO >= 23 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,22) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_23 " +
                     ",CASE WHEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,24) <= SYSDATE AND PRN.PLAZO >= 24 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,23) + 1 AND fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,24) " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') " +
                     "WHEN (SYSDATE - fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,24)) < 7 AND PRN.PLAZO >= 24 THEN " +
                            "(SELECT NVL(SUM(PAGO_REAL),0) " +
				            "FROM CONTROL_PAGOS_DETALLE " +
                            "WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGCL = PRC.CDGCL " +
				            "AND CDGCLNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND FMOVIMIENTO BETWEEN fnFechaProxPago(PRN.INICIO,PRN.PERIODICIDAD,23) + 1 AND SYSDATE " +
                            "AND CLNS = 'G' AND ESTATUS <> 'CA') ELSE 0 END SEMANA_24 " +
                    "FROM PRC,PRN,CL  " +
                    "WHERE PRN.CDGEM = PRC.CDGEM  " +
       	            "AND PRN.CDGNS = PRC.CDGNS  " +
                    "AND PRN.CICLO = PRC.CICLO  " +
                    "AND PRC.CDGEM = CL.CDGEM  " +
                    "AND PRC.CDGCL = CL.CODIGO  " +
                    "AND PRN.CDGEM = '" + empresa + "'  " +
                    "AND PRN.CDGNS = '" + grupo + "' " +
                    "AND PRN.CICLO = '" + ciclo + "' " +
                    "AND PRC.CANTENTRE > 0  " +
                    "ORDER BY CL.PRIMAPE ASC";
            
            iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);

            if (dsGrupo.Tables[0].Rows.Count > 0 && dsAcred.Tables[0].Rows.Count > 0)
            {
                LlenaControlSemanalImp(dsGrupo, dsAcred);
            }
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    //METODO QUE GENERA INFORMACIÓN DEL ESTADO DE CUENTA GRUPAL
    private void getEstadoCuentaGrupal(string grupo, string ciclo)
    {
        string empresa = cdgEmpresa;
        string strFecha = string.Empty;
        int iRes;
        try
        {
            DataSet dsEnc = new DataSet();
            DataSet dsDet = new DataSet();

            //"TRUNC(SYSDATE) - 1";
            strFecha = "TRUNC(SYSDATE)";

            string query = "SELECT MP.CDGEM, " +
                           "NS.CODIGO CDGCLNS, " +
                           "NS.NOMBRE NOMCLNS, " +
                           "PRN.CDGOCPE, " +
                           "NS.CALLE, " +
                           "RG.CODIGO CDGRG, " +
                           "RG.NOMBRE NOMRG, " +
                           "CO.CODIGO CDGCO, " +
                           "CO.NOMBRE NOMCO, " +
                           "PE.NOMBRE1, " +
                           "PE.NOMBRE2, " +
                           "PE.PRIMAPE, " +
                           "PE.SEGAPE, " +
                           "LO.NOMBRE NOMLO, " +
                           "PRN.CICLO, " +
                           "PRN.INICIO, " +
                           "FNFECHAPROXPAGO(PRN.INICIO,PRN.PERIODICIDAD,PRN.PLAZO) FIN , " +
                           "PRN.PERIODICIDAD, " +
                           "PRN.CANTAUTOR, " +
                           "PRN.CANTENTRE, " +
                           "CASE WHEN (CDGTPC = '01') THEN " +
                                "PRN.CANTENTRE * .1 " +
                           "ELSE " +
                                "0 " +
                           "END GARANTIA, " +
                           "PRN.SITUACION, " +
                           "CASE WHEN (SELECT SITUACION " +
                                      "FROM TBL_CIERRE_DIA " +
                                      "WHERE CDGEM = MP.CDGEM " +
                                      "AND CDGCLNS = MP.CDGCLNS " +
                                      "AND CLNS = MP.CLNS " +
                                      "AND CICLO = MP.CICLO " +
                                      "AND FECHA_CALC = " + strFecha + ") IS NOT NULL THEN " +
                                            "(SELECT SITUACION " +
                                            "FROM TBL_CIERRE_DIA " +
                                            "WHERE CDGEM = MP.CDGEM " +
                                            "AND CDGCLNS = MP.CDGCLNS " +
                                            "AND CLNS = MP.CLNS " +
                                            "AND CICLO = MP.CICLO " +
                                            "AND FECHA_CALC = " + strFecha + ") " +
                                 "ELSE PRN.SITUACION " +
                           "END SITUA, " +
                           "PRN.TASA, " +
                           "PRN.PLAZO, " +
                           "COL.NOMBRE NOMCOL, " +
                           "MU.NOMBRE NOMMU, " +
                           "MP.TIPO, " +
                           "MP.CLNS, " +
                           "MP.CANTIDAD, " +
                           strFecha + " FECHA, " +
                           "CASE WHEN (SELECT COUNT(*) " +
                                      "FROM GRACIAPRN " +
                                      "WHERE CDGEM = PRN.CDGEM " +
                                      "AND CDGCLNS = PRN.CDGNS " +
                                      "AND CICLO = PRN.CICLO " +
                                      "AND CLNS = 'G' " +
                                      "AND INICIOGR <= " + strFecha + " " +
                                      "AND FINGR >= " + strFecha + ") = 0 THEN " + strFecha + " " +
                           "ELSE " + 
                               "(SELECT INICIOGR " +
                               "FROM GRACIAPRN " +
                               "WHERE CDGEM = PRN.CDGEM " +
                               "AND CDGCLNS = PRN.CDGNS " +
                               "AND CICLO = PRN.CICLO " +
                               "AND CLNS = 'G' " +
                               "AND INICIOGR <= " + strFecha + " " +
                               "AND FINGR >= " + strFecha + ") " +
                           "END FECHACALC, " +
                           "(SELECT NVL(SUM(CANTIDAD),0) " +
                              "FROM PAG_GAR_SIM " +
                              "WHERE CDGEM = NS.CDGEM " +
                              "AND CDGCLNS = NS.CODIGO " +
                              "AND ESTATUS = 'RE' " +
                            //"AND FPAGO <= " + strFecha + "                                                  //INSTRUCCIONES UTILIZADAS PARA EDO CUENTA HISTORICOS
                              "AND CLNS = 'G') TOTALABONOS, " +
                           "(SELECT NVL(SUM(CANTIDAD),0) " +
                              "FROM PAG_GAR_SIM " +
                              "WHERE CDGEM = NS.CDGEM " +
                              "AND CDGCLNS = NS.CODIGO " +
                              "AND ESTATUS NOT IN ('RE','CA') " +
                            //"AND FPAGO <= " + strFecha + "                                                  //INSTRUCCIONES UTILIZADAS PARA EDO CUENTA HISTORICOS
                              "AND CLNS = 'G') TOTALCARGOS, " +
                            //"FNSDOGARANTIA(PRN.CDGEM,PRN.CDGNS,PRN.CICLO,'G'," + strFecha + ") SALDOGL, " +    //INSTRUCCIONES UTILIZADAS PARA EDO CUENTA HISTORICOS
                           "FNSDOGARANTIA(PRN.CDGEM,PRN.CDGNS,PRN.CICLO,'G') SALDOGL, " +
                           "CASE WHEN (SELECT COUNT(*) " +
                                      "FROM TBL_CIERRE_DIA " +
                                      "WHERE CDGEM = MP.CDGEM " +
                                      "AND CDGCLNS = MP.CDGCLNS " +
                                      "AND CLNS = MP.CLNS " +
                                      "AND CICLO = MP.CICLO " +
                                      "AND FECHA_CALC = " + strFecha + ") = 1 THEN " +
                                          "(SELECT NVL(SUM(SDO_RECARGOS),0) " +
                                          "FROM TBL_CIERRE_DIA " +
                                          "WHERE CDGEM = MP.CDGEM " +
                                          "AND CDGCLNS = MP.CDGCLNS " +
                                          "AND CLNS = MP.CLNS " +
                                          "AND CICLO = MP.CICLO " +
                                          "AND FECHA_CALC = " + strFecha + ")  " +
                                "ELSE " +
                                      "(SELECT NVL(SUM(SDO_RECARGOS),0) " +
                                      "FROM TBL_CIERRE_DIA " +
                                      "WHERE CDGEM = MP.CDGEM " +
                                      "AND CDGCLNS = MP.CDGCLNS " +
                                      "AND CLNS = MP.CLNS " +
                                      "AND CICLO = MP.CICLO " +
                                      "AND FECHA_CALC = " + strFecha + " - 1)  " +
                           "END SALDOREC, " +
                           "CASE WHEN (SELECT COUNT(*) " +
                                      "FROM TBL_CIERRE_DIA " +
                                      "WHERE CDGEM = MP.CDGEM " +
                                      "AND CDGCLNS = MP.CDGCLNS " +
                                      "AND CLNS = MP.CLNS " +
                                      "AND CICLO = MP.CICLO " +
                                      "AND FECHA_CALC = " + strFecha + ") = 1 THEN " +
                                          "(SELECT NVL(DIAS_MORA,0) " +
                                          "FROM TBL_CIERRE_DIA " +
                                          "WHERE CDGEM = MP.CDGEM " +
                                          "AND CDGCLNS = MP.CDGCLNS " +
                                          "AND CLNS = MP.CLNS " +
                                          "AND CICLO = MP.CICLO " +
                                          "AND FECHA_CALC = " + strFecha + ") " +
                                "ELSE " +
                                      "(SELECT NVL(DIAS_MORA,0) " +
                                      "FROM TBL_CIERRE_DIA " +
                                      "WHERE CDGEM = MP.CDGEM " +
                                      "AND CDGCLNS = MP.CDGCLNS " +
                                      "AND CLNS = MP.CLNS " +
                                      "AND CICLO = MP.CICLO " +
                                      "AND FECHA_CALC = " + strFecha + " - 1) " +
                           "END DIASMORA, " +
                           "(SELECT (CF.IVA/CF.PORCENTAJE) " +
                            "FROM CF,PRN " +
                            "WHERE PRN.CDGEM = CF.CDGEM " +
                            "AND PRN.CDGFDI = CF.CDGFDI " +
                            "AND PRN.CDGEM =  MP.CDGEM " +
                            "AND PRN.CDGNS =  MP.CDGCLNS " +
                            "AND PRN.CICLO = MP.CICLO) IVA, " +
                           "(SELECT (IVA/PORCENTAJE) " +
                              "FROM CF " +
                              "WHERE CDGEM = PRN.CDGEM " +
                              "AND CDGFDI = PRN.CDGFDI) TASAANUAL, " +
                          "(SELECT SUM(CANTIDAD)" +
                            "FROM MP WHERE CDGEM = PRN.CDGEM " +
                            "AND CDGNS = PRN.CDGNS " +
                            "AND CICLO = PRN.CICLO " +
                            "AND TIPO = 'IN' " +
                            "AND ESTATUS <> 'E') INTERES, " +
                          "FNCALTOTALSINIVA(PRN.CDGEM,PRN.CDGNS,'G',PRN.CICLO) TOTALSINIVA, " +
                          "DECODE(PRN.PERIODICIDAD,'S','SEMANAL' " +
                                                  ",'C','BISEMANAL' " +
                                                  ",'Q','QUINCENAL' " +
                                                  ",'M','MENSUAL') NOMPER, " +
                          "FNCALDIASATRASO(PRN.CDGEM,PRN.CDGNS,LPAD(PRN.CICLO,2,'0'),'G') DIASATRASO " +
                          "FROM MP left join prn on prn.cdgem = mp.cdgem and prn.cdgns = mp.cdgns and prn.ciclo = mp.ciclo " +
                                  "left join ns on ns.cdgem = prn.cdgem and ns.codigo = prn.cdgns " +
                                  "left join co on co.codigo = prn.cdgco and co.cdgem = prn.cdgem " +
                                  "left join rg on co.cdgrg = rg.codigo and co.cdgem = rg.cdgem " +
                                  "left join pai on pai.codigo = ns.cdgpai " +
                                  "left join ef on ef.cdgpai = ns.cdgpai and ef.codigo = ns.cdgef " +
                                  "left join mu on mu.cdgpai = ns.cdgpai and mu.cdgef = ns.cdgef and mu.codigo = ns.cdgmu " +
                                  "left join lo on lo.cdgpai = ns.cdgpai and lo.cdgef = ns.cdgef and lo.cdgmu = ns.cdgmu and lo.codigo = ns.cdglo " +
                                  "left join col on col.cdgpai = ns.cdgpai and col.cdgef = ns.cdgef and col.cdgmu = ns.cdgmu and col.cdglo = ns.cdglo and col.codigo = ns.cdgcol " +
                                  "left join pe on pe.cdgem = prn.cdgem and pe.codigo = prn.cdgocpe " +
                          "WHERE MP.CDGEM = '" + empresa + "' " +
                          "AND MP.CDGNS = '" + grupo + "' " +
                          "AND MP.CICLO = '" + ciclo + "' " +
                          "AND MP.CLNS = 'G' " +
                          "AND MP.ESTATUS <> 'E' " +
                          "AND MP.TIPO = 'IN' " +
                          "ORDER BY MP.FREALDEP ASC, MP.PERIODO ASC, TO_NUMBER(MP.SECUENCIA), MP.TIPO";

            iRes = db.ExecuteDS(ref dsEnc, query, CommandType.Text);

            query = "SELECT MP.CDGEM, " +
                    "MP.CLNS, " +
                    "MP.CDGCLNS, " +
                    "MP.CDGCL, " +
                    "MP.PERIODO, " +
                    "MP.TIPO, " +
                    "MP.FREALDEP, " +
                    "MP.CANTIDAD, " +
                    "MP.PAGADOCAP, " +
                    "MP.PAGADOINT, " +
                    "MP.PAGADOREC, " +
                    "NVL((SELECT CASE WHEN FDEPOSITO IS NOT NULL THEN " +
                         "' (' || TO_CHAR(FDEPOSITO,'DD/MM/YYYY') || ')' " +
                         "ELSE '' " +
                         "END " +
                         "FROM PDI " +
                         "WHERE CDGEM = MP.CDGEM " +
                         "AND CDGCLNS = MP.CDGCLNS " +
                         "AND CLNS = MP.CLNS " +
                         "AND CDGCB = MP.CDGCB " +
                         "AND SECUENCIAIM = MP.SECUENCIA " +
                         "AND CANTIDAD = MP.CANTIDAD " +
                         "AND FECHAIM = MP.FDEPOSITO),'') PAGOIDEN, " +
                    "(SELECT (IVA/PORCENTAJE) " +
                    "FROM CF " +
                    "WHERE PRN.CDGEM = CF.CDGEM " +
                    "AND PRN.CDGFDI = CF.CDGFDI) IVA " +
                    "FROM MP, PRN " +
                    "WHERE MP.CDGEM = '" + empresa + "' " +
                    "AND MP.CDGNS = '" + grupo + "' " +
                    "AND MP.CICLO = '" + ciclo + "' " +
                    "AND MP.CLNS = 'G' " +
                    "AND MP.ESTATUS <> 'E' " +
                    "AND PRN.CDGEM = MP.CDGEM " +
                    "AND PRN.CDGNS = MP.CDGCLNS " +
                    "AND PRN.CICLO = MP.CICLO " +
                    //"AND TRUNC(MP.FREALDEP) <= " + strFecha + " " +
                    "ORDER BY MP.FREALDEP ASC, MP.PERIODO ASC, TO_NUMBER(MP.SECUENCIA), MP.TIPO";

            iRes = db.ExecuteDS(ref dsDet, query, CommandType.Text);

            if (dsEnc.Tables[0].Rows.Count > 0)
                LlenaEstadoCuenta(dsEnc, dsDet);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getFormatoControl(string id, string grupo, string ciclo)
    {
        string empresa = cdgEmpresa;
        int iRes;
        try
        {
            DataSet dsGrupo = new DataSet();
            DataSet dsAcred = new DataSet();
            DataSet dsDoc = new DataSet();
            string query = string.Empty;

            switch (id)
            {
                case "5":
                case "6":
                case "7":
                case "10":
                    query = "SELECT PRN.*, " +
                            "DECODE(NVL(PRN.PERIODICIDAD,''), " +
                                           "'S', PRN.INICIO + (7 * NVL(PRN.PLAZO,0)), " +
                                           "'Q', PRN.INICIO + (15 * NVL(PRN.PLAZO,0)), " +
                                           "'C', PRN.INICIO + (14 * NVL(PRN.PLAZO,0)), " +
                                           "'M', PRN.INICIO + (30 * NVL(PRN.PLAZO,0)), " +
                                           "'', '') FECHAFIN, " +
                                   "NS.NOMBRE GRUPO, " +
                                   "NOMBREC(NULL,NULL,'I','N',A.NOMBRE1,A.NOMBRE2,A.PRIMAPE,A.SEGAPE) NOMBREGERENTE, " +
                                   "NOMBREC(NULL,NULL,'I','N',B.NOMBRE1,B.NOMBRE2,B.PRIMAPE,B.SEGAPE) NOMBRECOORDINADOR, " +
                                   "NOMBREC(NULL,NULL,'I','N',PE.NOMBRE1,PE.NOMBRE2,PE.PRIMAPE,PE.SEGAPE) NOMBREASESOR, " +
                                   "CO.NOMBRE COORDINACION " +
                                   "FROM PRN, NS, CO, PE, PE A, PE B " +
                                   "WHERE PRN.CDGEM = '" + empresa + "' " +
                                   "AND PRN.CDGNS = '" + grupo + "' " +
                                   "AND PRN.CICLO = '" + ciclo + "' " +
                                   "AND NS.CDGEM = PRN.CDGEM " +
                                   "AND NS.CODIGO = PRN.CDGNS " +
                                   "AND CO.CDGEM = NS.CDGEM " +
                                   "AND CO.CODIGO = NS.CDGCO " +
                                   "AND PRN.CDGEM = PE.CDGEM " +
                                   "AND PRN.CDGOCPE = PE.CODIGO " +
                                   "AND CO.CDGEM = A.CDGEM " +
                                   "AND CO.CDGPE = A.CODIGO " +
                                   "AND PE.CDGEM = B.CDGEM " +
                                   "AND PE.CALLE = B.TELEFONO";
                    break;
                case "24":
                    query = "SELECT SN.*, " +
                            "DECODE(NVL(SN.PERIODICIDAD,''), " +
                                           "'S', SN.INICIO + (7 * NVL(SN.DURACION,0)), " +
                                           "'Q', SN.INICIO + (15 * NVL(SN.DURACION,0)), " +
                                           "'C', SN.INICIO + (14 * NVL(SN.DURACION,0)), " +
                                           "'M', SN.INICIO + (30 * NVL(SN.DURACION,0)), " +
                                           "'', '') FECHAFIN, " +
                                   "NS.NOMBRE GRUPO, " +
                                   "NOMBREC(NULL,NULL,'I','N',A.NOMBRE1,A.NOMBRE2,A.PRIMAPE,A.SEGAPE) NOMBREGERENTE, " +
                                   "NOMBREC(NULL,NULL,'I','N',B.NOMBRE1,B.NOMBRE2,B.PRIMAPE,B.SEGAPE) NOMBRECOORDINADOR, " +
                                   "NOMBREC(NULL,NULL,'I','N',PE.NOMBRE1,PE.NOMBRE2,PE.PRIMAPE,PE.SEGAPE) NOMBREASESOR, " +
                                   "CO.NOMBRE COORDINACION " +
                                   "FROM SN, NS, CO, PE, PE A, PE B " +
                                   "WHERE SN.CDGEM = '" + empresa + "' " +
                                   "AND SN.CDGNS = '" + grupo + "' " +
                                   "AND SN.CICLO = '" + ciclo + "' " +
                                   "AND NS.CDGEM = SN.CDGEM " +
                                   "AND NS.CODIGO = SN.CDGNS " +
                                   "AND CO.CDGEM = NS.CDGEM " +
                                   "AND CO.CODIGO = NS.CDGCO " +
                                   "AND SN.CDGEM = PE.CDGEM " +
                                   "AND SN.CDGOCPE = PE.CODIGO " +
                                   "AND CO.CDGEM = A.CDGEM " +
                                   "AND CO.CDGPE = A.CODIGO " +
                                   "AND PE.CDGEM = B.CDGEM " +
                                   "AND PE.CALLE = B.TELEFONO";
                    break;
            }
            

            iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);

            //Control semanal
            if (id == "5" || id == "10")
            {
                query = "SELECT D.ARCHIVO " +
                        "FROM PRN, DOCTPC D " +
                        "WHERE PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND D.CDGEM = PRN.CDGEM " +
                        "AND D.CDGTPC = PRN.CDGTPC " +
                        "AND D.FORMATO = 'CTRLSEMANAL'";

                iRes = db.ExecuteDS(ref dsDoc, query, CommandType.Text);

                query = "SELECT CASE WHEN PRC.CDGCL = PRN.PRESIDENTE THEN 'P' " +
                                    "WHEN PRC.CDGCL = PRN.TESORERO THEN 'T' " +
                                    "WHEN PRC.CDGCL = PRN.SECRETARIO THEN 'SE' " +
                                "END PUESTO, " +
                        "NOMBREC(CL.CDGEM,CL.CODIGO,'I','A','','','','') CLIENTE, " +
                        "round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(4 * 100), " +
                                "'Q', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0) * 15)/(30 * 100), " +
                                "'C', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(2 * 100), " +
                                "'M', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(100), " +
                                "'',  ''),0) as TOTAL_A_PAGAR, " +
                        "PRC.CANTENTRE, " +
                        "round((round(decode(nvl(PRN.periodicidad,''), 'S', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(4 * 100), " +
                                "'Q', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0) * 15)/(30 * 100), " +
                                "'C', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(2 * 100), " +
                                "'M', nvl(PRC.cantentre,0) + (nvl(PRN.tasa,0) * nvl(PRN.plazo,0) * nvl(PRC.cantentre,0))/(100), " +
                                "'',  ''),0)) / PRN.PLAZO,2) AS PARCIALIDAD " +
                        "FROM PRC,PRN,CL " +
                        "WHERE PRN.CDGEM = PRC.CDGEM " +
                        "AND PRN.CDGNS = PRC.CDGNS " +
                        "AND PRN.CICLO = PRC.CICLO " +
                        "AND PRC.CDGEM = CL.CDGEM " +
                        "AND PRC.CDGCL = CL.CODIGO " +
                        "AND PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND PRC.CANTENTRE > 0 " +
                        "ORDER BY CL.PRIMAPE ASC";
            }
            //Control de Ahorros
            else if (id == "6")
            {
                query = "SELECT D.ARCHIVO " +
                        "FROM PRN, DOCTPC D " +
                        "WHERE PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND D.CDGEM = PRN.CDGEM " +
                        "AND D.CDGTPC = PRN.CDGTPC " +
                        "AND D.FORMATO = 'CTRLAHORRO'";

                iRes = db.ExecuteDS(ref dsDoc, query, CommandType.Text);

                query = "SELECT NOMBREC(CL.CDGEM,CL.CODIGO,'I','A','','','','') CLIENTE, " +
                        "CASE WHEN PRN.CDGTPC = '01' THEN " +
                            "(PRC.CANTENTRE * .1) " +
                        "ELSE " + 
                            "0 " +     
                        "END GARANTIA " +
                        "FROM PRC,PRN,CL " +
                        "WHERE PRN.CDGEM = PRC.CDGEM " +
                        "AND PRN.CDGNS = PRC.CDGNS " +
                        "AND PRN.CICLO = PRC.CICLO " +
                        "AND PRC.CDGEM = CL.CDGEM " +
                        "AND PRC.CDGCL = CL.CODIGO " +
                        "AND PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +  
                        "AND PRC.CANTENTRE > 0 " +
                        "ORDER BY CL.PRIMAPE ASC";
            }
            //Control de Asistencia
            else if (id == "7")
            {
                query = "SELECT D.ARCHIVO " +
                        "FROM PRN, DOCTPC D " +
                        "WHERE PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND D.CDGEM = PRN.CDGEM " +
                        "AND D.CDGTPC = PRN.CDGTPC " +
                        "AND D.FORMATO = 'CTRLASIST'";

                iRes = db.ExecuteDS(ref dsDoc, query, CommandType.Text);

                query = "SELECT NOMBREC(CL.CDGEM,CL.CODIGO,'I','A','','','','') CLIENTE " +
                        "FROM PRC,PRN,CL " +
                        "WHERE PRN.CDGEM = PRC.CDGEM " +
                        "AND PRN.CDGNS = PRC.CDGNS " +
                        "AND PRN.CICLO = PRC.CICLO " +
                        "AND PRC.CDGEM = CL.CDGEM " +
                        "AND PRC.CDGCL = CL.CODIGO " +
                        "AND PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND PRC.CANTENTRE > 0 " +
                        "ORDER BY CL.PRIMAPE ASC";
            }

            //Autorizacion de Pago con Garantia
            else if (id == "24")
            {
                query = "SELECT D.ARCHIVO " +
                        "FROM PRN, DOCTPC D " +
                        "WHERE PRN.CDGEM = '" + empresa + "' " +
                        "AND PRN.CDGNS = '" + grupo + "' " +
                        "AND PRN.CICLO = '" + ciclo + "' " +
                        "AND D.CDGEM = PRN.CDGEM " +
                        "AND D.CDGTPC = PRN.CDGTPC " +
                        "AND D.FORMATO = 'AUTPAGOGL'";

                iRes = db.ExecuteDS(ref dsDoc, query, CommandType.Text);

                query = "SELECT NOMBREC(CL.CDGEM,CL.CODIGO,'I','A','','','','') CLIENTE " +
                        "FROM SC,SN,CL " +
                        "WHERE SN.CDGEM = SC.CDGEM " +
                        "AND SN.CDGNS = SC.CDGNS " +
                        "AND SN.CICLO = SC.CICLO " +
                        "AND SC.CDGEM = CL.CDGEM " +
                        "AND SC.CDGCL = CL.CODIGO " +
                        "AND SN.CDGEM = '" + empresa + "' " +
                        "AND SN.CDGNS = '" + grupo + "' " +
                        "AND SN.CICLO = '" + ciclo + "' " +
                        "AND SC.SITUACION = 'A' " +
                        "ORDER BY CL.PRIMAPE ASC";
            }

            iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);

            if (dsGrupo.Tables[0].Rows.Count > 0 && dsAcred.Tables[0].Rows.Count > 0 && dsDoc.Tables[0].Rows.Count > 0)
            {
                if (id == "5" || id == "10")
                    LlenaControlSemanal(dsGrupo, dsAcred, dsDoc, id);
                else if (id == "6")
                    LlenaControlAhorro(dsGrupo, dsAcred, dsDoc);
                else if (id == "7")
                    LlenaControlAsistencia(dsGrupo, dsAcred, dsDoc);
                else if (id == "24")
                    LlenaAutPagoGL(dsGrupo, dsAcred, dsDoc);
            }
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
            
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getImpresionCheques(string grupo, string ciclo, string cuenta, string coord, string chqIni, string chqFin)
    {
        string empresa = cdgEmpresa;
        string strCheque = string.Empty;
        string strCuenta = string.Empty;
        string strCoord = string.Empty;
        int iRes;
        try
        {
            DataSet dsAcred = new DataSet();

            if (chqIni != null && chqFin != null)
            {
                strCheque = "AND TO_NUMBER(PRC.NOCHEQUE) BETWEEN TO_NUMBER('" + chqIni + "') AND TO_NUMBER('" + chqFin + "') ";
            }
            if (coord != null)
            {
                strCoord = "AND PRN.CDGCO = '" + coord + "' ";
            }
            if (cuenta != null)
            {
                strCuenta = "AND PRC.CDGCB = '" + cuenta + "' ";
            }

            string query = "SELECT PRN.CDGCO, " +
                           "PRN.CDGNS, " +
                           "PRN.CICLO, " +
                           "PRC.CDGCL, " +
                           "PRC.ENTRREAL MONTO, " +
                           "PRC.NOCHEQUE, " +
                           "PRC.CDGCB, " +
                           "NOMBREC(CL.CDGEM,CL.CODIGO,'I','A','','','','') NOMBREC, " +
                           "TO_CHAR(PRN.INICIO, 'DD/MM/YYYY') FECHA, " +
                           "CASE WHEN EXISTS (SELECT CDGCL FROM CHEQUE_NONEG NN " +
                                             "WHERE NN.CDGEM = PRC.CDGEM " +
                                             "AND NN.CDGCLNS = PRC.CDGNS " +
                                             "AND NN.CICLO = PRC.CICLO " +
                                             "AND NN.CDGCL = PRC.CDGCL " +
                                             "AND NN.TIPO = 'DE') THEN '' " +
                           "ELSE " +
                                "'NO NEGOCIABLE' " +
                           "END NO_NEGOCIABLE " +
                           "FROM PRN, PRC, CL " +
                           "WHERE PRN.CDGEM = '" + empresa + "' " +
                           "AND PRN.CDGNS = '" + grupo + "' " +
                           "AND PRN.CICLO = '" + ciclo + "' " +
                           strCoord +
                           "AND PRC.CDGEM = PRN.CDGEM " +
                           "AND PRC.CDGNS = PRN.CDGNS " +
                           "AND PRC.CICLO = PRN.CICLO " +
                           strCuenta +
                           strCheque +
                           "AND CL.CDGEM = PRC.CDGEM " +
                           "AND CL.CODIGO = PRC.CDGCL " +
                           "ORDER BY PRC.NOCHEQUE";

            iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);

            if (dsAcred.Tables[0].Rows.Count > 0)
                LlenaImpresionCheques(dsAcred, cuenta);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getPagare(string grupo, string ciclo, string usuario)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        string status = string.Empty;
        int iRes;
        try
        {
            DataSet dsDet = new DataSet();
            DataSet dsAmort = new DataSet();

            query = "SELECT TIPO_DOC, " +
                    "NOMBRE_EMP, " +
                    "CODIGO_EMP, " +
                    "CODIGO_CTE, " +
                    "NOMBRE_CTE, " +
                    "APELLIDOS_CTE, " +
                    "CODIGO_GPO, " +
                    "NOMBRE_GPO, " +
                    "CALLE_CTE, " +
                    "COLONIA_CTE, " +
                    "LOCALIDAD_CTE, " +
                    "MUNICIPIO_CTE, " +
                    "ENTIDAD_CTE, " +
                    "PAIS_CTE, " +
                    "CP_CTE, " +
                    "FECHA_INICIO, " +
                    "FECHA_FIN, " +
                    "CANTIDAD_LETRA, " +
                    "CANTIDAD_NUMERO, " +
                    "TASA_CTE, " +
                    "NOMBRE_AVAL1, " +
                    "DIREC1_AVAL1, " +
                    "NOMBRE_AVAL2, " +
                    "DIREC1_AVAL2, " +
                    "NOMBRE_AVAL3, " +
                    "DIREC1_AVAL3, " +
                    "NOMBRE_AVAL4, " +
                    "DIREC1_AVAL4, " +
                    "DIREC2_AVAL1, " +
                    "DIREC2_AVAL2, " + 
                    "DIREC2_AVAL3, " +
                    "DIREC2_AVAL4, " +
                    "NO_AVALES, " +
                    "NOMBRE_COM_CTE, " +
                    "CTAS_BANCOS, " +
                    "PLAZO_CTE, " +
                    "PERIODO_CTE, " +
                    "CANT_ENT_GPO, " +
                    "CANT_ENT_GPO_LETRA, " +
                    "CTAS_BANCOS, " +
                    "DIRECCION_SUC, " +
                    "LOCALIDAD_SUC, " +
                    "MUNICIPIO_SUC, " +
                    "PAIS_SUC, " +
                    "INTERESES_GPO, " +
                    "COD_PRE_GPO, " +
                    "COD_SEC_GPO, " +
                    "COD_TES_GPO, " +
                    "NOM_PRE_GPO, " +
                    "NOM_SEC_GPO, " +
                    "NOM_TES_GPO, " +
                    "DIR_PRE_GPO, " +
                    "DIR_SEC_GPO, " +
                    "DIR_TES_GPO, " +
                    "(SELECT (CF.IVA/CF.PORCENTAJE) " +
                    "FROM CF,PRN " +
                    "WHERE PRN.CDGEM = '" + empresa + "' " +
                    "AND PRN.CDGNS = '" + grupo + "' " +
                    "AND PRN.CICLO = '" + ciclo + "' " +
                    "AND PRN.CDGEM = CF.CDGEM " +
                    "AND PRN.CDGFDI = CF.CDGFDI) IVA, " +
                    "TIPO_PROD, " +
                    "NOM_PROD, " +
                    "(SELECT CALLE || ', Col. ' || INITCAP(COL.NOMBRE) || ', ' || MU.NOMBRE || ', ' || EF.NOM_ACEN || ', CP ' || COL.CDGPOSTAL " +
                     "FROM EM, EF, MU, COL " +
                     "WHERE EM.CODIGO = '" + empresa + "' " +
                     "AND EF.CODIGO = EM.CDGEF " +
                     "AND MU.CDGEF = EM.CDGEF " +
                     "AND MU.CODIGO = EM.CDGMU " +
                     "AND COL.CDGEF = EM.CDGEF " +
                     "AND COL.CDGMU = EM.CDGMU " +
                     "AND COL.CDGLO = EM.CDGLO " +
                     "AND COL.CODIGO = EM.CDGCOL) DIRECCION_EMP " +
                    "FROM IMPPAG " +
                    "WHERE CODIGO_EMP = '" + empresa + "' " +
                    "AND CODIGO_GPO = '" + grupo + "' " +
                    "AND CDGPE = '" + usuario + "' " +
                    "AND TIPO_DOC = 'PAGARE' " +
                    "ORDER BY TO_NUMBER(CANTIDAD_NUMERO) DESC, NOMBRE_CTE ";

            iRes = db.ExecuteDS(ref dsDet, query, CommandType.Text);

            query = "SELECT * " +
                    "FROM TAFIN " +
                    "WHERE CDG_EM = '" + empresa + "' " +
                    "AND CDG_GRUPO = '" + grupo + "' " +
                    "AND CDGPE = '" + usuario + "' " +
                    "AND TIPO_DOC = 'PAGARE' " +
                    "ORDER BY CDGCL, NOPAGO";

            iRes = db.ExecuteDS(ref dsAmort, query, CommandType.Text);

            if (dsDet.Tables[0].Rows.Count > 0 && dsAmort.Tables[0].Rows.Count > 0)
                LlenaPagare(dsDet, dsAmort);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getPagareGrupal(string grupo, string ciclo, string usuario)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        string status = string.Empty;
        int iRes;
        try
        {
            DataSet dsDet = new DataSet();
            DataSet dsAval = new DataSet();

            query = "SELECT TIPO_DOC, " +
                    "NOMBRE_EMP, " +
                    "CODIGO_EMP, " +
                    "CODIGO_CTE, " +
                    "NOMBRE_CTE, " +
                    "APELLIDOS_CTE, " +
                    "CODIGO_GPO, " +
                    "NOMBRE_GPO, " +
                    "CALLE_CTE, " +
                    "COLONIA_CTE, " +
                    "LOCALIDAD_CTE, " +
                    "MUNICIPIO_CTE, " +
                    "ENTIDAD_CTE, " +
                    "PAIS_CTE, " +
                    "CP_CTE, " +
                    "FECHA_INICIO, " +
                    "FECHA_FIN, " +
                    "CANTIDAD_LETRA, " +
                    "CANTIDAD_NUMERO, " +
                    "TASA_CTE, " +
                    "NOMBRE_AVAL1, " +
                    "DIREC1_AVAL1, " +
                    "NOMBRE_AVAL2, " +
                    "DIREC1_AVAL2, " +
                    "NOMBRE_AVAL3, " +
                    "DIREC1_AVAL3, " +
                    "NOMBRE_AVAL4, " +
                    "DIREC1_AVAL4, " +
                    "DIREC2_AVAL1, " +
                    "DIREC2_AVAL2, " +
                    "DIREC2_AVAL3, " +
                    "DIREC2_AVAL4, " +
                    "NO_AVALES, " +
                    "NOMBRE_COM_CTE, " +
                    "CTAS_BANCOS, " +
                    "PLAZO_CTE, " +
                    "PERIODO_CTE, " +
                    "CANT_ENT_GPO, " +
                    "CANT_ENT_GPO_LETRA, " +
                    "CTAS_BANCOS, " +
                    "DIRECCION_SUC, " +
                    "LOCALIDAD_SUC, " +
                    "MUNICIPIO_SUC, " +
                    "PAIS_SUC, " +
                    "INTERESES_GPO, " +
                    "COD_PRE_GPO, " +
                    "COD_SEC_GPO, " +
                    "COD_TES_GPO, " +
                    "NOM_PRE_GPO, " +
                    "NOM_SEC_GPO, " +
                    "NOM_TES_GPO, " +
                    "DIR_PRE_GPO, " +
                    "DIR_SEC_GPO, " +
                    "DIR_TES_GPO, " +
                    "(SELECT (CF.IVA/CF.PORCENTAJE) " +
                    "FROM CF,PRN " +
                    "WHERE PRN.CDGEM = '" + empresa + "' " +
                    "AND PRN.CDGNS = '" + grupo + "' " +
                    "AND PRN.CICLO = '" + ciclo + "' " +
                    "AND PRN.CDGEM = CF.CDGEM " +
                    "AND PRN.CDGFDI = CF.CDGFDI) IVA, " +
                    "TIPO_PROD, " +
                    "NOM_PROD, " +
                    "(SELECT CALLE || ', Col. ' || INITCAP(COL.NOMBRE) || ', ' || MU.NOMBRE || ', ' || EF.NOM_ACEN || ', CP ' || COL.CDGPOSTAL " +
                     "FROM EM, EF, MU, COL " +
                     "WHERE EM.CODIGO = '" + empresa + "' " +
                     "AND EF.CODIGO = EM.CDGEF " +
                     "AND MU.CDGEF = EM.CDGEF " +
                     "AND MU.CODIGO = EM.CDGMU " +
                     "AND COL.CDGEF = EM.CDGEF " +
                     "AND COL.CDGMU = EM.CDGMU " +
                     "AND COL.CDGLO = EM.CDGLO " +
                     "AND COL.CODIGO = EM.CDGCOL) DIRECCION_EMP " +
                    "FROM IMPPAG " +
                    "WHERE CODIGO_EMP = '" + empresa + "' " +
                    "AND CODIGO_GPO = '" + grupo + "' " +
                    "AND CDGPE = '" + usuario + "' " +
                    "AND TIPO_DOC = 'PAGARE' " +
                    "ORDER BY TO_NUMBER(CANTIDAD_NUMERO) DESC, NOMBRE_CTE ";

            iRes = db.ExecuteDS(ref dsDet, query, CommandType.Text);

            query = "SELECT * " +
                    "FROM REP_AVAL_PAGARE " +
                    "WHERE CDGEM = '" + empresa + "' " +
                    "AND CDGCLNS = '" + grupo + "' " +
                    "AND CDGPE = '" + usuario + "' " +
                    "ORDER BY NOMBRE";

            iRes = db.ExecuteDS(ref dsAval, query, CommandType.Text);

            if (dsDet.Tables[0].Rows.Count > 0 && dsAval.Tables[0].Rows.Count > 0)
                LlenaPagareGrupal(dsDet, dsAval);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getReimpCheques(string tipo, string codigo, string ciclo, string clns, string cuenta, string chqIni)
    {
        string empresa = cdgEmpresa;
        string[] strChq;
        string cheque = string.Empty;
        string query = string.Empty;
        int iRes;
        int i;
        try
        {
            DataSet dsAcred = new DataSet();
            
            strChq = chqIni.Split('|');
            for (i = 0; i < strChq.Length; i++)
            {
                cheque += (cheque != "" ? "," : "") + "'" + strChq[i] + "'";
            }

            if (tipo == "D")
            {
                query = "SELECT PRC.CDGNS " +
                        ",PRC.CICLO " +
                        ",NOMBREC(CL.CDGEM,CL.CODIGO,'I','A','','','','') NOMBREC " +
                        ",PRC.ENTRREAL MONTO " +
                        ",PRC.CDGCB " +
                        ",PRC.NOCHEQUE " +
                        ",TO_CHAR(PRN.INICIO, 'DD/MM/YYYY') FECHA " +
                        ",PRN.CDGCO " +
                        ",CASE WHEN EXISTS (SELECT CDGCL FROM CHEQUE_NONEG NN " +
                                             "WHERE NN.CDGEM = PRC.CDGEM " +
                                             "AND NN.CDGCLNS = PRC.CDGNS " +
                                             "AND NN.CICLO = PRC.CICLO " +
                                             "AND NN.CDGCL = PRC.CDGCL " +
                                             "AND NN.TIPO = 'DE') THEN '' " +
                        "ELSE " +
                            "'NO NEGOCIABLE' " +
                        "END NO_NEGOCIABLE " +
                        "FROM PRC, PRN, CL " +
                        "WHERE PRC.CDGEM = '" + empresa + "' " +
                        "AND PRC.CDGNS = '" + codigo + "' " +
                        "AND PRC.CICLO = '" + ciclo + "' " +
                        "AND PRC.NOCHEQUE IN (" + cheque + ") " +
                        "AND PRN.CDGEM = PRC.CDGEM " +
                        "AND PRN.CDGNS = PRC.CDGNS " +
                        "AND PRN.CICLO = PRC.CICLO " +
                        "AND CL.CDGEM = PRC.CDGEM " +
                        "AND CL.CODIGO = PRC.CDGCL " +
                        "ORDER BY PRC.NOCHEQUE";
            }
            else if(tipo == "G")
            {
                query = "SELECT PGS.CDGCLNS CDGNS " +
                        ",PGS.CICLO " +
                        ",PGS.NOMBRE NOMBREC " +
                        ",ABS(PGS.CANTIDAD) MONTO " +
                        ",PGS.CDGCB " +
                        ",PGS.NOCHEQUE " +
                        ",TO_CHAR(PGS.FPAGO, 'DD/MM/YYYY') FECHA " +
                        ",NULL CDGCO " +
                        ",CASE WHEN (SELECT COUNT(*) " +
                                   "FROM CHEQUE_NONEG NN " +
                                   "WHERE NN.CDGEM = PGS.CDGEM " +
                                   "AND NN.CDGCLNS = PGS.CDGCLNS " +
                                   "AND NN.FPAGO = PGS.FPAGO " +
                                   "AND NN.SEC = PGS.SECPAGO " +
                                   "AND NN.TIPO = 'GL' ) > 0 THEN '' " +
                        "ELSE " +
                             "'NO NEGOCIABLE' " +
                        "END NO_NEGOCIABLE " +
                        "FROM PAG_GAR_SIM PGS " +
                        "WHERE PGS.CDGEM = '" + empresa + "' " +
                        "AND PGS.CDGCLNS = '" + codigo + "' " +
                        "AND PGS.CLNS = '" + clns + "' " +
                        "AND PGS.NOCHEQUE IN (" + cheque + ") " +
                        "ORDER BY PGS.NOCHEQUE";
            }
            else if (tipo == "E")
            {
                query = "SELECT PDE.CDGCLNS CDGNS, " +
                        ",PDE.CICLO " +
                        ",PDE.NOMBRE NOMBREC " +
                        ",PDE.CANTIDAD MONTO " +
                        ",PDE.CDGCB " +
                        ",PDE.NOCHEQUE " +
                        ",TO_CHAR(PDE.FPAGO, 'DD/MM/YYYY') FECHA " +
                        ",NULL CDGCO " +
                        ",CASE WHEN (SELECT COUNT(*) " +
                                   "FROM CHEQUE_NONEG NN " +
                                   "WHERE NN.CDGEM = PDE.CDGEM " +
                                   "AND NN.CDGCLNS = PDE.CDGCLNS " +
                                   "AND NN.FPAGO = PDE.FPAGO " +
                                   "AND NN.SEC = PDE.SECMOV " +
                                   "AND NN.TIPO = 'EX' ) > 0 THEN '' " +
                        "ELSE " +
                             "'NO NEGOCIABLE' " +
                        "END NO_NEGOCIABLE " +
                        "FROM PAG_DEV_EXC PDE " +
                        "WHERE PDE.CDGEM = '" + empresa + "' " +
                        "AND PDE.CDGCLNS = '" + codigo + "' " +
                        "AND PDE.CICLO = '" + ciclo + "' " +
                        "AND PDE.CLNS = '" + clns + "' " +
                        "AND PDE.NOCHEQUE IN (" + cheque + ") " +
                        "ORDER BY PDE.NOCHEQUE";
            }
            iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);

            if (dsAcred.Tables[0].Rows.Count > 0)
            {
                LlenaImpresionCheques(dsAcred, cuenta);
            }
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");

        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getRepSitCartera(string id, string fecha, string nivel, string cartVig, string cartVenc, string cartRest, string cartCast, string usuario,
                            string nomUsuario, string region, string sucursal, string coord, string asesor, string tipoProd, string nivelMora, string titulo)
    {
        string empresa = cdgEmpresa;
        string query = string.Empty;
        int iRes;
        try
        {
            DataSet ds = new DataSet();
            DataSet dsPDI = new DataSet();

            if (id == "1" || id == "17")
                iRes = db.myExecuteNonQuery("SP_REP_ANT_SALDOS", CommandType.StoredProcedure, oP.ParamsMora(empresa, Convert.ToDateTime(fecha),
                                        Convert.ToInt32(nivel), Convert.ToInt32(cartVig), Convert.ToInt32(cartVenc), Convert.ToInt32(cartRest),
                                        Convert.ToInt32(cartCast), usuario, region, sucursal, coord, asesor, tipoProd, nivelMora));

            else if (id == "46")
                iRes = db.myExecuteNonQuery("SP_REP_ANT_SALDOS_HIST", CommandType.StoredProcedure, oP.ParamsMora(empresa, Convert.ToDateTime(fecha),
                                            Convert.ToInt32(nivel), Convert.ToInt32(cartVig), Convert.ToInt32(cartVenc), Convert.ToInt32(cartRest),
                                            Convert.ToInt32(cartCast), usuario, region, sucursal, coord, asesor));

            query = "SELECT RAS.* " +
                    ",TO_CHAR(SYSDATE,'DD/MM/YYYY') FECHAIMP " +
                    ",TO_CHAR(SYSDATE,'HH24:MI:SS') HORAIMP " +
                    "FROM REP_ANT_SALDOS RAS " +
                    "WHERE CDGEM = '" + empresa + "' " +
                    "AND CDGPE = '" + usuario + "' " +
                    "ORDER BY CDGRG, CDGCO, SALDO DESC, DIAS_MORA DESC, CDGCLNS";

            iRes = db.ExecuteDS(ref ds, query, CommandType.Text);

            string queryPDI = "SELECT COUNT(*) REG_PDI " +
                              ",SUM(CANTIDAD) MONTO_PDI " +
                              "FROM PDI " +
                              "WHERE CDGEM = '" + empresa + "' " +
                              "AND ESTATUS = 'RE' " + 
                              "AND CDGCB <> '12'";

            iRes = db.ExecuteDS(ref dsPDI, queryPDI, CommandType.Text);

            if (ds.Tables[0].Rows.Count > 0)
                LlenaRptSitCartera(ds, dsPDI, id, Convert.ToDateTime(fecha), Convert.ToInt32(nivel), titulo, nomUsuario);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    private void getSolicitudCred(string grupo, string ciclo, string usuario, string nomUsuario)
    {
        string empresa = cdgEmpresa;
        int iRes;

        DataSet dsGrupo = new DataSet();
        DataSet dsAcred = new DataSet();
        DataSet dsExcAcred = new DataSet();
        DataSet dsMaxAcred = new DataSet();
        DataSet ds = new DataSet();

        try
        {
            //SE INVOCA EL PROCEDIMIENTO PARA LLENAR LA TABLA TEMPORAL DE ENCABEZADO DE SOLICITUD
            iRes = db.myExecuteNonQuery("SP_REP_SOLICITUD_ENC", CommandType.StoredProcedure,
                                        oP.ParamsRepSol(empresa, grupo, ciclo, usuario));

            string query = "SELECT * " +
                           "FROM REP_SOLICITUD_ENC " +
                           "WHERE CDGEM = '" + empresa +"' " +
                           "AND CDGPE = '" + usuario + "'";

            iRes = db.ExecuteDS(ref dsGrupo, query, CommandType.Text);

            query = "SELECT A.CDGCL NO_CTE " +
                    ",NOMBREC(CL.CDGEM,CL.CODIGO,'I','A',NULL,NULL,NULL,NULL) CLIENTE " +
                    ",A.CANTENTRE " +
                    ",LPAD(TO_CHAR((SELECT COUNT(*) FROM PRC WHERE CDGEM = A.CDGEM AND CDGCL = A.CDGCL AND SITUACION IN ('E', 'L') AND CANTENTRE > 0 AND TRUNC(SOLICITUD) <= TRUNC(PRN.SOLICITUD) GROUP BY CDGCL)+1),2,'0') CICLO_SOL  " +
                    "FROM PRC A, PRN, CL " +
                    "WHERE PRN.CDGEM = A.CDGEM " +
                    "AND PRN.CDGNS = A.CDGNS " +
                    "AND PRN.CICLO = A.CICLO " +
                    "AND A.CDGEM = CL.CDGEM " +
                    "AND A.CDGCL = CL.CODIGO " +
                    "AND PRN.CDGEM = '" + empresa + "' " +
                    "AND PRN.CDGNS = '" + grupo + "' " +
                    "AND PRN.CICLO = '" + ciclo + "' " +
                    "AND A.CANTENTRE > 0 " +
                    "ORDER BY CLIENTE";

            iRes = db.ExecuteDS(ref dsAcred, query, CommandType.Text);

            //SE INVOCA EL PROCEDIMIENTO PARA LLENAR LA TABLA TEMPORAL DE EXCEPCIONES DE CREDITO
            iRes = db.myExecuteNonQuery("SP_EXPACREDSOL", CommandType.StoredProcedure,
                                        oP.ParamsRepSol(empresa, grupo, ciclo, usuario));

            //CONSULTA QUE OBTIENE LOS DATOS GENERADOS EN LA EJECUCION DEL PROCEDIMIENTO SP_EXPACREDSOL 
            query = "SELECT CDGCL ACRED " + 
                    ",NOMBRE " +
                    ",DESCRIPCION " +
                    ",OBSERVACION " +
                    "FROM REP_SOL_EXC " +
                    "WHERE CDGEM = '" + empresa + "' " +
                    "AND CDGPE = '" + usuario + "' ";

            iRes = db.ExecuteDS(ref dsExcAcred, query, CommandType.Text);

            query = "SELECT PRC.CDGCL ACRED " +
                 ",NOMBREC(CL.CDGEM,CL.CODIGO,'I','A',NULL,NULL,NULL,NULL) NOMBRE_ACRED " +
                 ",CM.MONTOMAX " +
                 "FROM PRC, CL_MARCA CM, CL " +
                 "WHERE PRC.CDGEM = '" + empresa + "' " +
                 "AND PRC.CDGNS = '" + grupo + "' " +
                 "AND PRC.CICLO = '" + ciclo + "' " +
                 "AND PRC.CLNS = 'G' " +
                 "AND PRC.CANTENTRE > 0 " +
                 "AND PRC.SITUACION IN ('E','L') " +
                 "AND CL.CDGEM = PRC.CDGEM " +
                 "AND CL.CODIGO = PRC.CDGCL " +
                 "AND CM.CDGEM = PRC.CDGEM " +
                 "AND CM.CDGCL = PRC.CDGCL " +
                 "AND CM.TIPOMARCA = 'EN' " +
                 "AND CM.ESTATUS = 'A' " +
                 "AND CM.BAJA IS NULL " +
                 "ORDER BY CL.PRIMAPE, CL.SEGAPE, CL.NOMBRE1, CL.NOMBRE2";

            iRes = db.ExecuteDS(ref dsMaxAcred, query, CommandType.Text);

            query = "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY') FECHAIMP, " +
                    "TO_CHAR(SYSDATE,'HH24:MI:SS') HORAIMP " +
                    "FROM DUAL";

            iRes = db.ExecuteDS(ref ds, query, CommandType.Text);

            if (dsGrupo.Tables[0].Rows.Count > 0 && dsAcred.Tables[0].Rows.Count > 0)
                LlenaSolicitudCred(ds, dsGrupo, dsAcred, dsExcAcred, dsMaxAcred, nomUsuario);
            else
                LlenaRptError("", "No se encontró información relacionada con los datos de consulta");
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    #region "DespliegaPdf"

    private void DisplayPdf(ReportDocument grcReporte)
    {
        try
        {
            MemoryStream oStream = new MemoryStream();
            try
            {
                if(tipoDoc == "PDF")
                    oStream = (MemoryStream)grcReporte.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                else if(tipoDoc == "DOC")
                    oStream = (MemoryStream)grcReporte.ExportToStream(CrystalDecisions.Shared.ExportFormatType.WordForWindows);

            }
            catch (Exception ex)
            {
                string mensaje = ex.Message;
            }
            Response.Clear();
            Response.ClearContent();
            Response.ClearHeaders();
            Response.Buffer = true;
            if(tipoDoc == "PDF")
                Response.ContentType = "application/pdf";
            else if(tipoDoc == "DOC")
                Response.ContentType = "application/msword";
            Response.BinaryWrite(oStream.ToArray());
            oStream.Dispose();
            oStream.Close();
            grcReporte.Dispose();
            grcReporte.Close();
            Response.Flush();
            Response.Close();
        }
        catch (Exception exAsync)
        {
            string mensaje = exAsync.Message;
            Response.Close();
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }

    #endregion

    #region "CargaReporte"

    private Boolean CargaReporte(string nombreReporte)
    {
        try
        {
            grcReporte.Load(Server.MapPath("") + "\\" + nombreReporte);
            return true;
        }
        catch (Exception ex)
        {
            String detalle;
            detalle = ex.Message;
            return false;
        }
    }

#endregion

    private void LlenaRptError(string LD, string descripcion)
    {
        DsError.dtErrorDataTable dtError = new DsError.dtErrorDataTable();
        DataTable dtReporteError = (DataTable)dtError;
        DataSet dsRError = new DsError();
        DataRow drError = dtReporteError.NewRow();
        drError["Descripcion"] = descripcion;
        dtReporteError.Rows.Add(drError);
        dsRError.Tables["dtError"].ImportRow(drError);
        CargaReporte("Error.rpt");
        grcReporte.SetDataSource(dsRError);
        DisplayPdf(grcReporte);
    }

    private void LlenaAutPagoGL(DataSet dsGrupo, DataSet dsAcred, DataSet dsDoc)
    {
        dsControl.dtAcreditadoDataTable dt = new dsControl.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControl.dtGrupoDataTable dtG = new dsControl.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        DataSet dsCont = new dsControl();

        int i;
        int filasAcred;

        string fecInicio = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["INICIO"]).ToString("dd/MM/yyyy");
        string fecFin = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["FECHAFIN"]).ToString("dd/MM/yyyy");
        string doc = dsDoc.Tables[0].Rows[0]["ARCHIVO"].ToString();

        DataRow drGrupo = dtGrupo.NewRow();
        drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
        drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
        drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
        drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORDINACION"];
        drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["NOMBREASESOR"];
        drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["NOMBREGERENTE"];
        drGrupo["DIA_INICIO"] = fecInicio.Substring(0, 2);
        drGrupo["MES_INICIO"] = fecInicio.Substring(3, 2);
        drGrupo["ANIO_INICIO"] = fecInicio.Substring(6, 4);
        drGrupo["DIA_FIN"] = fecFin.Substring(0, 2);
        drGrupo["MES_FIN"] = fecFin.Substring(3, 2);
        drGrupo["ANIO_FIN"] = fecFin.Substring(6, 4);

        dtGrupo.Rows.Add(drGrupo);
        dsCont.Tables["dtGrupo"].ImportRow(drGrupo);

        filasAcred = dsAcred.Tables[0].Rows.Count;

        for (i = 0; i < filasAcred; i++)
        {
            DataRow drAcred = dtAcred.NewRow();
            drAcred["NUMERO"] = i + 1;
            drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
            dtAcred.Rows.Add(drAcred);
            dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
        }

        CargaReporte(doc);
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaCartaGarantia(DataSet dsCarta, string grupo)
    {
        dsDocumentos.dtCartaPagoDataTable dt = new dsDocumentos.dtCartaPagoDataTable();
        DataTable dtRepCarta = (DataTable)dt;

        DataSet dsRepCarta = new dsDocumentos();

        string cuenta = string.Empty;
        string sucursal = string.Empty;
        string tipoProd = string.Empty;

        BarcodeLib.Barcode bar = new BarcodeLib.Barcode();
        bar.IncludeLabel = true;
        System.Drawing.Image imgBarcodeOpenPay = bar.Encode(TYPE.CODE128, dsCarta.Tables[0].Rows[0]["REF_OPENPAY"].ToString(), System.Drawing.Color.Black, System.Drawing.Color.White, 400, 100);
        System.Drawing.Image imgBarcodePayCash = bar.Encode(TYPE.CODE128, dsCarta.Tables[0].Rows[0]["REF_PAYCASH"].ToString(), System.Drawing.Color.Black, System.Drawing.Color.White, 400, 100);
        //imgBarcodeOpenPay.Save(@"C:\imgBarcodeOpenPay.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
        //imgBarcodePayCash.Save(@"C:\imgBarcodePayCash.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

        tipoProd = dsCarta.Tables[0].Rows[0]["TIPO_PROD"].ToString();

        DataRow drCarta = dtRepCarta.NewRow();
        drCarta["CODIGO_GPO"] = dsCarta.Tables[0].Rows[0]["CODIGO_GPO"];
        drCarta["NOMBRE_GPO"] = dsCarta.Tables[0].Rows[0]["NOMBRE_GPO"];
        drCarta["CODIGO_CTE"] = dsCarta.Tables[0].Rows[0]["CODIGO_CTE"];
        drCarta["NOMBRE_CTE"] = dsCarta.Tables[0].Rows[0]["NOMBRE_CTE"];
        drCarta["PRESIDENTE"] = dsCarta.Tables[0].Rows[0]["PRESIDENTE"];
        drCarta["TESORERO"] = dsCarta.Tables[0].Rows[0]["TESORERO"];
        drCarta["SECRETARIO"] = dsCarta.Tables[0].Rows[0]["SECRETARIO"];
        drCarta["SUPERVISOR"] = dsCarta.Tables[0].Rows[0]["SUPERVISOR"];
        drCarta["NOM_SUCURSAL"] = dsCarta.Tables[0].Rows[0]["NOM_SUCURSAL"];
        drCarta["PAGO_PARCIAL_LETRA"] = dsCarta.Tables[0].Rows[0]["PAGO_PARCIAL_LETRA"];
        drCarta["NOM_ASESOR"] = dsCarta.Tables[0].Rows[0]["NOM_GERENTE_SUC"];
        drCarta["REFERENCIA"] = dsCarta.Tables[0].Rows[0]["REF_TELECOMM"];
        drCarta["REF_BANCOMER"] = dsCarta.Tables[0].Rows[0]["REF_BANCOMER"];
        drCarta["REF_BANSEFI"] = dsCarta.Tables[0].Rows[0]["REF_BANSEFI"];
        drCarta["REF_OPENPAY"] = dsCarta.Tables[0].Rows[0]["REF_OPENPAY"];
        drCarta["REF_PAYCASH"] = dsCarta.Tables[0].Rows[0]["REF_PAYCASH"];
        drCarta["NOM_PROD"] = dsCarta.Tables[0].Rows[0]["NOM_PROD"];
        drCarta["RECA"] = dsCarta.Tables[0].Rows[0]["NOM_INTEG_DNI"];
        drCarta["BARCODE_OPENPAY"] = Funciones.ImagenToByte(imgBarcodeOpenPay);
        drCarta["BARCODE_PAYCASH"] = Funciones.ImagenToByte(imgBarcodePayCash);
        sucursal = dsCarta.Tables[0].Rows[0]["CDGCO"].ToString();

        dtRepCarta.Rows.Add(drCarta);
        dsRepCarta.Tables["dtCartaPago"].ImportRow(drCarta);


        if (tipoProd == "01")
            CargaReporte("GarantiaComunal_04.rpt");
        else if (tipoProd == "04")
            CargaReporte("GarantiaSolidario_04.rpt");

        grcReporte.SetDataSource(dsRepCarta);
        DisplayPdf(grcReporte);
    }

    private void LlenaCartaPago(DataSet dsCarta, DataSet dsCartaDet, string grupo)
    {
        dsDocumentos.dtCartaPagoDataTable dt = new dsDocumentos.dtCartaPagoDataTable();
        DataTable dtRepCarta = (DataTable)dt;

        dsDocumentos.dtCartaPagoDetDataTable dtD = new dsDocumentos.dtCartaPagoDetDataTable();
        DataTable dtRepCartaDet = (DataTable)dtD;

        DataSet dsRepCarta = new dsDocumentos();

        BarcodeLib.Barcode bar = new BarcodeLib.Barcode();
        bar.IncludeLabel = true;
        System.Drawing.Image imgBarcodePayCash = bar.Encode(TYPE.CODE128, dsCarta.Tables[0].Rows[0]["REF_PAYCASH"].ToString(), System.Drawing.Color.Black, System.Drawing.Color.White, 400, 100);
        System.Drawing.Image imgBarcodePayCashGl = bar.Encode(TYPE.CODE128, dsCarta.Tables[0].Rows[0]["REF_PAYCASH_GL"].ToString(), System.Drawing.Color.Black, System.Drawing.Color.White, 400, 100);

        string cuenta = string.Empty;
        string sucursal = string.Empty;
        string tipoProd = string.Empty;
        int i;
        int filasAcred;

        tipoProd = dsCarta.Tables[0].Rows[0]["TIPO_PROD"].ToString();

        DataRow drCarta = dtRepCarta.NewRow();
        drCarta["CODIGO_GPO"] = dsCarta.Tables[0].Rows[0]["CODIGO_GPO"];
        drCarta["NOMBRE_GPO"] = dsCarta.Tables[0].Rows[0]["NOMBRE_GPO"];
        drCarta["CICLO"] = dsCarta.Tables[0].Rows[0]["CICLO"];
        drCarta["NUM_INT"] = dsCarta.Tables[0].Rows[0]["NUM_INT"];
        drCarta["PLAZO"] = dsCarta.Tables[0].Rows[0]["PLAZO"];
        drCarta["TASA"] = dsCarta.Tables[0].Rows[0]["TASA"];
        drCarta["INICIO"] = dsCarta.Tables[0].Rows[0]["INICIO"];
        drCarta["FIN"] = dsCarta.Tables[0].Rows[0]["FIN"];
        drCarta["CANTENTRE"] = dsCarta.Tables[0].Rows[0]["CANTENTRE"];
        drCarta["TOTAL"] = dsCarta.Tables[0].Rows[0]["TOTAL"];
        drCarta["PARCIALIDAD"] = dsCarta.Tables[0].Rows[0]["PARCIALIDAD"];
        drCarta["REF_PAYCASH"] = dsCarta.Tables[0].Rows[0]["REF_PAYCASH"];
        drCarta["REF_PAYCASH_GL"] = dsCarta.Tables[0].Rows[0]["REF_PAYCASH_GL"];
        drCarta["BARCODE_PAYCASH"] = Funciones.ImagenToByte(imgBarcodePayCash);
        drCarta["BARCODE_PAYCASH_GL"] = Funciones.ImagenToByte(imgBarcodePayCashGl);

        dtRepCarta.Rows.Add(drCarta);
        dsRepCarta.Tables["dtCartaPago"].ImportRow(drCarta);

        filasAcred = dsCartaDet.Tables[0].Rows.Count;

        for (i = 0; i < filasAcred; i++)
        {
            DataRow drCartaDet = dtRepCartaDet.NewRow();
            drCartaDet["NUMERO"] = i + 1;
            drCartaDet["NOMCL"] = dsCartaDet.Tables[0].Rows[i]["NOMBRE"];
            drCartaDet["CANTENTRE"] = dsCartaDet.Tables[0].Rows[i]["CANTENTRE"];
            drCartaDet["TOTAL"] = dsCartaDet.Tables[0].Rows[i]["TOTAL"];
            drCartaDet["PARCIALIDAD"] = dsCartaDet.Tables[0].Rows[i]["PARCIALIDAD"];
            dtRepCartaDet.Rows.Add(drCartaDet);
            dsRepCarta.Tables["dtCartaPagoDet"].ImportRow(drCartaDet);
        }

        if (grupo != null)
        {
            if (tipoProd == "01" || tipoProd == "03")
                CargaReporte("CartaPagoComunalPayCash.rpt");
            else if (tipoProd == "02")
                CargaReporte("CartaPagoAdicionalPayCash.rpt");
        }
        else if (grupo == null)
            CargaReporte("CartaPagoIndividual.rpt");

        grcReporte.SetDataSource(dsRepCarta);
        DisplayPdf(grcReporte);
    }

    private void LlenaDocMicroseguro(string id, DataSet dsCons, string tipoProd, string formato)
    {
        dsDocumentos.dtDocMicrosegDataTable dt = new dsDocumentos.dtDocMicrosegDataTable();
        DataTable dtDocMicroseg = (DataTable)dt;

        DataSet dsDocMicroseg = new dsDocumentos();

        int i;
        int contFilas;

        if (dsCons != null)
        {
            if (dsCons.Tables.Count > 0)
            {
                contFilas = dsCons.Tables[0].Rows.Count;

                for (i = 0; i < contFilas; i++)
                {
                    DataRow drCons = dtDocMicroseg.NewRow();
                    drCons["CODIGO_CTE"] = dsCons.Tables[0].Rows[i]["CODIGO_CTE"];
                    drCons["NOMBRE_CTE"] = dsCons.Tables[0].Rows[i]["NOMBRE_CTE"];
                    drCons["FECHA_INICIO"] = dsCons.Tables[0].Rows[i]["FECINI"];
                    drCons["FECHA_FIN"] = dsCons.Tables[0].Rows[i]["FECFIN"];
                    drCons["FECHA_NAC"] = dsCons.Tables[0].Rows[i]["NACIMIENTO"];
                    dtDocMicroseg.Rows.Add(drCons);
                    dsDocMicroseg.Tables["dtDocMicroseg"].ImportRow(drCons);
                }
            }
        }
        CargaReporte(formato);
        grcReporte.SetDataSource(dsDocMicroseg);
        DisplayPdf(grcReporte);
    }


    private void LlenaConsultaBuro(DataSet dsInfo, DataSet dsNom, DataSet dsDir, DataSet dsCta, DataSet dsRes, DataSet dsCons, string nomUsuario)
    {
        dsBuro.dtNombreDataTable dtN = new dsBuro.dtNombreDataTable();
        DataTable dtNom = (DataTable)dtN;

        dsBuro.dtDireccionDataTable dtD = new dsBuro.dtDireccionDataTable();
        DataTable dtDir = (DataTable)dtD;

        dsBuro.dtCuentaDataTable dtC = new dsBuro.dtCuentaDataTable();
        DataTable dtCta = (DataTable)dtC;

        dsBuro.dtResumenDataTable dtR = new dsBuro.dtResumenDataTable();
        DataTable dtRes = (DataTable)dtR;

        dsBuro.dtConsultaDataTable dtCon = new dsBuro.dtConsultaDataTable();
        DataTable dtCons = (DataTable)dtCon;

        dsBuro.dtInfoRepDataTable dtInfo = new dsBuro.dtInfoRepDataTable();
        DataTable dtReporteInfo = (DataTable)dtInfo;

        DataSet dsBuro = new dsBuro();

        int i;
        int filasDir;
        int filasCta;
        int filasRes;
        int filasCons;

        DataRow drInfo = dtInfo.NewRow();
        drInfo["USUARIO"] = nomUsuario;
        drInfo["FEC_CONS"] = dsInfo.Tables[0].Rows[0]["FECHACONS"];
        drInfo["FEC_VIG"] = dsInfo.Tables[0].Rows[0]["FECHAVIG"];
        drInfo["FEC_IMP"] = dsInfo.Tables[0].Rows[0]["FECHAIMP"];
        drInfo["HORA_IMP"] = dsInfo.Tables[0].Rows[0]["HORAIMP"];
        dtInfo.Rows.Add(drInfo);
        dsBuro.Tables["dtInfoRep"].ImportRow(drInfo);


        DataRow drNom = dtNom.NewRow();
        drNom["APATERNO"] = dsNom.Tables[0].Rows[0]["PRIMAPE"];
        drNom["AMATERNO"] = dsNom.Tables[0].Rows[0]["SEGAPE"];
        drNom["AOPCIONAL"] = dsNom.Tables[0].Rows[0]["OPCAPE"];
        drNom["PNOMBRE"] = dsNom.Tables[0].Rows[0]["NOMBRE1"];
        drNom["SNOMBRE"] = dsNom.Tables[0].Rows[0]["NOMBRE2"];
        drNom["FECNAC"] = dsNom.Tables[0].Rows[0]["FECNAC"];
        drNom["RFC"] = dsNom.Tables[0].Rows[0]["RFC"];
        drNom["CURP"] = dsNom.Tables[0].Rows[0]["CURP"];
        drNom["PREFPERS"] = dsNom.Tables[0].Rows[0]["PREFPERS"];
        drNom["SUBFIJO"] = dsNom.Tables[0].Rows[0]["SUBFIJO"];
        drNom["NAC"] = dsNom.Tables[0].Rows[0]["NAC"];
        drNom["RESID"] = dsNom.Tables[0].Rows[0]["RESID"];
        drNom["LICENCIA"] = dsNom.Tables[0].Rows[0]["LICENCIA"];
        drNom["EDOCIVIL"] = dsNom.Tables[0].Rows[0]["EDOCIVIL"];
        drNom["SEXO"] = dsNom.Tables[0].Rows[0]["SEXO"];
        drNom["CEDULA"] = dsNom.Tables[0].Rows[0]["CEDULA"];
        drNom["IFE"] = dsNom.Tables[0].Rows[0]["IFE"];
        drNom["IMPOTROPAIS"] = dsNom.Tables[0].Rows[0]["IMPOTROPAI"];
        drNom["OTROPAIS"] = dsNom.Tables[0].Rows[0]["CDGPAI"];
        drNom["DEPEND"] = dsNom.Tables[0].Rows[0]["DEPEND"];
        drNom["EDADDEPEND"] = dsNom.Tables[0].Rows[0]["EDADDEPEND"];
        drNom["FECINFODEPEND"] = dsNom.Tables[0].Rows[0]["FECINFODEPEND"];
        drNom["FECDEFUNC"] = dsNom.Tables[0].Rows[0]["FECDEFUNC"];
        dtNom.Rows.Add(drNom);
        dsBuro.Tables["dtNombre"].ImportRow(drNom);

        filasDir = dsDir.Tables[0].Rows.Count;

        for (i = 0; i < filasDir; i++)
        {
            DataRow drDir = dtDir.NewRow();
            drDir["NUM"] = i + 1;
            drDir["DIR1"] = dsDir.Tables[0].Rows[i]["DIR1"];
            drDir["DIR2"] = dsDir.Tables[0].Rows[i]["DIR2"];
            drDir["COLONIA"] = dsDir.Tables[0].Rows[i]["COLONIA"];
            drDir["MUNICIPIO"] = dsDir.Tables[0].Rows[i]["MUNICIPIO"];
            drDir["CIUDAD"] = dsDir.Tables[0].Rows[i]["CIUDAD"];
            drDir["ESTADO"] = dsDir.Tables[0].Rows[i]["ESTADO"];
            drDir["CODPOSTAL"] = dsDir.Tables[0].Rows[i]["CDGPOSTAL"];
            drDir["FECRESID"] = dsDir.Tables[0].Rows[i]["FECRESID"];
            drDir["TELEFONO"] = dsDir.Tables[0].Rows[i]["TELEFONO"];
            drDir["EXT"] = dsDir.Tables[0].Rows[i]["EXT"];
            drDir["FAX"] = dsDir.Tables[0].Rows[i]["FAX"];
            drDir["TIPODOM"] = dsDir.Tables[0].Rows[i]["TIPODOM"];
            drDir["INDESPDOM"] = dsDir.Tables[0].Rows[i]["INDESPDOM"];
            drDir["FECDOM"] = dsDir.Tables[0].Rows[i]["FECREPDIR"];
            dtDir.Rows.Add(drDir);
            dsBuro.Tables["dtDireccion"].ImportRow(drDir);
        }

        filasCta = dsCta.Tables[0].Rows.Count;

        for (i = 0; i < filasCta; i++)
        {
            DataRow drCta = dtCta.NewRow();
            drCta["NUM"] = i + 1;
            drCta["FECACT"] = dsCta.Tables[0].Rows[i]["FECACT"];
            drCta["REGIMP"] = dsCta.Tables[0].Rows[i]["REGIMP"];
            drCta["CVEOTORGA"] = dsCta.Tables[0].Rows[i]["CDGOTOR"];
            drCta["NOMOTORGA"] = dsCta.Tables[0].Rows[i]["NOMOTOR"];
            drCta["IDSIC"] = dsCta.Tables[0].Rows[i]["SIC"];
            drCta["CUENTA"] = dsCta.Tables[0].Rows[i]["CUENTA"];
            drCta["TIPORESPONS"] = dsCta.Tables[0].Rows[i]["TIPORESP"];
            drCta["DESCTIPORESP"] = dsCta.Tables[0].Rows[i]["DESCTIPORESP"].ToString().ToUpper();
            drCta["TIPOCUENTA"] = dsCta.Tables[0].Rows[i]["TIPOCTA"];
            drCta["DESCTIPOCTA"] = dsCta.Tables[0].Rows[i]["DESCTIPOCTA"].ToString().ToUpper();
            drCta["TIPOCONTRATO"] = dsCta.Tables[0].Rows[i]["TIPOCONT"];
            drCta["DESCTIPOCONT"] = dsCta.Tables[0].Rows[i]["DESCTIPOCRED"].ToString().ToUpper();
            drCta["CVEMONEDA"] = "MN";
            drCta["VALORACTIVO"] = dsCta.Tables[0].Rows[i]["VALACT"];
            drCta["NUMPAGOS"] = dsCta.Tables[0].Rows[i]["NUMPAGOS"];
            drCta["FRECPAGOS"] = dsCta.Tables[0].Rows[i]["FRECPAGOS"];
            drCta["DESCFRECPAGOS"] = dsCta.Tables[0].Rows[i]["DESCFRECPAGOS"].ToString().ToUpper();
            drCta["MONTOPAGAR"] = dsCta.Tables[0].Rows[i]["MONTO"];
            drCta["FECCUENTA"] = dsCta.Tables[0].Rows[i]["FECINICTA"];
            drCta["FECULTPAGO"] = dsCta.Tables[0].Rows[i]["FECULTPAGO"];
            drCta["FECULTCOMPRA"] = dsCta.Tables[0].Rows[i]["FECULTCOMP"];
            drCta["FECCIERRECTA"] = dsCta.Tables[0].Rows[i]["FECFINCTA"];
            drCta["FECREP"] = dsCta.Tables[0].Rows[i]["FECREP"];
            drCta["MODOREP"] = dsCta.Tables[0].Rows[i]["MODOREP"];
            drCta["FECCTACERO"] = dsCta.Tables[0].Rows[i]["FECSINSALDO"];
            drCta["GARANTIA"] = dsCta.Tables[0].Rows[i]["GARANTIA"];
            drCta["CREDMAX"] = dsCta.Tables[0].Rows[i]["CREDMAX"];
            drCta["SALDOACT"] = dsCta.Tables[0].Rows[i]["SDOACT"];
            drCta["LIMCRED"] = dsCta.Tables[0].Rows[i]["LIMCRD"];
            drCta["SALDOVENC"] = dsCta.Tables[0].Rows[i]["SDOVENC"];
            drCta["PAGOSVENC"] = dsCta.Tables[0].Rows[i]["PAGOSVENC"];
            drCta["FORMAPAGO"] = dsCta.Tables[0].Rows[i]["MOP"];
            drCta["DESCFORMAPAGO"] = dsCta.Tables[0].Rows[i]["DESCMOP"].ToString().ToUpper();
            drCta["HISTPAGO"] = dsCta.Tables[0].Rows[i]["HISTPAGOS"];
            drCta["FECRECHISTPAGO"] = dsCta.Tables[0].Rows[i]["FECRECHPAGOS"];
            drCta["FECANTHISTPAGO"] = dsCta.Tables[0].Rows[i]["FECANTHPAGOS"];
            drCta["CVEOBS"] = dsCta.Tables[0].Rows[i]["CDGOBS"];
            drCta["DESCCVEOBS"] = dsCta.Tables[0].Rows[i]["DESCCDGOBS"];
            drCta["TOTALPAGOS"] = dsCta.Tables[0].Rows[i]["PAGOSREP"];
            drCta["TOTALPAGOSMOP02"] = dsCta.Tables[0].Rows[i]["MOP02"];
            drCta["TOTALPAGOSMOP03"] = dsCta.Tables[0].Rows[i]["MOP03"];
            drCta["TOTALPAGOSMOP04"] = dsCta.Tables[0].Rows[i]["MOP04"];
            drCta["TOTALPAGOSMOP05"] = dsCta.Tables[0].Rows[i]["MOP05"];
            drCta["IMPSALDOMORA"] = dsCta.Tables[0].Rows[i]["IMPMOR"];
            drCta["FECSALDOMORA"] = dsCta.Tables[0].Rows[i]["FECIMPMOR"];
            drCta["MOPHISTORICO"] = dsCta.Tables[0].Rows[i]["MOPMOR"];
            drCta["FECRESTRUC"] = dsCta.Tables[0].Rows[i]["FECREST"];
            drCta["ANIO1"] = dsCta.Tables[0].Rows[i]["ANIO1"];
            drCta["ANIO2"] = dsCta.Tables[0].Rows[i]["ANIO2"];
            drCta["ANIO3"] = dsCta.Tables[0].Rows[i]["ANIO3"];
            drCta["MES1"] = dsCta.Tables[0].Rows[i]["MES1"];
            drCta["MES2"] = dsCta.Tables[0].Rows[i]["MES2"];
            drCta["MES3"] = dsCta.Tables[0].Rows[i]["MES3"];
            drCta["SEGMES1"] = dsCta.Tables[0].Rows[i]["SEGMES1"];
            drCta["SEGMES2"] = dsCta.Tables[0].Rows[i]["SEGMES2"];
            drCta["SEGMES3"] = dsCta.Tables[0].Rows[i]["SEGMES3"];

            dtCta.Rows.Add(drCta);
            dsBuro.Tables["dtCuenta"].ImportRow(drCta);
        }

        filasRes = dsRes.Tables[0].Rows.Count;

        for (i = 0; i < filasRes; i++)
        {
            DataRow drRes = dtRes.NewRow();
            drRes["MOP"] = dsRes.Tables[0].Rows[i]["MOP"];
            drRes["NUMABR"] = dsRes.Tables[0].Rows[i]["NUMABR"];
            drRes["LIMABR"] = dsRes.Tables[0].Rows[i]["LIMABR"];
            drRes["CREDMAXABR"] = dsRes.Tables[0].Rows[i]["CREDMAXABR"];
            drRes["SDOACTABR"] = dsRes.Tables[0].Rows[i]["SDOACTABR"];
            drRes["SDOVENCABR"] = dsRes.Tables[0].Rows[i]["SDOVENCABR"];
            drRes["MONTOABR"] = dsRes.Tables[0].Rows[i]["MONTOABR"];
            drRes["NUMCER"] = dsRes.Tables[0].Rows[i]["NUMCER"];
            drRes["LIMCER"] = dsRes.Tables[0].Rows[i]["LIMCER"];
            drRes["CREDMAXCER"] = dsRes.Tables[0].Rows[i]["CREDMAXCER"];
            drRes["SDOACTCER"] = dsRes.Tables[0].Rows[i]["SDOACTCER"];
            drRes["MONTOCER"] = dsRes.Tables[0].Rows[i]["MONTOCER"];

            dtRes.Rows.Add(drRes);
            dsBuro.Tables["dtResumen"].ImportRow(drRes);
        }

        filasCons = dsCons.Tables[0].Rows.Count;

        for (i = 0; i < filasCons; i++)
        {
            DataRow drCons = dtCons.NewRow();
            drCons["FECREGCONS"] = dsCons.Tables[0].Rows[i]["FECREGCONS"];
            drCons["IDBURO"] = dsCons.Tables[0].Rows[i]["IDBURO"];
            drCons["CVEOTORGA"] = dsCons.Tables[0].Rows[i]["CDGOTOR"];
            drCons["NOMOTORGA"] = dsCons.Tables[0].Rows[i]["NOMOTOR"];
            drCons["TELOTORGA"] = dsCons.Tables[0].Rows[i]["TELOTOR"];
            drCons["TIPOCONTRATO"] = dsCons.Tables[0].Rows[i]["TIPOCONT"];
            drCons["DESCTIPOCONT"] = dsCons.Tables[0].Rows[i]["DESCTIPOCONT"];
            drCons["CVEMONEDA"] = dsCons.Tables[0].Rows[i]["CDGMO"];
            drCons["IMPORTE"] = dsCons.Tables[0].Rows[i]["IMPORTE"];
            drCons["TIPORESPONS"] = dsCons.Tables[0].Rows[i]["TIPORESP"];
            drCons["DESCTIPORESP"] = dsCons.Tables[0].Rows[i]["DESCTIPORESP"];
            drCons["CONSNUEVO"] = dsCons.Tables[0].Rows[i]["CONSNUEVO"];
            drCons["RESFINAL"] = dsCons.Tables[0].Rows[i]["RESFINAL"];

            dtCons.Rows.Add(drCons);
            dsBuro.Tables["dtConsulta"].ImportRow(drCons);
        }

        CargaReporte("ConsultaBuro.rpt");
        grcReporte.SetDataSource(dsBuro);
        DisplayPdf(grcReporte);
    }

    private void LlenaConsultaRepCredito(string sic, DataSet dsInfo, DataSet dsNom, DataSet dsDir, DataSet dsCta, DataSet dsRes, DataSet dsCons,
                                         DataSet dsMens, DataSet dsSco, string nomUsuario)
    {
        dsBuro.dtNombreDataTable dtN = new dsBuro.dtNombreDataTable();
        DataTable dtNom = (DataTable)dtN;

        dsBuro.dtDireccionDataTable dtD = new dsBuro.dtDireccionDataTable();
        DataTable dtDir = (DataTable)dtD;

        dsBuro.dtCuentaDataTable dtC = new dsBuro.dtCuentaDataTable();
        DataTable dtCta = (DataTable)dtC;

        dsBuro.dtResumenDataTable dtR = new dsBuro.dtResumenDataTable();
        DataTable dtRes = (DataTable)dtR;

        dsBuro.dtConsultaDataTable dtCon = new dsBuro.dtConsultaDataTable();
        DataTable dtCons = (DataTable)dtCon;

        dsBuro.dtMensajeDataTable dtM = new dsBuro.dtMensajeDataTable();
        DataTable dtMens = (DataTable)dtM;

        dsBuro.dtScoreDataTable dtS = new dsBuro.dtScoreDataTable();
        DataTable dtSco = (DataTable)dtS;

        dsBuro.dtInfoRepDataTable dtInfo = new dsBuro.dtInfoRepDataTable();
        DataTable dtReporteInfo = (DataTable)dtInfo;

        DataSet dsRep = new dsBuro();

        int i;
        int filasDir;
        int filasCta;
        int filasRes;
        int filasCons;
        int filasMens;
        DateTime fec;

        DataRow drInfo = dtInfo.NewRow();
        drInfo["USUARIO"] = nomUsuario;
        drInfo["TITULO"] = dsInfo.Tables[0].Rows[0]["TITULO"];
        drInfo["FEC_CONS"] = dsInfo.Tables[0].Rows[0]["FECHACONS"];
        drInfo["FEC_VIG"] = dsInfo.Tables[0].Rows[0]["FECHAVIG"];
        drInfo["FEC_IMP"] = dsInfo.Tables[0].Rows[0]["FECHAIMP"];
        drInfo["HORA_IMP"] = dsInfo.Tables[0].Rows[0]["HORAIMP"];
        drInfo["REGISTROS"] = dsInfo.Tables[0].Rows[0]["REGISTROS"];
        drInfo["NUM_CONS"] = dsInfo.Tables[0].Rows[0]["NUMCONS"];
        drInfo["FEC_CRED"] = dsInfo.Tables[0].Rows[0]["FECCRED"];
        drInfo["MONTO_CRED"] = dsInfo.Tables[0].Rows[0]["MONTOCRED"];
        dtInfo.Rows.Add(drInfo);
        dsRep.Tables["dtInfoRep"].ImportRow(drInfo);


        DataRow drNom = dtNom.NewRow();
        drNom["APATERNO"] = dsNom.Tables[0].Rows[0]["PRIMAPE"];
        drNom["AMATERNO"] = dsNom.Tables[0].Rows[0]["SEGAPE"];
        drNom["AOPCIONAL"] = dsNom.Tables[0].Rows[0]["OPCAPE"];
        drNom["PNOMBRE"] = dsNom.Tables[0].Rows[0]["NOMBRE1"];
        drNom["SNOMBRE"] = dsNom.Tables[0].Rows[0]["NOMBRE2"];
        drNom["FECNAC"] = DateTime.TryParse(dsNom.Tables[0].Rows[0]["FECNAC"].ToString(), out fec) ? Convert.ToDateTime(dsNom.Tables[0].Rows[0]["FECNAC"]).ToString("dd-MMM-yyyy").ToUpper() : "";
        drNom["RFC"] = dsNom.Tables[0].Rows[0]["RFC"];
        drNom["CURP"] = dsNom.Tables[0].Rows[0]["CURP"];
        drNom["PREFPERS"] = dsNom.Tables[0].Rows[0]["PREFPERS"];
        drNom["SUBFIJO"] = dsNom.Tables[0].Rows[0]["SUBFIJO"];
        drNom["NAC"] = dsNom.Tables[0].Rows[0]["NAC"];
        drNom["RESID"] = dsNom.Tables[0].Rows[0]["RESID"];
        drNom["LICENCIA"] = dsNom.Tables[0].Rows[0]["LICENCIA"];
        drNom["EDOCIVIL"] = dsNom.Tables[0].Rows[0]["EDOCIVIL"];
        drNom["SEXO"] = dsNom.Tables[0].Rows[0]["SEXO"];
        drNom["CEDULA"] = dsNom.Tables[0].Rows[0]["CEDULA"];
        drNom["IFE"] = dsNom.Tables[0].Rows[0]["IFE"];
        drNom["IMPOTROPAIS"] = dsNom.Tables[0].Rows[0]["IMPOTROPAI"];
        drNom["OTROPAIS"] = dsNom.Tables[0].Rows[0]["CDGPAI"];
        drNom["DEPEND"] = dsNom.Tables[0].Rows[0]["DEPEND"];
        drNom["EDADDEPEND"] = dsNom.Tables[0].Rows[0]["EDADDEPEND"];
        drNom["FECINFODEPEND"] = dsNom.Tables[0].Rows[0]["FECINFODEPEND"];
        drNom["FECDEFUNC"] = dsNom.Tables[0].Rows[0]["FECDEFUNC"];
        dtNom.Rows.Add(drNom);
        dsRep.Tables["dtNombre"].ImportRow(drNom);

        filasDir = dsDir.Tables[0].Rows.Count;

        for (i = 0; i < filasDir; i++)
        {
            DataRow drDir = dtDir.NewRow();
            drDir["NUM"] = i + 1;
            drDir["DIR1"] = dsDir.Tables[0].Rows[i]["DIR1"];
            drDir["DIR2"] = dsDir.Tables[0].Rows[i]["DIR2"];
            drDir["COLONIA"] = dsDir.Tables[0].Rows[i]["COLONIA"];
            drDir["MUNICIPIO"] = dsDir.Tables[0].Rows[i]["MUNICIPIO"];
            drDir["CIUDAD"] = dsDir.Tables[0].Rows[i]["CIUDAD"];
            drDir["ESTADO"] = dsDir.Tables[0].Rows[i]["ESTADO"];
            drDir["CODPOSTAL"] = dsDir.Tables[0].Rows[i]["CDGPOSTAL"];
            drDir["FECRESID"] = dsDir.Tables[0].Rows[i]["FECRESID"];
            drDir["TELEFONO"] = dsDir.Tables[0].Rows[i]["TELEFONO"];
            drDir["EXT"] = dsDir.Tables[0].Rows[i]["EXT"];
            drDir["FAX"] = dsDir.Tables[0].Rows[i]["FAX"];
            drDir["TIPODOM"] = dsDir.Tables[0].Rows[i]["TIPODOM"];
            drDir["INDESPDOM"] = dsDir.Tables[0].Rows[i]["INDESPDOM"];
            drDir["FECDOM"] = DateTime.TryParse(dsDir.Tables[0].Rows[i]["FECREPDIR"].ToString(), out fec)?Convert.ToDateTime(dsDir.Tables[0].Rows[i]["FECREPDIR"]).ToString("dd-MMM-yyyy").ToUpper():"";
            dtDir.Rows.Add(drDir);
            dsRep.Tables["dtDireccion"].ImportRow(drDir);
        }

        filasCta = dsCta.Tables[0].Rows.Count;

        for (i = 0; i < filasCta; i++)
        {
            DataRow drCta = dtCta.NewRow();
            drCta["NUM"] = i + 1;
            drCta["FECACT"] = dsCta.Tables[0].Rows[i]["FECACT"];
            drCta["REGIMP"] = dsCta.Tables[0].Rows[i]["REGIMP"];
            drCta["CVEOTORGA"] = dsCta.Tables[0].Rows[i]["CDGOTOR"];
            drCta["NOMOTORGA"] = dsCta.Tables[0].Rows[i]["NOMOTOR"];
            drCta["IDSIC"] = dsCta.Tables[0].Rows[i]["SIC"];
            drCta["CUENTA"] = dsCta.Tables[0].Rows[i]["CUENTA"];
            drCta["TIPORESPONS"] = dsCta.Tables[0].Rows[i]["TIPORESP"];
            drCta["DESCTIPORESP"] = dsCta.Tables[0].Rows[i]["DESCTIPORESP"].ToString().ToUpper();
            drCta["TIPOCUENTA"] = dsCta.Tables[0].Rows[i]["TIPOCTA"];
            drCta["DESCTIPOCTA"] = dsCta.Tables[0].Rows[i]["DESCTIPOCTA"].ToString().ToUpper();
            drCta["TIPOCONTRATO"] = dsCta.Tables[0].Rows[i]["TIPOCONT"];
            drCta["DESCTIPOCONT"] = dsCta.Tables[0].Rows[i]["DESCTIPOCRED"].ToString().ToUpper();
            drCta["CVEMONEDA"] = "MN";
            drCta["VALORACTIVO"] = dsCta.Tables[0].Rows[i]["VALACT"];
            drCta["NUMPAGOS"] = dsCta.Tables[0].Rows[i]["NUMPAGOS"];
            drCta["FRECPAGOS"] = dsCta.Tables[0].Rows[i]["FRECPAGOS"];
            drCta["DESCFRECPAGOS"] = dsCta.Tables[0].Rows[i]["DESCFRECPAGOS"].ToString().ToUpper();
            drCta["MONTOPAGAR"] = dsCta.Tables[0].Rows[i]["MONTO"];
            drCta["FECCUENTA"] = dsCta.Tables[0].Rows[i]["FECINICTA"];
            drCta["FECULTPAGO"] = dsCta.Tables[0].Rows[i]["FECULTPAGO"];
            drCta["FECULTCOMPRA"] = dsCta.Tables[0].Rows[i]["FECULTCOMP"];
            drCta["FECCIERRECTA"] = dsCta.Tables[0].Rows[i]["FECFINCTA"];
            drCta["FECREP"] = dsCta.Tables[0].Rows[i]["FECREP"];
            drCta["MODOREP"] = dsCta.Tables[0].Rows[i]["MODOREP"];
            drCta["FECCTACERO"] = dsCta.Tables[0].Rows[i]["FECSINSALDO"];
            drCta["ESTATUS"] = dsCta.Tables[0].Rows[i]["ESTATUS"];
            drCta["GARANTIA"] = dsCta.Tables[0].Rows[i]["GARANTIA"];
            drCta["CREDMAX"] = dsCta.Tables[0].Rows[i]["CREDMAX"];
            drCta["SALDOACT"] = dsCta.Tables[0].Rows[i]["SDOACT"];
            drCta["LIMCRED"] = dsCta.Tables[0].Rows[i]["LIMCRD"];
            drCta["SALDOVENC"] = dsCta.Tables[0].Rows[i]["SDOVENC"];
            drCta["PAGOSVENC"] = dsCta.Tables[0].Rows[i]["PAGOSVENC"];
            drCta["FORMAPAGO"] = dsCta.Tables[0].Rows[i]["MOP"];
            drCta["DESCFORMAPAGO"] = dsCta.Tables[0].Rows[i]["DESCMOP"].ToString().ToUpper();
            drCta["HISTPAGO"] = dsCta.Tables[0].Rows[i]["HISTPAGOS"];
            drCta["FECRECHISTPAGO"] = dsCta.Tables[0].Rows[i]["FECRECHPAGOS"];
            drCta["FECANTHISTPAGO"] = dsCta.Tables[0].Rows[i]["FECANTHPAGOS"];
            drCta["CVEOBS"] = dsCta.Tables[0].Rows[i]["CDGOBS"];
            drCta["DESCCVEOBS"] = dsCta.Tables[0].Rows[i]["DESCCDGOBS"];
            drCta["TOTALPAGOS"] = dsCta.Tables[0].Rows[i]["PAGOSREP"];
            drCta["TOTALPAGOSMOP02"] = dsCta.Tables[0].Rows[i]["MOP02"];
            drCta["TOTALPAGOSMOP03"] = dsCta.Tables[0].Rows[i]["MOP03"];
            drCta["TOTALPAGOSMOP04"] = dsCta.Tables[0].Rows[i]["MOP04"];
            drCta["TOTALPAGOSMOP05"] = dsCta.Tables[0].Rows[i]["MOP05"];
            drCta["IMPSALDOMORA"] = dsCta.Tables[0].Rows[i]["IMPMOR"];
            drCta["FECSALDOMORA"] = DateTime.TryParse(dsCta.Tables[0].Rows[i]["FECIMPMOR"].ToString(), out fec) ? Convert.ToDateTime(dsCta.Tables[0].Rows[i]["FECIMPMOR"]).ToString("dd-MMM-yyyy").ToUpper() : "";
            drCta["MOPHISTORICO"] = dsCta.Tables[0].Rows[i]["MOPMOR"];
            drCta["FECRESTRUC"] = dsCta.Tables[0].Rows[i]["FECREST"];
            drCta["ANIO1"] = dsCta.Tables[0].Rows[i]["ANIO1"];
            drCta["ANIO2"] = dsCta.Tables[0].Rows[i]["ANIO2"];
            drCta["ANIO3"] = dsCta.Tables[0].Rows[i]["ANIO3"];
            drCta["MES1"] = dsCta.Tables[0].Rows[i]["MES1"];
            drCta["MES2"] = dsCta.Tables[0].Rows[i]["MES2"];
            drCta["MES3"] = dsCta.Tables[0].Rows[i]["MES3"];
            drCta["SEGMES1"] = dsCta.Tables[0].Rows[i]["SEGMES1"];
            drCta["SEGMES2"] = dsCta.Tables[0].Rows[i]["SEGMES2"];
            drCta["SEGMES3"] = dsCta.Tables[0].Rows[i]["SEGMES3"];

            dtCta.Rows.Add(drCta);
            dsRep.Tables["dtCuenta"].ImportRow(drCta);
        }

        filasRes = dsRes.Tables[0].Rows.Count;

        for (i = 0; i < filasRes; i++)
        {
            DataRow drRes = dtRes.NewRow();
            if (sic == "BC")
            {
                drRes["MOP"] = dsRes.Tables[0].Rows[i]["MOP"];
                drRes["NUMCER"] = dsRes.Tables[0].Rows[i]["NUMCER"];
                drRes["LIMCER"] = dsRes.Tables[0].Rows[i]["LIMCER"];         
            }
            else if (sic == "CC")
            {
                drRes["MOP"] = dsRes.Tables[0].Rows[i]["MOP"];
                drRes["PRODUCTO"] = dsRes.Tables[0].Rows[i]["DESCTIPOCONT"];
            }
            drRes["NUMABR"] = dsRes.Tables[0].Rows[i]["NUMABR"];
            drRes["LIMABR"] = dsRes.Tables[0].Rows[i]["LIMABR"];
            drRes["CREDMAXABR"] = dsRes.Tables[0].Rows[i]["CREDMAXABR"];
            drRes["SDOACTABR"] = dsRes.Tables[0].Rows[i]["SDOACTABR"];
            drRes["SDOVENCABR"] = dsRes.Tables[0].Rows[i]["SDOVENCABR"];
            drRes["MONTOABR"] = dsRes.Tables[0].Rows[i]["MONTOABR"];
            drRes["CREDMAXCER"] = dsRes.Tables[0].Rows[i]["CREDMAXCER"];
            drRes["SDOACTCER"] = dsRes.Tables[0].Rows[i]["SDOACTCER"];
            drRes["MONTOCER"] = dsRes.Tables[0].Rows[i]["MONTOCER"];

            dtRes.Rows.Add(drRes);
            dsRep.Tables["dtResumen"].ImportRow(drRes);
        }

        filasCons = dsCons.Tables[0].Rows.Count;

        for (i = 0; i < filasCons; i++)
        {
            DataRow drCons = dtCons.NewRow();
            drCons["FECREGCONS"] = dsCons.Tables[0].Rows[i]["FECREGCONS"];
            drCons["IDBURO"] = dsCons.Tables[0].Rows[i]["IDBURO"];
            drCons["CVEOTORGA"] = dsCons.Tables[0].Rows[i]["CDGOTOR"];
            drCons["NOMOTORGA"] = dsCons.Tables[0].Rows[i]["NOMOTOR"];
            drCons["TELOTORGA"] = dsCons.Tables[0].Rows[i]["TELOTOR"];
            drCons["TIPOCONTRATO"] = dsCons.Tables[0].Rows[i]["TIPOCONT"];
            drCons["DESCTIPOCONT"] = dsCons.Tables[0].Rows[i]["DESCTIPOCONT"];
            drCons["CVEMONEDA"] = dsCons.Tables[0].Rows[i]["CDGMO"];
            drCons["IMPORTE"] = dsCons.Tables[0].Rows[i]["IMPORTE"];
            drCons["TIPORESPONS"] = dsCons.Tables[0].Rows[i]["TIPORESP"];
            drCons["DESCTIPORESP"] = dsCons.Tables[0].Rows[i]["DESCTIPORESP"];
            drCons["CONSNUEVO"] = dsCons.Tables[0].Rows[i]["CONSNUEVO"];
            drCons["RESFINAL"] = dsCons.Tables[0].Rows[i]["RESFINAL"];

            dtCons.Rows.Add(drCons);
            dsRep.Tables["dtConsulta"].ImportRow(drCons);
        }

        filasMens = dsMens.Tables[0].Rows.Count;

        for (i = 0; i < filasMens; i++)
        {
            DataRow drMens = dtMens.NewRow();
            drMens["TIPOMENS"] = dsMens.Tables[0].Rows[i]["TIPOMENS"];
            drMens["LEYENDA"] = dsMens.Tables[0].Rows[i]["LEYENDA"];
            drMens["DESCTIPOMENS"] = dsMens.Tables[0].Rows[i]["DESCTIPOMENS"];
            drMens["DESCLEYENDA"] = dsMens.Tables[0].Rows[i]["DESCLEYENDA"];
            
            dtMens.Rows.Add(drMens);
            dsRep.Tables["dtMensaje"].ImportRow(drMens);
        }

        if (sic == "CC")
        {
            if (dsSco.Tables.Count > 0)
            {
                if (dsSco.Tables[0].Rows.Count > 0)
                {
                    DataRow drSco = dtSco.NewRow();
                    if (dsSco.Tables[0].Columns.Contains("NOMBRESCORE")) drSco["NOMBRESCORE"] = dsSco.Tables[0].Rows[0]["NOMBRESCORE"]; else drSco["NOMBRESCORE"] = "";
                    if (dsSco.Tables[0].Columns.Contains("CODIGO")) drSco["CODIGO"] = dsSco.Tables[0].Rows[0]["CODIGO"]; else drSco["CODIGO"] = "";
                    if (dsSco.Tables[0].Columns.Contains("VALOR")) drSco["VALOR"] = dsSco.Tables[0].Rows[0]["VALOR"]; else drSco["VALOR"] = "";
                    if (dsSco.Tables[0].Columns.Contains("RAZON1")) drSco["RAZON1"] = dsSco.Tables[0].Rows[0]["RAZON1"]; else drSco["RAZON1"] = "";
                    if (dsSco.Tables[0].Columns.Contains("RAZON2")) drSco["RAZON2"] = dsSco.Tables[0].Rows[0]["RAZON2"]; else drSco["RAZON2"] = "";
                    if (dsSco.Tables[0].Columns.Contains("RAZON3")) drSco["RAZON3"] = dsSco.Tables[0].Rows[0]["RAZON3"]; else drSco["RAZON3"] = "";
                    if (dsSco.Tables[0].Columns.Contains("RAZON4")) drSco["RAZON4"] = dsSco.Tables[0].Rows[0]["RAZON4"]; else drSco["RAZON4"] = "";
                    dtSco.Rows.Add(drSco);
                    dsRep.Tables["dtScore"].ImportRow(drSco);

                    CargaReporte("ConsultaRepCredito_Gauge.rpt");
                }
                else
                    CargaReporte("ConsultaRepCredito.rpt");
            }
            else
                CargaReporte("ConsultaRepCredito.rpt");

        }
        else if(sic == "BC")
            CargaReporte("ConsultaBuro.rpt");
        grcReporte.SetDataSource(dsRep);
        DisplayPdf(grcReporte);
    }

    private void LlenaContrato(DataSet dsContrato, DataSet dsCliente)
    {
        dsDocumentos.dtContratoDataTable dt = new dsDocumentos.dtContratoDataTable();
        DataTable dtRepContrato = (DataTable)dt;

        dsDocumentos.dtClienteDataTable dtC = new dsDocumentos.dtClienteDataTable();
        DataTable dtRepCliente = (DataTable)dtC;

        DataSet dsRepContrato = new dsDocumentos();

        int i;
        int contFilas;
        int plazo;
        string cat = string.Empty;
        string cuenta = string.Empty;
        string tipoProd = string.Empty;
        string periodo =string.Empty;
        decimal cantEntre;
        decimal pagoParc;
        decimal totalPagar;

        contFilas = dsContrato.Tables[0].Rows.Count;
        tipoProd = dsContrato.Tables[0].Rows[0]["TIPO_PROD"].ToString();
        periodo = dsContrato.Tables[0].Rows[0]["PER_SINGULAR"].ToString();
        plazo = Convert.ToInt32(dsContrato.Tables[0].Rows[0]["PLAZO"].ToString());
        cantEntre = Convert.ToDecimal(dsContrato.Tables[0].Rows[0]["CANTIDAD_NUMERO"].ToString().Replace("$", ""));
        pagoParc = Convert.ToDecimal(dsContrato.Tables[0].Rows[0]["PARC_SIN_IVA"].ToString().Replace("$", ""));
        totalPagar = Convert.ToDecimal(dsContrato.Tables[0].Rows[0]["TOTAL_SIN_IVA"].ToString().Replace("$", ""));

        cat = Funciones.GeneraCAT(cantEntre, totalPagar, pagoParc, plazo, periodo);

        for (i = 0; i < contFilas; i++)
        {
            DataRow drContrato = dtRepContrato.NewRow();
            drContrato["CODIGO_GPO"] = dsContrato.Tables[0].Rows[i]["CODIGO_GPO"];
            drContrato["NOMBRE_GPO"] = dsContrato.Tables[0].Rows[i]["NOMBRE_GPO"];
            drContrato["MUNICIPIO_CTE"] = dsContrato.Tables[0].Rows[i]["MUNICIPIO_CTE"];
            drContrato["ENTIDAD_CTE"] = dsContrato.Tables[0].Rows[i]["ENTIDAD_CTE"];
            drContrato["FECHA_INICIO"] = dsContrato.Tables[0].Rows[i]["FECHA_INICIO"];
            drContrato["FECHA_PAGO"] = dsContrato.Tables[0].Rows[i]["FECHA_PAGO"];
            drContrato["TASA_CTE"] = dsContrato.Tables[0].Rows[i]["TASA_CTE"];
            drContrato["TASA_ANUAL"] = dsContrato.Tables[0].Rows[i]["TASA_ANUAL"];
            drContrato["PLAZO"] = dsContrato.Tables[0].Rows[i]["PLAZO"];
            drContrato["CANTIDAD_NUMERO"] = dsContrato.Tables[0].Rows[i]["CANTIDAD_NUMERO"].ToString().Replace("$","");
            drContrato["CANTIDAD_LETRA"] = Funciones.convierteNumeroaLetra(Convert.ToDecimal(dsContrato.Tables[0].Rows[i]["CANTIDAD_NUMERO"].ToString().Replace("$", "")));
            drContrato["FECHA_FIN_LETRA"] = dsContrato.Tables[0].Rows[i]["FECHA_FIN_LETRA"];
            drContrato["PRESIDENTE"] = dsContrato.Tables[0].Rows[i]["PRESIDENTE"];
            drContrato["TESORERO"] = dsContrato.Tables[0].Rows[i]["TESORERO"];
            drContrato["SECRETARIO"] = dsContrato.Tables[0].Rows[i]["SECRETARIO"];
            drContrato["SUPERVISOR"] = dsContrato.Tables[0].Rows[i]["SUPERVISOR"];
            drContrato["PAGO_PARCIAL"] = dsContrato.Tables[0].Rows[i]["PAGO_PARCIAL"].ToString().Replace("$", "");
            drContrato["TOTAL_A_PAGAR"] = dsContrato.Tables[0].Rows[i]["TOTAL_A_PAGAR"].ToString().Replace("$", "");
            drContrato["PER_SINGULAR"] = dsContrato.Tables[0].Rows[i]["PER_SINGULAR"];
            drContrato["PER_PLURAL"] = dsContrato.Tables[0].Rows[i]["PER_PLURAL"];
            drContrato["NOM_SUCURSAL"] = dsContrato.Tables[0].Rows[i]["NOM_SUCURSAL"];
            drContrato["NOM_INTEG"] = dsContrato.Tables[0].Rows[i]["NOM_INTEG"];
            drContrato["PAGO_PARCIAL_LETRA"] = Funciones.convierteNumeroaLetra(Convert.ToDecimal(dsContrato.Tables[0].Rows[i]["PAGO_PARCIAL"].ToString().Replace("$", "")));
            drContrato["TOTAL_A_PAGAR_LETRA"] = Funciones.convierteNumeroaLetra(Convert.ToDecimal(dsContrato.Tables[0].Rows[i]["TOTAL_A_PAGAR"].ToString().Replace("$", "")));
            drContrato["CUENTA_BANCARIA"] = dsContrato.Tables[0].Rows[i]["CUENTA_BANCARIA"];
            drContrato["NOM_GERENTE_SUC"] = dsContrato.Tables[0].Rows[i]["NOM_GERENTE_SUC"];
            drContrato["NOM_ASESOR"] = dsContrato.Tables[0].Rows[i]["NOM_ASESOR"];
            drContrato["DIR_SUCURSAL"] = dsContrato.Tables[0].Rows[i]["DIR_SUCURSAL"];
            drContrato["CTAS_BANCOS"] = dsContrato.Tables[0].Rows[i]["CUENTA_BANCARIA"];
            drContrato["IVA"] = dsContrato.Tables[0].Rows[i]["IVA"];
            drContrato["CAT"] = cat;

            dtRepContrato.Rows.Add(drContrato);
            dsRepContrato.Tables["dtContrato"].ImportRow(drContrato);
        }

        contFilas = dsCliente.Tables[0].Rows.Count;

        for (i = 0; i < contFilas; i++)
        {
            DataRow drCliente = dtRepCliente.NewRow();
            drCliente["CDGCL"] = dsCliente.Tables[0].Rows[i]["CODIGO"];
            drCliente["NOMCL"] = dsCliente.Tables[0].Rows[i]["NOMBRE_CL"];
            drCliente["EDOCIVIL"] = dsCliente.Tables[0].Rows[i]["EDO_CIVIL"];
            drCliente["OCUPACION"] = dsCliente.Tables[0].Rows[i]["OCUPACION"];
            drCliente["CALLE"] = dsCliente.Tables[0].Rows[i]["CALLE"];
            drCliente["COLONIA"] = dsCliente.Tables[0].Rows[i]["COLONIA"]; 
            drCliente["LOCALIDAD"] = dsCliente.Tables[0].Rows[i]["LOCALIDAD"];
            drCliente["MUNICIPIO"] = dsCliente.Tables[0].Rows[i]["MUNIC"];
            drCliente["CDGPOSTAL"] = dsCliente.Tables[0].Rows[i]["CODPOSTAL"];
            drCliente["IFE"] = dsCliente.Tables[0].Rows[i]["IFE"];

            dtRepCliente.Rows.Add(drCliente);
            dsRepContrato.Tables["dtCliente"].ImportRow(drCliente);
        }

        if(tipoProd == "01")
            CargaReporte("ContratoComunalFirmas.rpt");
        else if(tipoProd == "02")
            CargaReporte("ContratoAdcicionalFirmas.rpt");
        else if(tipoProd == "03")
            CargaReporte("ContratoIndividualFirmas.rpt");
        grcReporte.SetDataSource(dsRepContrato);
        DisplayPdf(grcReporte);
    }

    private void LlenaControlAhorro(DataSet dsGrupo, DataSet dsAcred, DataSet dsDoc)
    {
        dsControl.dtAcreditadoDataTable dt = new dsControl.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControl.dtGrupoDataTable dtG = new dsControl.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        DataSet dsCont = new dsControl();

        int i;
        int filasAcred;

        string fecInicio = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["INICIO"]).ToString("dd/MM/yyyy");
        string fecFin = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["FECHAFIN"]).ToString("dd/MM/yyyy");
        string doc = dsDoc.Tables[0].Rows[0]["ARCHIVO"].ToString();

        DataRow drGrupo = dtGrupo.NewRow();
        drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
        drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
        drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
        drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORDINACION"];
        drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["NOMBREASESOR"];
        drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["NOMBREGERENTE"];
        drGrupo["DIA_INICIO"] = fecInicio.Substring(0, 2);
        drGrupo["MES_INICIO"] = fecInicio.Substring(3, 2);
        drGrupo["ANIO_INICIO"] = fecInicio.Substring(6, 4);
        drGrupo["DIA_FIN"] = fecFin.Substring(0, 2);
        drGrupo["MES_FIN"] = fecFin.Substring(3, 2);
        drGrupo["ANIO_FIN"] = fecFin.Substring(6, 4);

        dtGrupo.Rows.Add(drGrupo);
        dsCont.Tables["dtGrupo"].ImportRow(drGrupo);

        filasAcred = dsAcred.Tables[0].Rows.Count;

        for (i = 0; i < filasAcred; i++)
        {
            DataRow drAcred = dtAcred.NewRow();
            drAcred["NUMERO"] = i + 1;
            drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
            drAcred["TOTAL_PAGAR"] = dsAcred.Tables[0].Rows[i]["GARANTIA"];
            dtAcred.Rows.Add(drAcred);
            dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
        }

        CargaReporte(doc);
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaControlAsistencia(DataSet dsGrupo, DataSet dsAcred, DataSet dsDoc)
    {
        dsControl.dtAcreditadoDataTable dt = new dsControl.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControl.dtGrupoDataTable dtG = new dsControl.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        DataSet dsCont = new dsControl();

        int i;
        int filasAcred;

        string fecInicio = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["INICIO"]).ToString("dd/MM/yyyy");
        string fecFin = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["FECHAFIN"]).ToString("dd/MM/yyyy");
        string doc = dsDoc.Tables[0].Rows[0]["ARCHIVO"].ToString();

        DataRow drGrupo = dtGrupo.NewRow();
        drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
        drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
        drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
        drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORDINACION"];
        drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["NOMBREASESOR"];
        drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["NOMBREGERENTE"];
        drGrupo["DIA_INICIO"] = fecInicio.Substring(0, 2);
        drGrupo["MES_INICIO"] = fecInicio.Substring(3, 2);
        drGrupo["ANIO_INICIO"] = fecInicio.Substring(6, 4);
        drGrupo["DIA_FIN"] = fecFin.Substring(0, 2);
        drGrupo["MES_FIN"] = fecFin.Substring(3, 2);
        drGrupo["ANIO_FIN"] = fecFin.Substring(6, 4);

        dtGrupo.Rows.Add(drGrupo);
        dsCont.Tables["dtGrupo"].ImportRow(drGrupo);

        filasAcred = dsAcred.Tables[0].Rows.Count;

        for (i = 0; i < filasAcred; i++)
        {
            DataRow drAcred = dtAcred.NewRow();
            drAcred["NUMERO"] = i + 1;
            drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
            dtAcred.Rows.Add(drAcred);
            dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
        }

        CargaReporte(doc);
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaControlPagos(string id, DataSet ds, DataSet dsGrupo, DataSet dsAcred, DateTime fecha, string nomUsuario)
    {
        dsControlEmp.dtAcreditadoDataTable dt = new dsControlEmp.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControlEmp.dtGrupoDataTable dtG = new dsControlEmp.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        dsControlEmp.dtInfoRepDataTable dtInfo = new dsControlEmp.dtInfoRepDataTable();
        DataTable dtReporteInfo = (DataTable)dtInfo;

        DataSet dsCont = new dsControlEmp();

        int i;
        int contFilas = 0;

        if (id == "30")
        {
            contFilas = dsGrupo.Tables[0].Rows.Count;

            for (i = 0; i < contFilas; i++)
            {
                DataRow drGrupo = dtGrupo.NewRow();
                drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[i]["CDGNS"];
                drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[i]["GRUPO"];
                drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[i]["CICLO"];
                drGrupo["INICIO"] = dsGrupo.Tables[0].Rows[i]["FINICIO"];
                drGrupo["FIN"] = dsGrupo.Tables[0].Rows[i]["FFIN"];
                drGrupo["CANTENTRE"] = dsGrupo.Tables[0].Rows[i]["CANTENTRE"];
                drGrupo["PAGO_SEM"] = dsGrupo.Tables[0].Rows[i]["PAGOSEM"];
                drGrupo["PAGO_EXT"] = dsGrupo.Tables[0].Rows[i]["PAGOEXT"];
                drGrupo["APORT_GPO"] = dsGrupo.Tables[0].Rows[i]["APORTA"];
                drGrupo["TOTAL"] = dsGrupo.Tables[0].Rows[i]["TOTAL"];
                drGrupo["ABONO_BANCOS"] = dsGrupo.Tables[0].Rows[i]["TOTAL_PAGADO"];
                drGrupo["DIFERENCIA"] = dsGrupo.Tables[0].Rows[i]["DIFERENCIA"];
                drGrupo["DEV_GPO"] = dsGrupo.Tables[0].Rows[i]["DEV_GPO"];
                drGrupo["AHORRO_GPO"] = dsGrupo.Tables[0].Rows[i]["AHORRO_GPO"];
                drGrupo["MULTA_GPO"] = dsGrupo.Tables[0].Rows[i]["MULTA_GPO"];
                drGrupo["PARCIALIDAD"] = dsGrupo.Tables[0].Rows[i]["PARCIALIDAD"];
                drGrupo["MORA_TOTAL"] = dsGrupo.Tables[0].Rows[i]["MORA_TOTAL"];
                drGrupo["SALDO_TOTAL"] = dsGrupo.Tables[0].Rows[i]["SALDO_TOTAL"];
                drGrupo["DIAS_MORA"] = dsGrupo.Tables[0].Rows[i]["DIAS_MORA"];  
                drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[i]["COORD"];
                drGrupo["REGION"] = dsGrupo.Tables[0].Rows[i]["REGION"];
                drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[i]["ASESOR"];
                dtGrupo.Rows.Add(drGrupo);
                dsCont.Tables["dtGrupo"].ImportRow(drGrupo);
            }
        }

        else if (id == "31")
        {
            contFilas = dsAcred.Tables[0].Rows.Count;

            DataRow drGrupo = dtGrupo.NewRow();
            drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
            drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
            drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
            drGrupo["INICIO"] = dsGrupo.Tables[0].Rows[0]["FINICIO"];
            drGrupo["FIN"] = dsGrupo.Tables[0].Rows[0]["FFIN"];
            drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORD"];
            drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["ASESOR"];
            drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["GERENTE"];
            dtGrupo.Rows.Add(drGrupo);
            dsCont.Tables["dtGrupo"].ImportRow(drGrupo);

            for (i = 0; i < contFilas; i++)
            {
                DataRow drAcred = dtAcred.NewRow();
                drAcred["COD_GRUPO"] = dsAcred.Tables[0].Rows[i]["CDGNS"];
                drAcred["GRUPO"] = dsAcred.Tables[0].Rows[i]["GRUPO"];
                drAcred["COD_ACRED"] = dsAcred.Tables[0].Rows[i]["CDGCL"];
                drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["ACRED"];
                drAcred["CANTENTRE"] = dsAcred.Tables[0].Rows[i]["CANTENTRE"];
                drAcred["PAGO_SEM"] = dsAcred.Tables[0].Rows[i]["PAGO_SEM"];
                drAcred["PAGO_EXT"] = dsAcred.Tables[0].Rows[i]["PAGO_EXT"];
                drAcred["PAGO_REAL"] = dsAcred.Tables[0].Rows[i]["PAGOREAL"];
                drAcred["PAGO_TEO"] = dsAcred.Tables[0].Rows[i]["PAGOCOMP"];
                drAcred["DIFERENCIA"] = dsAcred.Tables[0].Rows[i]["DIFERENCIA"];
                drAcred["APORT"] = dsAcred.Tables[0].Rows[i]["APORTACRED"];
                drAcred["AHORRO"] = dsAcred.Tables[0].Rows[i]["AHORROACRED"];
                drAcred["DEVOLUCION"] = dsAcred.Tables[0].Rows[i]["DEVACRED"];
                drAcred["MULTA"] = dsAcred.Tables[0].Rows[i]["MULTAACRED"];
                drAcred["TOTAL_PAGAR"] = dsAcred.Tables[0].Rows[i]["SALDO"];
                drAcred["PARCIALIDAD"] = dsAcred.Tables[0].Rows[i]["PARCIALIDAD"];
                dtAcred.Rows.Add(drAcred);
                dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
            }
        }

        else if (id == "32")
        {
            contFilas = dsAcred.Tables[0].Rows.Count;

            DataRow drGrupo = dtGrupo.NewRow();
            drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
            drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
            drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
            drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
            drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
            drGrupo["INICIO"] = dsGrupo.Tables[0].Rows[0]["FINICIO"];
            drGrupo["FIN"] = dsGrupo.Tables[0].Rows[0]["FFIN"];
            drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORD"];
            drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["ASESOR"];
            drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["GERENTE"];
            dtGrupo.Rows.Add(drGrupo);
            dsCont.Tables["dtGrupo"].ImportRow(drGrupo);

            for (i = 0; i < contFilas; i++)
            {
                DataRow drAcred = dtAcred.NewRow();
                drAcred["GRUPO"] = dsAcred.Tables[0].Rows[i]["GRUPO"];
                drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["NOMBRE_CL"];
                drAcred["COD_ACRED"] = dsAcred.Tables[0].Rows[i]["CDGCL"];
                drAcred["CANTENTRE"] = dsAcred.Tables[0].Rows[i]["CANTENTRE"];
                drAcred["FECHA"] = dsAcred.Tables[0].Rows[i]["FPAGO"];
                drAcred["TIPO"] = dsAcred.Tables[0].Rows[i]["TIPOREG"];
                drAcred["PAGO_REAL"] = dsAcred.Tables[0].Rows[i]["PAGOREAL"];
                drAcred["APORT"] = dsAcred.Tables[0].Rows[i]["APORT"];
                drAcred["AHORRO"] = dsAcred.Tables[0].Rows[i]["AHORRO"];
                drAcred["DEVOLUCION"] = dsAcred.Tables[0].Rows[i]["DEVOLUCION"];
                drAcred["MULTA"] = dsAcred.Tables[0].Rows[i]["MULTA"];
                drAcred["ASISTENCIA"] = dsAcred.Tables[0].Rows[i]["ASIST"];
                dtAcred.Rows.Add(drAcred);
                dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
            }
        }

        DataRow drInfo = dtReporteInfo.NewRow();
        drInfo["FEC_IMP"] = ds.Tables[0].Rows[0]["FECHAIMP"];
        drInfo["HORA_IMP"] = ds.Tables[0].Rows[0]["HORAIMP"];
        drInfo["FECHA"] = fecha.ToString("dd/MM/yyyy");
        drInfo["USUARIO"] = nomUsuario;
        drInfo["REGISTROS"] = contFilas;
        dtReporteInfo.Rows.Add(drInfo);
        dsCont.Tables["dtInfoRep"].ImportRow(drInfo);

        if (id == "30")
            CargaReporte("controlPagosGrupo.rpt");
        else if(id == "31")
            CargaReporte("controlPagosAcum.rpt");
        else if(id == "32")
            CargaReporte("controlPagosAcred.rpt");
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaControlSemanal(DataSet dsGrupo, DataSet dsAcred, DataSet dsDoc, string id)
    {
        dsControl.dtAcreditadoDataTable dt = new dsControl.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControl.dtGrupoDataTable dtG = new dsControl.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        DataSet dsCont = new dsControl();

        int i;
        int filasAcred;

        string fecInicio = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["INICIO"]).ToString("dd/MM/yyyy");
        string fecFin = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["FECHAFIN"]).ToString("dd/MM/yyyy");
        string doc = dsDoc.Tables[0].Rows[0]["ARCHIVO"].ToString();

        DataRow drGrupo = dtGrupo.NewRow();
        //drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
        //drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
        //drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
        //drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORDINACION"];
        //drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["NOMBREASESOR"];
        //drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["NOMBREGERENTE"];
        //drGrupo["DIA_INICIO"] = fecInicio.Substring(0,2);
        //drGrupo["MES_INICIO"] = fecInicio.Substring(3,2);
        //drGrupo["ANIO_INICIO"] = fecInicio.Substring(6,4);
        //drGrupo["DIA_FIN"] = fecFin.Substring(0,2);
        //drGrupo["MES_FIN"] = fecFin.Substring(3,2);
        //drGrupo["ANIO_FIN"] = fecFin.Substring(6,4);

        dtGrupo.Rows.Add(drGrupo);
        dsCont.Tables["dtGrupo"].ImportRow(drGrupo);
        
        filasAcred = dsAcred.Tables[0].Rows.Count;

        for (i = 0; i < 18; i++)
        {
            DataRow drAcred = dtAcred.NewRow();
            drAcred["NUMERO"] = i + 1;
            //drAcred["CARGO"] = dsAcred.Tables[0].Rows[i]["PUESTO"];
            //drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
            //drAcred["CANTENTRE"] = dsAcred.Tables[0].Rows[i]["CANTENTRE"];
            //drAcred["TOTAL_PAGAR"] = dsAcred.Tables[0].Rows[i]["TOTAL_A_PAGAR"];
            //drAcred["PARCIALIDAD"] = dsAcred.Tables[0].Rows[i]["PARCIALIDAD"];
            dtAcred.Rows.Add(drAcred);
            dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
        }

        CargaReporte(doc);
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaControlSemanalEmp(DataSet dsGrupo, DataSet dsAcred)
    {
        dsControlEmp.dtAcreditadoDataTable dt = new dsControlEmp.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControlEmp.dtGrupoDataTable dtG = new dsControlEmp.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        DataSet dsCont = new dsControlEmp();

        int i;
        int j;
        int filasAcred;
        int filasGrupo;

        string fecInicio = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["INICIO"]).ToString("dd/MM/yyyy");
        string fecFin = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["FECHAFIN"]).ToString("dd/MM/yyyy");
        string semAct = string.Empty;
        string semSig = string.Empty;

        filasGrupo = dsGrupo.Tables[0].Rows.Count;

        for (i = 0; i < filasGrupo; i++)
        {
            DataRow drGrupo = dtGrupo.NewRow();
            drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[i]["CDGNS"];
            drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[i]["GRUPO"];
            drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[i]["CICLO"];
            drGrupo["FECHA"] = dsGrupo.Tables[0].Rows[i]["FECHA"];
            drGrupo["SECUENCIA"] = Convert.ToDateTime(dsGrupo.Tables[0].Rows[i]["FECHA"].ToString()).ToString("yyyy/MM/dd") + dsGrupo.Tables[0].Rows[i]["SECUENCIA"].ToString();
            drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[i]["COORDINACION"];
            drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[i]["NOMBREASESOR"];
            drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[i]["NOMBREGERENTE"];
            drGrupo["SEMANA_GPO"] = dsGrupo.Tables[0].Rows[i]["FECHA"].ToString() + dsGrupo.Tables[0].Rows[i]["SECUENCIA"].ToString(); //dsGrupo.Tables[0].Rows[i]["SEMANA"];
            drGrupo["PAGO_TEO_GPO"] = dsGrupo.Tables[0].Rows[i]["PAGOTEO"];
            drGrupo["PAGO_REAL_GPO"] = dsGrupo.Tables[0].Rows[i]["PAGOREAL"];
            drGrupo["APORT_GPO"] = dsGrupo.Tables[0].Rows[i]["APORT"];
            drGrupo["DEV_GPO"] = dsGrupo.Tables[0].Rows[i]["DEVOLUCION"];
            drGrupo["AHORRO_GPO"] = dsGrupo.Tables[0].Rows[i]["AHORRO"];
            drGrupo["MULTA_GPO"] = dsGrupo.Tables[0].Rows[i]["MULTA"];
            drGrupo["MORA"] = dsGrupo.Tables[0].Rows[i]["MORA"];
            drGrupo["OBSERVACION"] = dsGrupo.Tables[0].Rows[i]["OBSERVACION"];
            drGrupo["DIA_INICIO"] = fecInicio.Substring(0, 2);
            drGrupo["MES_INICIO"] = fecInicio.Substring(3, 2);
            drGrupo["ANIO_INICIO"] = fecInicio.Substring(6, 4);
            drGrupo["DIA_FIN"] = fecFin.Substring(0, 2);
            drGrupo["MES_FIN"] = fecFin.Substring(3, 2);
            drGrupo["ANIO_FIN"] = fecFin.Substring(6, 4);
            drGrupo["TIPO"] = dsGrupo.Tables[0].Rows[i]["TIPOREG"];
            dtGrupo.Rows.Add(drGrupo);
            dsCont.Tables["dtGrupo"].ImportRow(drGrupo);
        }

        filasAcred = dsAcred.Tables[0].Rows.Count;
        j = 1;

        for (i = 0; i < filasAcred; i++)
        {
            DataRow drAcred = dtAcred.NewRow();
            drAcred["NUMERO"] = j;
            drAcred["CARGO"] = dsAcred.Tables[0].Rows[i]["PUESTO"];
            drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
            drAcred["CANTENTRE"] = dsAcred.Tables[0].Rows[i]["CANTENTRE"];
            drAcred["TOTAL_PAGAR"] = dsAcred.Tables[0].Rows[i]["TOTAL_A_PAGAR"];
            drAcred["PARCIALIDAD"] = dsAcred.Tables[0].Rows[i]["PARCIALIDAD"];
            drAcred["FECHA"] = dsAcred.Tables[0].Rows[i]["FECHA"];
            drAcred["SECUENCIA"] = Convert.ToDateTime(dsAcred.Tables[0].Rows[i]["FECHA"].ToString()).ToString("yyyy/MM/dd") + dsAcred.Tables[0].Rows[i]["SECUENCIA"].ToString();
            drAcred["SEMANA"] = dsAcred.Tables[0].Rows[i]["FECHA"].ToString() + dsAcred.Tables[0].Rows[i]["SECUENCIA"].ToString(); //dsAcred.Tables[0].Rows[i]["SEMANA"];
            drAcred["PAGO_REAL"] = dsAcred.Tables[0].Rows[i]["PAGOREAL"];
            drAcred["APORT"] = dsAcred.Tables[0].Rows[i]["APORT"];
            drAcred["DEVOLUCION"] = dsAcred.Tables[0].Rows[i]["DEVOLUCION"];
            drAcred["AHORRO"] = dsAcred.Tables[0].Rows[i]["AHORRO"];
            drAcred["MULTA"] = dsAcred.Tables[0].Rows[i]["MULTA"];
            drAcred["ASISTENCIA"] = dsAcred.Tables[0].Rows[i]["ASIST"];
            dtAcred.Rows.Add(drAcred);
            dsCont.Tables["dtAcreditado"].ImportRow(drAcred);

            if (i + 1 < filasAcred)
            {
                semSig = Convert.ToDateTime(dsAcred.Tables[0].Rows[i + 1]["FECHA"].ToString()).ToString("yyyy/MM/dd") + dsAcred.Tables[0].Rows[i + 1]["SECUENCIA"].ToString(); //dsAcred.Tables[0].Rows[i + 1]["SEMANA"].ToString();
            }
            semAct = Convert.ToDateTime(dsAcred.Tables[0].Rows[i]["FECHA"].ToString()).ToString("yyyy/MM/dd") + dsAcred.Tables[0].Rows[i]["SECUENCIA"].ToString(); //dsAcred.Tables[0].Rows[i]["SEMANA"].ToString(); 

            if (semSig != semAct)
            {
                j = 1;
            }
            else
            {
                j++;
            }
        }

        CargaReporte("controlPagosEmp.rpt");
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaControlSemanalImp(DataSet dsGrupo, DataSet dsAcred)
    {
        dsControlExt.dtAcreditadoDataTable dt = new dsControlExt.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsControlExt.dtGrupoDataTable dtG = new dsControlExt.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        DataSet dsCont = new dsControlExt();

        int i;
        int filasAcred;

        string fecInicio = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["INICIO"]).ToString("dd/MM/yyyy");
        string fecFin = Convert.ToDateTime(dsGrupo.Tables[0].Rows[0]["FECHAFIN"]).ToString("dd/MM/yyyy");

        DataRow drGrupo = dtGrupo.NewRow();
        drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
        drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
        drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
        drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORDINACION"];
        drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["NOMBREASESOR"];
        drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["NOMBREGERENTE"];
        drGrupo["DIA_INICIO"] = fecInicio.Substring(0, 2);
        drGrupo["MES_INICIO"] = fecInicio.Substring(3, 2);
        drGrupo["ANIO_INICIO"] = fecInicio.Substring(6, 4);
        drGrupo["DIA_FIN"] = fecFin.Substring(0, 2);
        drGrupo["MES_FIN"] = fecFin.Substring(3, 2);
        drGrupo["ANIO_FIN"] = fecFin.Substring(6, 4);

        dtGrupo.Rows.Add(drGrupo);
        dsCont.Tables["dtGrupo"].ImportRow(drGrupo);

        filasAcred = dsAcred.Tables[0].Rows.Count;

        for (i = 0; i < filasAcred; i++)
        {
            DataRow drAcred = dtAcred.NewRow();
            drAcred["NUMERO"] = i + 1;
            drAcred["CARGO"] = dsAcred.Tables[0].Rows[i]["PUESTO"];
            drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
            drAcred["CANTENTRE"] = dsAcred.Tables[0].Rows[i]["CANTENTRE"];
            drAcred["TOTAL_PAGAR"] = dsAcred.Tables[0].Rows[i]["TOTAL_A_PAGAR"];
            drAcred["PARCIALIDAD"] = dsAcred.Tables[0].Rows[i]["PARCIALIDAD"];

            drAcred["SEM1"] = dsAcred.Tables[0].Rows[i]["SEMANA_1"];
            drAcred["SEM2"] = dsAcred.Tables[0].Rows[i]["SEMANA_2"];
            drAcred["SEM3"] = dsAcred.Tables[0].Rows[i]["SEMANA_3"];
            drAcred["SEM4"] = dsAcred.Tables[0].Rows[i]["SEMANA_4"];
            drAcred["SEM5"] = dsAcred.Tables[0].Rows[i]["SEMANA_5"];
            drAcred["SEM6"] = dsAcred.Tables[0].Rows[i]["SEMANA_6"];
            drAcred["SEM7"] = dsAcred.Tables[0].Rows[i]["SEMANA_7"];
            drAcred["SEM8"] = dsAcred.Tables[0].Rows[i]["SEMANA_8"];
            drAcred["SEM9"] = dsAcred.Tables[0].Rows[i]["SEMANA_9"];
            drAcred["SEM10"] = dsAcred.Tables[0].Rows[i]["SEMANA_10"];
            drAcred["SEM11"] = dsAcred.Tables[0].Rows[i]["SEMANA_11"];
            drAcred["SEM12"] = dsAcred.Tables[0].Rows[i]["SEMANA_12"];
            drAcred["SEM13"] = dsAcred.Tables[0].Rows[i]["SEMANA_13"];
            drAcred["SEM14"] = dsAcred.Tables[0].Rows[i]["SEMANA_14"];
            drAcred["SEM15"] = dsAcred.Tables[0].Rows[i]["SEMANA_15"];
            drAcred["SEM16"] = dsAcred.Tables[0].Rows[i]["SEMANA_16"];

            dtAcred.Rows.Add(drAcred);
            dsCont.Tables["dtAcreditado"].ImportRow(drAcred);
        }

        CargaReporte("controlSemanalImp.rpt");
        grcReporte.SetDataSource(dsCont);
        DisplayPdf(grcReporte);
    }

    private void LlenaEstadoCuenta(DataSet dsEnc, DataSet dsDet)
    {
        dsEdoCuenta.dtEncDataTable dtE = new dsEdoCuenta.dtEncDataTable();
        DataTable dtEnc = (DataTable)dtE;

        dsEdoCuenta.dtDetDataTable dtD = new dsEdoCuenta.dtDetDataTable();
        DataTable dtDet = (DataTable)dtD;

        DataSet dsEdoCta = new dsEdoCuenta();

        int i;
        int plazo;
        int contFilas;
        string cat = string.Empty;
        string periodo = string.Empty;
        decimal cantEntre;
        decimal pagoParc;
        decimal totalPagar;
        
        periodo = dsEnc.Tables[0].Rows[0]["NOMPER"].ToString();
        plazo = Convert.ToInt32(dsEnc.Tables[0].Rows[0]["PLAZO"].ToString());
        cantEntre = Convert.ToDecimal(dsEnc.Tables[0].Rows[0]["CANTENTRE"].ToString().Replace("$", ""));
        totalPagar = Convert.ToDecimal(dsEnc.Tables[0].Rows[0]["TOTALSINIVA"].ToString().Replace("$", ""));

        pagoParc = Math.Round(totalPagar / plazo, 2);
        cat = Funciones.GeneraCAT(cantEntre, totalPagar, pagoParc, plazo, periodo);

        contFilas = dsDet.Tables[0].Rows.Count;

        DataRow drEnc = dtEnc.NewRow();
        drEnc["CDGCLNS"] = dsEnc.Tables[0].Rows[0]["CDGCLNS"];
        drEnc["NOMCLNS"] = dsEnc.Tables[0].Rows[0]["NOMCLNS"];
        drEnc["CDGRG"] = dsEnc.Tables[0].Rows[0]["CDGRG"];
        drEnc["NOMRG"] = dsEnc.Tables[0].Rows[0]["NOMRG"];
        drEnc["CDGCO"] = dsEnc.Tables[0].Rows[0]["CDGCO"];
        drEnc["NOMCO"] = dsEnc.Tables[0].Rows[0]["NOMCO"];
        drEnc["CDGFUNPE"] = dsEnc.Tables[0].Rows[0]["CDGOCPE"];
        drEnc["NOMBRE1"] = dsEnc.Tables[0].Rows[0]["NOMBRE1"];
        drEnc["NOMBRE2"] = dsEnc.Tables[0].Rows[0]["NOMBRE2"];
        drEnc["PRIMAPE"] = dsEnc.Tables[0].Rows[0]["PRIMAPE"];
        drEnc["SEGAPE"] = dsEnc.Tables[0].Rows[0]["SEGAPE"];
        drEnc["CALLE"] = dsEnc.Tables[0].Rows[0]["CALLE"];
        drEnc["NOMCOL"] = dsEnc.Tables[0].Rows[0]["NOMCOL"];
        drEnc["NOMMU"] = dsEnc.Tables[0].Rows[0]["NOMMU"];
        drEnc["NOMLO"] = dsEnc.Tables[0].Rows[0]["NOMLO"];
        drEnc["CICLO"] = dsEnc.Tables[0].Rows[0]["CICLO"];
        drEnc["INICIO"] = dsEnc.Tables[0].Rows[0]["INICIO"];
        drEnc["FIN"] = dsEnc.Tables[0].Rows[0]["FIN"];
        drEnc["PLAZO"] = dsEnc.Tables[0].Rows[0]["PLAZO"];
        drEnc["CANTAUTOR"] = dsEnc.Tables[0].Rows[0]["CANTAUTOR"];
        drEnc["CANTENTRE"] = dsEnc.Tables[0].Rows[0]["CANTENTRE"];
        drEnc["GARANTIA"] = dsEnc.Tables[0].Rows[0]["GARANTIA"];
        drEnc["PERIODICIDAD"] = dsEnc.Tables[0].Rows[0]["PERIODICIDAD"];
        drEnc["SITUACION"] = dsEnc.Tables[0].Rows[0]["SITUA"];
        drEnc["TASA"] = dsEnc.Tables[0].Rows[0]["TASA"];
        drEnc["CDGEM"] = dsEnc.Tables[0].Rows[0]["CDGEM"];
        drEnc["TIPO"] = dsEnc.Tables[0].Rows[0]["TIPO"];
        drEnc["CLNS"] = dsEnc.Tables[0].Rows[0]["CLNS"];
        drEnc["CANTIDAD"] = dsEnc.Tables[0].Rows[0]["INTERES"];
        drEnc["TOTALABONOS"] = dsEnc.Tables[0].Rows[0]["TOTALABONOS"];
        drEnc["TOTALCARGOS"] = dsEnc.Tables[0].Rows[0]["TOTALCARGOS"];
        drEnc["SALDOGL"] = dsEnc.Tables[0].Rows[0]["SALDOGL"];
        drEnc["SALDOREC"] = dsEnc.Tables[0].Rows[0]["SALDOREC"];
        drEnc["DIASMORA"] = dsEnc.Tables[0].Rows[0]["DIASMORA"];
        drEnc["DIASATRASO"] = dsEnc.Tables[0].Rows[0]["DIASATRASO"];
        drEnc["FECHA"] = dsEnc.Tables[0].Rows[0]["FECHA"];
        drEnc["FECHACALC"] = dsEnc.Tables[0].Rows[0]["FECHACALC"];
        drEnc["IVA"] = dsEnc.Tables[0].Rows[0]["IVA"];
        drEnc["TASAANUAL"] = dsEnc.Tables[0].Rows[0]["TASAANUAL"];
        drEnc["CAT"] = cat;
        dtEnc.Rows.Add(drEnc);
        dsEdoCta.Tables["dtEnc"].ImportRow(drEnc);

        for (i = 0; i < contFilas; i++)
        {
            DataRow drDet = dtDet.NewRow();
            drDet["PERIODO"] = dsDet.Tables[0].Rows[i]["PERIODO"];
            drDet["FREALDEP"] = dsDet.Tables[0].Rows[i]["FREALDEP"];
            drDet["PAGOIDEN"] = dsDet.Tables[0].Rows[i]["PAGOIDEN"];
            drDet["CANTIDAD"] = dsDet.Tables[0].Rows[i]["CANTIDAD"];
            drDet["PAGADOCAP"] = dsDet.Tables[0].Rows[i]["PAGADOCAP"];
            drDet["PAGADOINT"] = dsDet.Tables[0].Rows[i]["PAGADOINT"];
            drDet["PAGADOREC"] = dsDet.Tables[0].Rows[i]["PAGADOREC"];
            drDet["CLNS"] = dsDet.Tables[0].Rows[i]["CLNS"];
            drDet["CDGCLNS"] = dsDet.Tables[0].Rows[i]["CDGCLNS"];
            drDet["CDGCL"] = dsDet.Tables[0].Rows[i]["CDGCL"];
            drDet["TIPO"] = dsDet.Tables[0].Rows[i]["TIPO"];
            drDet["IVA"] = dsDet.Tables[0].Rows[i]["IVA"];
            dtDet.Rows.Add(drDet);
            dsEdoCta.Tables["dtDet"].ImportRow(drDet);
        }

        CargaReporte("EstadoCuenta.rpt");
        grcReporte.SetDataSource(dsEdoCta);
        DisplayPdf(grcReporte);
    }

    private void LlenaImpresionCheques(DataSet ds, string cuenta)
    {
        dsCheque.dtChequeDataTable dt = new dsCheque.dtChequeDataTable();
        DataTable dtReporteCheque = (DataTable)dt;

        DataSet dsCheque = new dsCheque();

        int i;
        int contFilas;

        contFilas = ds.Tables[0].Rows.Count;

        for (i = 0; i < contFilas; i++)
        {
            DataRow drCheque = dtReporteCheque.NewRow();
            drCheque["NOMBRE"] = ds.Tables[0].Rows[i]["NOMBREC"];
            drCheque["CANTIDAD"] = ds.Tables[0].Rows[i]["MONTO"];
            drCheque["CANTIDAD_LETRA"] = "(" + Funciones.convierteNumeroaLetra(Convert.ToDecimal(ds.Tables[0].Rows[i]["MONTO"])) + ")";
            drCheque["CUENTA"] = ds.Tables[0].Rows[i]["CDGCB"];
            drCheque["COORDINACION"] = ds.Tables[0].Rows[i]["CDGCO"];
            drCheque["GRUPO"] = ds.Tables[0].Rows[i]["CDGNS"];
            drCheque["CICLO"] = ds.Tables[0].Rows[i]["CICLO"];
            drCheque["NOCHEQUE"] = ds.Tables[0].Rows[i]["NOCHEQUE"];
            drCheque["FECHA_IMP"] = ds.Tables[0].Rows[i]["FECHA"];
            try
            {//Si existe el campo NO_NEGOCIABLE
                drCheque["NO_NEGOCIABLE"] = ds.Tables[0].Rows[i]["NO_NEGOCIABLE"];
            }
            catch (Exception ex)
            {// si no existe en las demas funciones que mandan a llamar esta funcion
                drCheque["NO_NEGOCIABLE"] = "";
            }
          
            dtReporteCheque.Rows.Add(drCheque);
            dsCheque.Tables["dtCheque"].ImportRow(drCheque);
        }

        //IMPRESION DE CHEQUES BANCOMER
        if (cuenta == "06" || cuenta == "07" || cuenta == "08" || cuenta == "10" ||
            cuenta == "11" || cuenta == "13" || cuenta == "14" || cuenta == "15" ||
            cuenta == "16" || cuenta == "17" || cuenta == "18" || cuenta == "20" || 
            cuenta == "21" || cuenta == "22" || cuenta == "23" || cuenta == "24" ||
            cuenta == "25" || cuenta == "26" || cuenta == "27" || cuenta == "32" ||
            cuenta == "33")
            CargaReporte("ChequeBancomerVST.rpt");

        grcReporte.SetDataSource(dsCheque);
        DisplayPdf(grcReporte);
    }

    private void LlenaPagare(DataSet dsDet, DataSet dsAmort)
    {
        dsPagare.dtPagareDataTable dt = new dsPagare.dtPagareDataTable();
        DataTable dtRepPagare = (DataTable)dt;

        dsPagare.dtAmortDataTable dtA = new dsPagare.dtAmortDataTable();
        DataTable dtRepAmort = (DataTable)dtA;

        DataSet dsRepPagare = new dsPagare();
                  
        int i;
        int contFilas;
        string cuenta = string.Empty;
        string tipoProd = string.Empty;

        contFilas = dsDet.Tables[0].Rows.Count;
        tipoProd = dsDet.Tables[0].Rows[0]["TIPO_PROD"].ToString();
        
        for (i = 0; i < contFilas; i++)
        {
            DataRow drPagare = dtRepPagare.NewRow();
            drPagare["TIPO_DOC"] = dsDet.Tables[0].Rows[i]["TIPO_DOC"];
            drPagare["NOMBRE_EMP"] = dsDet.Tables[0].Rows[i]["NOMBRE_EMP"];
            drPagare["CODIGO_EMP"] = dsDet.Tables[0].Rows[i]["CODIGO_EMP"];
            drPagare["CODIGO_CTE"] = dsDet.Tables[0].Rows[i]["CODIGO_CTE"];
            drPagare["nombre_cte"] = dsDet.Tables[0].Rows[i]["nombre_cte"];
            drPagare["apellidos_cte"] = dsDet.Tables[0].Rows[i]["apellidos_cte"];
            drPagare["codigo_gpo"] = dsDet.Tables[0].Rows[i]["codigo_gpo"];
            drPagare["nombre_gpo"] = dsDet.Tables[0].Rows[i]["nombre_gpo"];
            drPagare["calle_cte"] = dsDet.Tables[0].Rows[i]["calle_cte"];
            drPagare["colonia_cte"] = dsDet.Tables[0].Rows[i]["colonia_cte"];
            drPagare["localidad_cte"] = dsDet.Tables[0].Rows[i]["localidad_cte"];
            drPagare["municipio_cte"] = dsDet.Tables[0].Rows[i]["municipio_cte"];
            drPagare["entidad_cte"] = dsDet.Tables[0].Rows[i]["entidad_cte"];
            drPagare["pais_cte"] = dsDet.Tables[0].Rows[i]["pais_cte"];
            drPagare["cp_cte"] = dsDet.Tables[0].Rows[i]["cp_cte"];
            drPagare["fecha_inicio"] = dsDet.Tables[0].Rows[i]["fecha_inicio"];
            drPagare["fecha_fin"] = dsDet.Tables[0].Rows[i]["fecha_fin"];
            drPagare["cantidad_letra"] = Funciones.convierteNumeroaLetra(Convert.ToDecimal(dsDet.Tables[0].Rows[i]["cantidad_numero"].ToString()));
            drPagare["cantidad_numero"] = dsDet.Tables[0].Rows[i]["cantidad_numero"];
            drPagare["tasa_cte"] = dsDet.Tables[0].Rows[i]["tasa_cte"];
            drPagare["nombre_aval1"] = dsDet.Tables[0].Rows[i]["nombre_aval1"];
            drPagare["direc1_aval1"] = dsDet.Tables[0].Rows[i]["direc1_aval1"];
            drPagare["nombre_aval2"] = dsDet.Tables[0].Rows[i]["nombre_aval2"];
            drPagare["direc1_aval2"] = dsDet.Tables[0].Rows[i]["direc1_aval2"];
            drPagare["nombre_aval3"] = dsDet.Tables[0].Rows[i]["nombre_aval3"];
            drPagare["direc1_aval3"] = dsDet.Tables[0].Rows[i]["direc1_aval3"];
            drPagare["nombre_aval4"] = dsDet.Tables[0].Rows[i]["nombre_aval4"];
            drPagare["direc1_aval4"] = dsDet.Tables[0].Rows[i]["direc1_aval4"];
            drPagare["direc2_aval1"] = dsDet.Tables[0].Rows[i]["direc2_aval1"];
            drPagare["direc2_aval2"] = dsDet.Tables[0].Rows[i]["direc2_aval2"];
            drPagare["direc2_aval3"] = dsDet.Tables[0].Rows[i]["direc2_aval3"];
            drPagare["direc2_aval4"] = dsDet.Tables[0].Rows[i]["direc2_aval4"];
            drPagare["no_avales"] = dsDet.Tables[0].Rows[i]["no_avales"];
            drPagare["nombre_com_cte"] = dsDet.Tables[0].Rows[i]["nombre_com_cte"];
            drPagare["ctas_bancos"] = dsDet.Tables[0].Rows[i]["ctas_bancos"];
            drPagare["direccion_suc"] = dsDet.Tables[0].Rows[i]["direccion_suc"];
            drPagare["LOCALIDAD_SUC"] = dsDet.Tables[0].Rows[i]["LOCALIDAD_SUC"];    
            drPagare["municipio_suc"] = dsDet.Tables[0].Rows[i]["municipio_suc"];
            drPagare["pais_suc"] = dsDet.Tables[0].Rows[i]["pais_suc"];
            
            drPagare["plazo_cte"] = dsDet.Tables[0].Rows[i]["plazo_cte"];
            drPagare["periodo_cte"] = dsDet.Tables[0].Rows[i]["periodo_cte"];
            drPagare["cant_ent_gpo"] = dsDet.Tables[0].Rows[i]["cant_ent_gpo"];
            drPagare["cant_ent_gpo_letra"] = dsDet.Tables[0].Rows[i]["cant_ent_gpo_letra"];
            drPagare["ctas_bancos"] = dsDet.Tables[0].Rows[i]["ctas_bancos"];
            drPagare["intereses_gpo"] = dsDet.Tables[0].Rows[i]["intereses_gpo"];
            drPagare["cod_pre_gpo"] = dsDet.Tables[0].Rows[i]["cod_pre_gpo"];

            drPagare["cod_sec_gpo"] = dsDet.Tables[0].Rows[i]["cod_sec_gpo"];
            drPagare["cod_tes_gpo"] = dsDet.Tables[0].Rows[i]["cod_tes_gpo"];
            drPagare["nom_pre_gpo"] = dsDet.Tables[0].Rows[i]["nom_pre_gpo"];
            drPagare["nom_sec_gpo"] = dsDet.Tables[0].Rows[i]["nom_sec_gpo"];
            drPagare["nom_tes_gpo"] = dsDet.Tables[0].Rows[i]["nom_tes_gpo"];
            drPagare["dir_pre_gpo"] = dsDet.Tables[0].Rows[i]["dir_pre_gpo"];
            drPagare["dir_sec_gpo"] = dsDet.Tables[0].Rows[i]["dir_sec_gpo"];
            drPagare["dir_tes_gpo"] = dsDet.Tables[0].Rows[i]["dir_tes_gpo"];
            drPagare["IVA"] = dsDet.Tables[0].Rows[i]["IVA"];

            drPagare["DIRECCION_EMP"] = dsDet.Tables[0].Rows[i]["DIRECCION_EMP"];

            dtRepPagare.Rows.Add(drPagare);
            dsRepPagare.Tables["dtPagare"].ImportRow(drPagare);
        }

        contFilas = dsAmort.Tables[0].Rows.Count;

        for (i = 0; i < contFilas; i++)
        {
            DataRow drAmort = dtRepAmort.NewRow();
            drAmort["NOPAGO"] = dsAmort.Tables[0].Rows[i]["NOPAGO"];
            drAmort["FECHAPAGO"] = dsAmort.Tables[0].Rows[i]["FECHAPAGO"];
            drAmort["IMPORTE"] = dsAmort.Tables[0].Rows[i]["IMPORTE"];
            drAmort["CDG_GRUPO"] = dsAmort.Tables[0].Rows[i]["CDG_GRUPO"];
            drAmort["CDGEM"] = dsAmort.Tables[0].Rows[i]["CDG_EM"];
            drAmort["TIPO_DOC"] = dsAmort.Tables[0].Rows[i]["TIPO_DOC"];
            drAmort["CDGCL"] = dsAmort.Tables[0].Rows[i]["CDGCL"];
            drAmort["CICLO"] = dsAmort.Tables[0].Rows[i]["CICLO"];
            dtRepAmort.Rows.Add(drAmort);
            dsRepPagare.Tables["dtAmort"].ImportRow(drAmort);
        }

        if(tipoProd == "01" || tipoProd == "03")
            CargaReporte("PagareComunalVST.rpt");
        else if(tipoProd == "02")
            CargaReporte("PagareAdicionalVST.rpt");
        grcReporte.SetDataSource(dsRepPagare);
        DisplayPdf(grcReporte);
    }

    private void LlenaPagareGrupal(DataSet dsDet, DataSet dsAval)
    {
        dsPagare.dtPagareDataTable dt = new dsPagare.dtPagareDataTable();
        DataTable dtRepPagare = (DataTable)dt;

        dsPagare.dtAvalDataTable dtA = new dsPagare.dtAvalDataTable();
        DataTable dtRepAval = (DataTable)dtA;

        DataSet dsRepPagare = new dsPagare();

        int i;
        int contFilas;
        string tipoProd = string.Empty;

        contFilas = dsDet.Tables[0].Rows.Count;
        tipoProd = dsDet.Tables[0].Rows[0]["TIPO_PROD"].ToString();

        for (i = 0; i < contFilas; i++)
        {
            DataRow drPagare = dtRepPagare.NewRow();
            drPagare["TIPO_DOC"] = dsDet.Tables[0].Rows[i]["TIPO_DOC"];
            drPagare["NOMBRE_EMP"] = dsDet.Tables[0].Rows[i]["NOMBRE_EMP"];
            drPagare["CODIGO_EMP"] = dsDet.Tables[0].Rows[i]["CODIGO_EMP"];
            drPagare["CODIGO_CTE"] = dsDet.Tables[0].Rows[i]["CODIGO_CTE"];
            drPagare["nombre_cte"] = dsDet.Tables[0].Rows[i]["nombre_cte"];
            drPagare["apellidos_cte"] = dsDet.Tables[0].Rows[i]["apellidos_cte"];
            drPagare["codigo_gpo"] = dsDet.Tables[0].Rows[i]["codigo_gpo"];
            drPagare["nombre_gpo"] = dsDet.Tables[0].Rows[i]["nombre_gpo"];
            drPagare["calle_cte"] = dsDet.Tables[0].Rows[i]["calle_cte"];
            drPagare["colonia_cte"] = dsDet.Tables[0].Rows[i]["colonia_cte"];
            drPagare["localidad_cte"] = dsDet.Tables[0].Rows[i]["localidad_cte"];
            drPagare["municipio_cte"] = dsDet.Tables[0].Rows[i]["municipio_cte"];
            drPagare["entidad_cte"] = dsDet.Tables[0].Rows[i]["entidad_cte"];
            drPagare["pais_cte"] = dsDet.Tables[0].Rows[i]["pais_cte"];
            drPagare["cp_cte"] = dsDet.Tables[0].Rows[i]["cp_cte"];
            drPagare["fecha_inicio"] = dsDet.Tables[0].Rows[i]["fecha_inicio"];
            drPagare["fecha_fin"] = dsDet.Tables[0].Rows[i]["fecha_fin"];
            drPagare["cantidad_letra"] = Funciones.convierteNumeroaLetra(Convert.ToDecimal(dsDet.Tables[0].Rows[i]["cantidad_numero"].ToString()));
            drPagare["cantidad_numero"] = dsDet.Tables[0].Rows[i]["cantidad_numero"];
            drPagare["tasa_cte"] = dsDet.Tables[0].Rows[i]["tasa_cte"];
            drPagare["nombre_aval1"] = dsDet.Tables[0].Rows[i]["nombre_aval1"];
            drPagare["direc1_aval1"] = dsDet.Tables[0].Rows[i]["direc1_aval1"];
            drPagare["nombre_aval2"] = dsDet.Tables[0].Rows[i]["nombre_aval2"];
            drPagare["direc1_aval2"] = dsDet.Tables[0].Rows[i]["direc1_aval2"];
            drPagare["nombre_aval3"] = dsDet.Tables[0].Rows[i]["nombre_aval3"];
            drPagare["direc1_aval3"] = dsDet.Tables[0].Rows[i]["direc1_aval3"];
            drPagare["nombre_aval4"] = dsDet.Tables[0].Rows[i]["nombre_aval4"];
            drPagare["direc1_aval4"] = dsDet.Tables[0].Rows[i]["direc1_aval4"];
            drPagare["direc2_aval1"] = dsDet.Tables[0].Rows[i]["direc2_aval1"];
            drPagare["direc2_aval2"] = dsDet.Tables[0].Rows[i]["direc2_aval2"];
            drPagare["direc2_aval3"] = dsDet.Tables[0].Rows[i]["direc2_aval3"];
            drPagare["direc2_aval4"] = dsDet.Tables[0].Rows[i]["direc2_aval4"];
            drPagare["no_avales"] = dsDet.Tables[0].Rows[i]["no_avales"];
            drPagare["nombre_com_cte"] = dsDet.Tables[0].Rows[i]["nombre_com_cte"];
            drPagare["ctas_bancos"] = dsDet.Tables[0].Rows[i]["ctas_bancos"];
            drPagare["direccion_suc"] = dsDet.Tables[0].Rows[i]["direccion_suc"];
            drPagare["LOCALIDAD_SUC"] = dsDet.Tables[0].Rows[i]["LOCALIDAD_SUC"];
            drPagare["municipio_suc"] = dsDet.Tables[0].Rows[i]["municipio_suc"];
            drPagare["pais_suc"] = dsDet.Tables[0].Rows[i]["pais_suc"];

            drPagare["plazo_cte"] = dsDet.Tables[0].Rows[i]["plazo_cte"];
            drPagare["periodo_cte"] = dsDet.Tables[0].Rows[i]["periodo_cte"];
            drPagare["cant_ent_gpo"] = dsDet.Tables[0].Rows[i]["cant_ent_gpo"];
            drPagare["cant_ent_gpo_letra"] = dsDet.Tables[0].Rows[i]["cant_ent_gpo_letra"];
            drPagare["ctas_bancos"] = dsDet.Tables[0].Rows[i]["ctas_bancos"];
            drPagare["intereses_gpo"] = dsDet.Tables[0].Rows[i]["intereses_gpo"];
            drPagare["cod_pre_gpo"] = dsDet.Tables[0].Rows[i]["cod_pre_gpo"];

            drPagare["cod_sec_gpo"] = dsDet.Tables[0].Rows[i]["cod_sec_gpo"];
            drPagare["cod_tes_gpo"] = dsDet.Tables[0].Rows[i]["cod_tes_gpo"];
            drPagare["nom_pre_gpo"] = dsDet.Tables[0].Rows[i]["nom_pre_gpo"];
            drPagare["nom_sec_gpo"] = dsDet.Tables[0].Rows[i]["nom_sec_gpo"];
            drPagare["nom_tes_gpo"] = dsDet.Tables[0].Rows[i]["nom_tes_gpo"];
            drPagare["dir_pre_gpo"] = dsDet.Tables[0].Rows[i]["dir_pre_gpo"];
            drPagare["dir_sec_gpo"] = dsDet.Tables[0].Rows[i]["dir_sec_gpo"];
            drPagare["dir_tes_gpo"] = dsDet.Tables[0].Rows[i]["dir_tes_gpo"];
            drPagare["IVA"] = dsDet.Tables[0].Rows[i]["IVA"];

            drPagare["DIRECCION_EMP"] = dsDet.Tables[0].Rows[i]["DIRECCION_EMP"];

            dtRepPagare.Rows.Add(drPagare);
            dsRepPagare.Tables["dtPagare"].ImportRow(drPagare);
        }

        contFilas = dsAval.Tables[0].Rows.Count;

        for (i = 0; i < contFilas; i++)
        {
            DataRow drAval = dtRepAval.NewRow();
            drAval["CDGCLNS"] = dsAval.Tables[0].Rows[i]["CDGCLNS"];
            drAval["CDGCL"] = dsAval.Tables[0].Rows[i]["CDGCL"];
            drAval["NOMBRE"] = dsAval.Tables[0].Rows[i]["NOMBRE"];
            drAval["DIRECCION1"] = dsAval.Tables[0].Rows[i]["DIRECCION1"];
            drAval["DIRECCION2"] = dsAval.Tables[0].Rows[i]["DIRECCION2"];
            dtRepAval.Rows.Add(drAval);
            dsRepPagare.Tables["dtAval"].ImportRow(drAval);
        }

        CargaReporte("PagareComunalGrupalVST.rpt");
        grcReporte.SetDataSource(dsRepPagare);
        DisplayPdf(grcReporte);
    }

    private void LlenaRptSitCartera(DataSet ds, DataSet dsPDI, string id, DateTime fecha, int nivel, string titulo, string nomUsuario)
    {
        dsRepMora.dtMoraDataTable dtMora = new dsRepMora.dtMoraDataTable();
        DataTable dtReporteMora = (DataTable)dtMora;

        dsRepMora.dtInfoRepDataTable dtInfo = new dsRepMora.dtInfoRepDataTable();
        DataTable dtReporteInfo = (DataTable)dtInfo;

        DataSet dsMora = new dsRepMora();

        int i;
        int contFilas;
        string textoCart = string.Empty;

        contFilas = ds.Tables[0].Rows.Count;

        for (i = 0; i < contFilas; i++)
        {
            DataRow drMora = dtReporteMora.NewRow();
            drMora["REGIONAL"] = ds.Tables[0].Rows[i]["CDGRG"] + " - " + ds.Tables[0].Rows[i]["NOMRG"];
            drMora["COORDINACION"] = ds.Tables[0].Rows[i]["CDGCO"] + " - " + ds.Tables[0].Rows[i]["NOMCO"];
            drMora["COORDINADOR"] = ds.Tables[0].Rows[i]["CDGCOPE"] + " - " + ds.Tables[0].Rows[i]["NOMCOPE"];
            drMora["OF_CREDITO"] = ds.Tables[0].Rows[i]["CDGOCPE"] + " - " + ds.Tables[0].Rows[i]["NOMPE"];
            drMora["GRUPO"] = ds.Tables[0].Rows[i]["CDGCLNS"] + " " + ds.Tables[0].Rows[i]["NOMCLNS"];
            drMora["CICLO"] = ds.Tables[0].Rows[i]["CICLO"];
            drMora["INICIO"] = Convert.ToDateTime(ds.Tables[0].Rows[i]["INICIO"]).ToString("dd/MM/yyyy");
            drMora["CANT_ENTRE"] = ds.Tables[0].Rows[i]["CARTERAVIG"];
            drMora["SALDO_TOTAL"] = ds.Tables[0].Rows[i]["MORA3"];
            drMora["PAGOS_COMP"] = ds.Tables[0].Rows[i]["MORA1"];
            drMora["PAGOS_EFECT"] = ds.Tables[0].Rows[i]["MORA2"];
            drMora["MORA_TOTAL"] = ds.Tables[0].Rows[i]["SALDO"];
            drMora["DIAS_MORA"] = ds.Tables[0].Rows[i]["DIAS_MORA"];
            drMora["NUM_INTEG"] = ds.Tables[0].Rows[i]["MORA4"];
            drMora["SALDO_GL"] = ds.Tables[0].Rows[i]["MORA5"];
            drMora["MORATORIOS"] = ds.Tables[0].Rows[i]["MORA6"];
            drMora["PARCIALIDAD"] = ds.Tables[0].Rows[i]["PARCIALIDAD"];
            drMora["FEC_FIN"] = ds.Tables[0].Rows[i]["FECHA_FIN"];
            drMora["FEC_ULT_PAGO"] = ds.Tables[0].Rows[i]["FECHA_ULT_PAGO"];
            drMora["MONTO_ULT_PAGO"] = ds.Tables[0].Rows[i]["MONTO_ULT_PAGO"];
            drMora["PERIO_TRANS"] = ds.Tables[0].Rows[i]["PERIO_TRANS"];
            drMora["TASA"] = ds.Tables[0].Rows[i]["TASA"];
            drMora["TOTAL_PAGAR"] = ds.Tables[0].Rows[i]["TOTAL_PAGAR"];
            drMora["PLAZO"] = ds.Tables[0].Rows[i]["PLAZO"];
            drMora["TIPOPROD"] = ds.Tables[0].Rows[i]["TIPOPROD"];
            dtReporteMora.Rows.Add(drMora);
            dsMora.Tables["dtMora"].ImportRow(drMora);
        }

        DataRow drInfo = dtReporteInfo.NewRow();
        drInfo["FEC_IMP"] = ds.Tables[0].Rows[0]["FECHAIMP"];
        drInfo["HORA_IMP"] = ds.Tables[0].Rows[0]["HORAIMP"];
        drInfo["FECHA"] = fecha.ToString("dd/MM/yyyy");
        drInfo["USUARIO"] = nomUsuario;
        drInfo["REGISTROS"] = contFilas;
        drInfo["REG_PDI"] = dsPDI.Tables[0].Rows[0]["REG_PDI"];
        drInfo["MONTO_PDI"] = dsPDI.Tables[0].Rows[0]["MONTO_PDI"];
        drInfo["TIPO_CART"] = titulo;
        dtReporteInfo.Rows.Add(drInfo);
        dsMora.Tables["dtInfoRep"].ImportRow(drInfo);

        switch (nivel)
        {
            case 0:
                CargaReporte("RepSitCarteraReg.rpt");
                break;
            case 1:
                CargaReporte("RepSitCarteraSuc.rpt");
                break;
            case 2:
                CargaReporte("RepSitCarteraCoord.rpt");
                break;
            case 3:
                CargaReporte("RepSitCarteraAsesor.rpt");
                break;
        }
        grcReporte.SetDataSource(dsMora);
        DisplayPdf(grcReporte);
    }

    private void LlenaRptTarjetaPagos(DataSet ds)
    {
        dsControl.dtGrupoDataTable dtEnc = new dsControl.dtGrupoDataTable();
        DataTable dtReporteEnc = (DataTable)dtEnc;

        dsControl.dtAcreditadoDataTable dtDet = new dsControl.dtAcreditadoDataTable();
        DataTable dtReporteDet = (DataTable)dtDet;

        DataSet dsResp = new dsControl();

        int i;
        int contFilas;

        contFilas = ds.Tables[0].Rows.Count;
        
        for (i = 0; i < contFilas; i++)
        {
            DataRow drEnc = dtReporteEnc.NewRow();
            drEnc["NUMERO"] = i + 1;
            drEnc["COD_GRUPO"] = ds.Tables[0].Rows[i]["CDGNS"];
            drEnc["GRUPO"] = ds.Tables[0].Rows[i]["GRUPO"];
            drEnc["INICIO"] = ds.Tables[0].Rows[i]["INICIO"];
            drEnc["CICLO"] = ds.Tables[0].Rows[i]["CICLO"];
            drEnc["PLAZO"] = ds.Tables[0].Rows[i]["PLAZO"];
            drEnc["DIA_PAGO"] = ds.Tables[0].Rows[i]["DIA_PAGO"];
            drEnc["HORARIO"] = ds.Tables[0].Rows[i]["HORARIO"];
            drEnc["PRESIDENTE"] = ds.Tables[0].Rows[i]["NOM_PRES"];
            drEnc["TESORERO"] = ds.Tables[0].Rows[i]["NOM_TES"];
            drEnc["SECRETARIO"] = ds.Tables[0].Rows[i]["NOM_SEC"];
            drEnc["SUPERVISOR"] = ds.Tables[0].Rows[i]["NOM_SUP"];
            dtReporteEnc.Rows.Add(drEnc);
            dsResp.Tables["dtGrupo"].ImportRow(drEnc);

            DataRow drDet = dtReporteDet.NewRow();
            drDet["NUMERO"] = i + 1;
            drDet["NOMBRE"] = ds.Tables[0].Rows[i]["ACRED"];
            drDet["CANTENTRE"] = ds.Tables[0].Rows[i]["CANTENTRE"];
            drDet["PARCIALIDAD"] = ds.Tables[0].Rows[i]["PARCIALIDAD"];
            drDet["INTERES"] = ds.Tables[0].Rows[i]["INTERES"];
            drDet["TOTAL_PAGAR"] = ds.Tables[0].Rows[i]["TOTALPAGAR"];
            drDet["GARANTIA"] = ds.Tables[0].Rows[i]["GARANTIA"];
            dtReporteDet.Rows.Add(drDet);
            dsResp.Tables["dtAcreditado"].ImportRow(drDet);
        }

        CargaReporte("TarjetaPagos.rpt");
        grcReporte.SetDataSource(dsResp);
        DisplayPdf(grcReporte);
    }

    private void LlenaSolicitudCred(DataSet ds, DataSet dsGrupo, DataSet dsAcred, DataSet dsExcAcred, DataSet dsMaxAcred, string nomUsuario)
    {
        dsSolicitudCred.dtAcreditadoDataTable dt = new dsSolicitudCred.dtAcreditadoDataTable();
        DataTable dtAcred = (DataTable)dt;

        dsSolicitudCred.dtGrupoDataTable dtG = new dsSolicitudCred.dtGrupoDataTable();
        DataTable dtGrupo = (DataTable)dtG;

        dsSolicitudCred.dtInfoRepDataTable dtInfo = new dsSolicitudCred.dtInfoRepDataTable();
        DataTable dtReporteInfo = (DataTable)dtInfo;

        dsSolicitudCred.dtExcAcredDataTable dtExc = new dsSolicitudCred.dtExcAcredDataTable();
        DataTable dtExcAcred = (DataTable)dtExc;

        dsSolicitudCred.dtMaxAcredDataTable dtMax = new dsSolicitudCred.dtMaxAcredDataTable();
        DataTable dtMaxAcred = (DataTable)dtMax;

        DataSet dsSolic = new dsSolicitudCred();

        int i;
        int j;
        int n;
        int filasAcred;
        int numAdic;
        string tipoProd;

        try
        {
            /*Se obtiene el dato del tipo de producto*/
            tipoProd = dsGrupo.Tables[0].Rows[0]["CDGTPC"].ToString();
            
            DataRow drGrupo = dtGrupo.NewRow();
            drGrupo["COD_GRUPO"] = dsGrupo.Tables[0].Rows[0]["CDGNS"];
            drGrupo["GRUPO"] = dsGrupo.Tables[0].Rows[0]["GRUPO"];
            drGrupo["CICLO"] = dsGrupo.Tables[0].Rows[0]["CICLO"];
            drGrupo["COORDINACION"] = dsGrupo.Tables[0].Rows[0]["COORDINACION"];
            drGrupo["ASESOR"] = dsGrupo.Tables[0].Rows[0]["NOMBREASESOR"];
            drGrupo["COORDINADOR"] = dsGrupo.Tables[0].Rows[0]["NOMBRECOORDINADOR"];
            drGrupo["GERENTE"] = dsGrupo.Tables[0].Rows[0]["NOMBREGERENTE"];
            drGrupo["NOMBREASIST"] = dsGrupo.Tables[0].Rows[0]["NOMBREASIST"];
            drGrupo["GTE_REGIONAL"] = dsGrupo.Tables[0].Rows[0]["NOMBREREGIONAL"];
            drGrupo["PAGO_ANT"] = dsGrupo.Tables[0].Rows[0]["PAGO_ANT"];
            drGrupo["TASA"] = dsGrupo.Tables[0].Rows[0]["TASA"];
            drGrupo["NUMEXCACRED"] = dsExcAcred.Tables[0].Rows.Count;
            drGrupo["CICLOA"] = dsGrupo.Tables[0].Rows[0]["CICLOA"];
            drGrupo["CICLOB"] = dsGrupo.Tables[0].Rows[0]["CICLOB"];
            drGrupo["CICLOC"] = dsGrupo.Tables[0].Rows[0]["CICLOC"];
            drGrupo["CICLOD"] = dsGrupo.Tables[0].Rows[0]["CICLOD"];
            drGrupo["TASAA"] = dsGrupo.Tables[0].Rows[0]["TASAA"];
            drGrupo["TASAB"] = dsGrupo.Tables[0].Rows[0]["TASAB"];
            drGrupo["TASAC"] = dsGrupo.Tables[0].Rows[0]["TASAC"];
            drGrupo["TASAD"] = dsGrupo.Tables[0].Rows[0]["TASAD"];
            drGrupo["DIASA"] = dsGrupo.Tables[0].Rows[0]["DIASA"];
            drGrupo["DIASB"] = dsGrupo.Tables[0].Rows[0]["DIASB"];
            drGrupo["DIASC"] = dsGrupo.Tables[0].Rows[0]["DIASC"];
            drGrupo["DIASD"] = dsGrupo.Tables[0].Rows[0]["DIASD"];
            drGrupo["INSDOC"] = dsGrupo.Tables[0].Rows[0]["INSDOC"];
            drGrupo["INTEGRACION"] = dsGrupo.Tables[0].Rows[0]["INTEGRACION"];
            drGrupo["ANALISISIND"] = dsGrupo.Tables[0].Rows[0]["ANALISISIND"];

            dtGrupo.Rows.Add(drGrupo);
            dsSolic.Tables["dtGrupo"].ImportRow(drGrupo);

            filasAcred = dsAcred.Tables[0].Rows.Count;

            for (i = 0; i < filasAcred; i++)
            {
                DataRow drAcred = dtAcred.NewRow();
                drAcred["NUMERO"] = i + 1;
                drAcred["COD_ACRED"] = dsAcred.Tables[0].Rows[i]["NO_CTE"];
                drAcred["NOMBRE"] = dsAcred.Tables[0].Rows[i]["CLIENTE"];
                drAcred["CICLO_ACRED"] = dsAcred.Tables[0].Rows[i]["CICLO_SOL"];
                drAcred["PRES_ANTE"] = dsAcred.Tables[0].Rows[i]["CANTENTRE"];
                dtAcred.Rows.Add(drAcred);
                dsSolic.Tables["dtAcreditado"].ImportRow(drAcred);
            }

            if (filasAcred < 10)
                numAdic = 10 - filasAcred;
            else
                numAdic = 3;

            /*Filas de acreditado en blanco*/
            for (n = 0; n < numAdic; n++)
            {
                i += 1;
                DataRow drAd = dtAcred.NewRow();
                drAd["NUMERO"] = i;
                dtAcred.Rows.Add(drAd);
                dsSolic.Tables["dtAcreditado"].ImportRow(drAd);
            }

            j = i;

            if (i % 10 != 0)
            {
                for (n = 0; n < (9 - (j % 10)); n++)
                {
                    i += 1;
                    DataRow dr = dtAcred.NewRow();
                    dr["NUMERO"] = i;
                    dtAcred.Rows.Add(dr);
                    dsSolic.Tables["dtAcreditado"].ImportRow(dr);
                }
            }

            DataRow drInfo = dtReporteInfo.NewRow();
            drInfo["FEC_IMP"] = ds.Tables[0].Rows[0]["FECHAIMP"];
            drInfo["HORA_IMP"] = ds.Tables[0].Rows[0]["HORAIMP"];
            drInfo["USUARIO"] = nomUsuario;
            dtReporteInfo.Rows.Add(drInfo);
            dsSolic.Tables["dtInfoRep"].ImportRow(drInfo);

            filasAcred = dsExcAcred.Tables[0].Rows.Count;

            for (i = 0; i < filasAcred; i++)
            {
                DataRow drExc = dtExcAcred.NewRow();
                drExc["COD_ACRED"] = dsExcAcred.Tables[0].Rows[i]["ACRED"];
                drExc["NOM_ACRED"] = dsExcAcred.Tables[0].Rows[i]["NOMBRE"];
                drExc["DESC_EXCEP"] = dsExcAcred.Tables[0].Rows[i]["DESCRIPCION"];
                drExc["OBSERVACION"] = dsExcAcred.Tables[0].Rows[i]["OBSERVACION"];
                dtExcAcred.Rows.Add(drExc);
                dsSolic.Tables["dtExcAcred"].ImportRow(drExc);
            }

            filasAcred = dsMaxAcred.Tables[0].Rows.Count;

            for (i = 0; i < filasAcred; i++)
            {
                DataRow drMax = dtMaxAcred.NewRow();
                drMax["COD_ACRED"] = dsMaxAcred.Tables[0].Rows[i]["ACRED"];
                drMax["NOM_ACRED"] = dsMaxAcred.Tables[0].Rows[i]["NOMBRE_ACRED"];
                drMax["MONTO_MAX"] = dsMaxAcred.Tables[0].Rows[i]["MONTOMAX"];
                dtMaxAcred.Rows.Add(drMax);
                dsSolic.Tables["dtMaxAcred"].ImportRow(drMax);
            }

            if (tipoProd == "01" || tipoProd == "03")
                CargaReporte("SolicitudCredComVST.rpt");
            grcReporte.SetDataSource(dsSolic);
            DisplayPdf(grcReporte);
        }
        catch (Exception ex)
        {
            string mensaje = ex.Message;
            LlenaRptError("", "No es posible generar el documento.\n\nVerifique la información solicitada.");
        }
    }
}