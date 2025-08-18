using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Collections;

/// <summary>
/// Descripción breve de ReportePlanPagos
/// </summary>
public class ReportePlanPagos
{
    DBService db;
	public ReportePlanPagos()
	{
        
	}

    public int getReportePlanPagos(string empresa,string credito, string ciclo,ref DataTable dt)
    {
        db = new DBService();

        DateTime fechaTeorica = DateTime.Now;
        string sqlCredito = "select TO_CHAR(INICIO,'DD/MM/YYYY') AS INICIO,PERIODICIDAD,CANTENTRE,PLAZO,TASA, '' AS NUMPAGO,'' AS FECHATEORICA,'' AS FECHAREAL, '' AS PAGOTEORICO, '' AS PAGOREAL, '' AS DIASMORA,'' as ADEUDO,' ' AS PARCIALIDAD,"+
                "SITUACION,CO.NOMBRE AS COORDINACION,NS.NOMBRE AS GRUPO,'' AS CICLO, '' AS DIFERENCIA,CO.CODIGO AS CODIGO,RG.CODIGO AS CODREGION, RG.NOMBRE AS REGION,NS.CODIGO AS CODGRUPO, ' ' AS FECHAFIN" +
                " FROM PRN"+
                " INNER JOIN CO ON (CO.CODIGO=PRN.CDGCO)"+
                " INNER JOIN NS ON (NS.CODIGO=PRN.CDGNS)"+
                " INNER JOIN RG ON (CO.CDGRG=RG.CODIGO AND RG.CDGEM = 'EMPFIN')"+
                " WHERE PRN.CDGNS=" + credito + " AND PRN.CICLO='" + ciclo + "' AND PRN.SITUACION IN ('E','L') AND PRN.CDGEM='"+empresa+"'";
     

        string sqlPagos="SELECT  DISTINCT a.FREALDEP,(select sum(b.cantidad) from mp b where a.FREALDEP=b.FREALDEP and a.cdgns=b.cdgns and a.ciclo=b.ciclo and a.CDGEM=b.CDGEM) AS MONTO, 0 AS PAGADO from mp a"+
                  " where a.cdgns="+credito+" and a.ciclo = '"+ciclo+"' and a.PERIODO>0 AND a.CDGEM='"+empresa+"'"+
                  " ORDER BY a.FREALDEP";
        int year = fechaTeorica.Year;
        string sqlDNH = "select fecha from dnh";//where fecha between '01/01/" + year + "' and '31/12/" + year + "'";


        DataTable dtCredito = new DataTable();
        DataTable dtPagos = new DataTable();
        DataTable dnh = new DataTable();

        int result = -1;
        result = db.getDataTableFromQuery(sqlCredito, ref dtCredito);
        
        if (result == -1 || dtCredito.Rows.Count==0)
            return -1;

        db.getDataTableFromQuery(sqlPagos, ref dtPagos);
        db.getDataTableFromQuery(sqlDNH, ref dnh);

        int plazo = 0;

        double totalPagar = 0;
        double totalAcumulado = 0;

        double parcialidad = 0;
        double tasa = 0;
        double entregado = 0;
        string periodo = "";
        int periodoDias = 0;
        string sucursal = "";
        string situacion = "";
        string grupo = "";
        string fechaInicio = "";
        double porcentajeParcialidad = 0;
        bool evaluadoTop = false;

        string codigo="";
        string codregion = "";
        string region = "";
        string codgrupo = "";
        try
        {

            fechaTeorica = DateTime.Parse(dtCredito.Rows[0][0].ToString());
            fechaInicio = fechaTeorica.ToString("dd/MM/yyyy");

            plazo = int.Parse(dtCredito.Rows[0][3].ToString());
            tasa = Double.Parse(dtCredito.Rows[0][4].ToString());
            entregado = Double.Parse(dtCredito.Rows[0][2].ToString());
            periodo = dtCredito.Rows[0][1].ToString();
            sucursal = dtCredito.Rows[0][14].ToString();
            situacion = dtCredito.Rows[0][13].ToString();
            codigo = dtCredito.Rows[0][18].ToString();
            codregion = dtCredito.Rows[0][19].ToString();
            region = dtCredito.Rows[0][20].ToString();
            codgrupo = dtCredito.Rows[0][21].ToString();
            if (situacion == "L")
            {
                situacion = "LIQUIDADO";
            }
            if (situacion == "E")
            {
                situacion = "ENTREGADO";
            }
            grupo = dtCredito.Rows[0][15].ToString();
            switch (periodo)
            {
                case "S":
                    periodoDias = 7;
                    break;
                case "Q":
                    periodoDias = 15;
                    break;
                case "C":
                    periodoDias = 14;
                    break;
                case "M":
                    periodoDias = 30;
                    break;
            }
        }
        catch (Exception ex) { return -1; }

        int diasMora = 0;
        parcialidad = CalculaParcialidad(periodo, tasa, plazo, entregado);
        porcentajeParcialidad = parcialidad * .10;

        totalPagar = parcialidad * plazo;

        if (porcentajeParcialidad > 500)
        {
            porcentajeParcialidad = parcialidad - 500;
        }
        else { porcentajeParcialidad = parcialidad - porcentajeParcialidad; }


        int count = 1;
        DataTable dtPlanPagos = dtCredito.Clone();
        int jFechaIni = 0;
        double adeudo = 0;
        double adeudoOriginal = 0;

        //Calcula fecha final de pago
        DateTime _fechaIni = DateTime.Parse(dtCredito.Rows[0][0].ToString());
        DateTime _fechaValidada=new DateTime();
        for (int jFinal = 0; jFinal < plazo; jFinal++)
        {
            
            _fechaIni = _fechaIni.AddDays(periodoDias);
            _fechaValidada = ValidaDiaFestivo(_fechaIni,dnh);
            dtCredito.Rows[0][22] = _fechaValidada.ToString("dd/MM/yyyy");
        }
        DateTime fechaComparar = new DateTime();
        for (int i = 0; i < plazo; i++)
        {
            int anteriorValido = 0;
            DateTime fechaSiguientePago = new DateTime();
            Boolean ignorarSiguientePago = false;


            if (fechaComparar < _fechaValidada && totalAcumulado >= totalPagar)
                break;
                
            DataRow dr = dtPlanPagos.NewRow();
            DateTime fechaReal = DateTime.Now;
            string strFechaReal = "";
            double siguientePago = 0;
            double cantidadReal = 0;
            double cantidadConAdeudo = 0;
            dr[0] = fechaInicio.Substring(0, 10);
            dr[1] = periodo;
            dr[2] = entregado.ToString();
            dr[3] = plazo.ToString();
            dr[4] = tasa.ToString();
            dr[5] = count;                        

            fechaTeorica = fechaTeorica.AddDays(periodoDias);

            DateTime fechaTeoricaValidada = ValidaDiaFestivo(fechaTeorica, dnh);

            dr[6] = fechaTeoricaValidada.ToString("dd/MM/yyyy");
            
            try
            {
                for (int jFecha = jFechaIni; jFecha <= dtPagos.Rows.Count; jFecha++)
                {
                    try
                    {
                        fechaSiguientePago = DateTime.Parse(dtPagos.Rows[jFecha + 1][0].ToString());
                        fechaComparar = fechaSiguientePago;
                        
                    }
                    catch (Exception eFecha) { fechaSiguientePago = DateTime.Now; }

                    if (ignorarSiguientePago)
                    {
                        ignorarSiguientePago = false;
                        continue;
                    }

                    
                    
                    if (cantidadConAdeudo >= parcialidad)
                    {
                        jFechaIni = jFecha;
                        adeudoOriginal = adeudo;                        
                        adeudo = (parcialidad - cantidadConAdeudo)+adeudo;                                                
                        
                      
                        dr[9] = cantidadReal.ToString();

                        totalAcumulado += cantidadReal;

                        cantidadReal = 0;
                        cantidadConAdeudo = 0;

                        //if (fechaTeoricaValidada > _fechaValidada)
                            
                        break;
                    }
                    else
                    {                        
                        fechaReal = DateTime.Parse(dtPagos.Rows[jFecha][0].ToString());
                        if (fechaTeoricaValidada >= fechaReal)
                        {
                            strFechaReal = fechaTeoricaValidada.ToString("dd/MM/yyyy") ;
                        }
                        else
                        {
                            strFechaReal = fechaReal.ToString("dd/MM/yyyy");
                        }
    
                        if (-adeudo >= parcialidad)
                        {
                            cantidadReal =0;
                            cantidadConAdeudo += -adeudo;
                            adeudo = 0;
                            //adeudo = parcialidad - -(adeudo);
                            jFecha--;
                            fechaReal = DateTime.Parse(dtPagos.Rows[jFecha][0].ToString());
                            if (fechaTeorica >= fechaReal)
                            {
                                strFechaReal = fechaTeorica.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                strFechaReal = fechaReal.ToString("dd/MM/yyyy");
                            }


                        }
                        else
                        {
                            cantidadReal += Double.Parse(dtPagos.Rows[jFecha][1].ToString());
                            cantidadConAdeudo += Double.Parse(dtPagos.Rows[jFecha][1].ToString());

                            if (cantidadReal < 0)
                            {
                                cantidadReal = 0;
                                cantidadConAdeudo = 0;
                                continue;
                            }
                            try
                            {
                                siguientePago = Double.Parse(dtPagos.Rows[jFecha + 1][1].ToString());

                                //NUEVA MAGIA
                                if (siguientePago > 0 && fechaSiguientePago <= fechaTeoricaValidada)
                                {
                                    cantidadReal += siguientePago;
                                    cantidadConAdeudo += siguientePago;
                                    anteriorValido = jFecha + 1;
                                    ignorarSiguientePago = true;
                                }

                                if (siguientePago < 0)
                                {
                                    cantidadReal = cantidadReal- -siguientePago;
                                    cantidadConAdeudo = cantidadConAdeudo - -siguientePago;
                                }
                            }
                            catch (Exception ePago) { siguientePago = 0; }
                        }
                        
                        if (adeudo < 0 && cantidadConAdeudo < parcialidad)
                        {
                            //cantidadReal += -(adeudo);
                            cantidadConAdeudo += -(adeudo);

                            //adeudo = cantidadConAdeudo - adeudo;
                            adeudo = 0;
                        }
                    }
                }
            }
            catch (Exception exFecha) { 
                strFechaReal = dtPlanPagos.Rows[dtPlanPagos.Rows.Count - 1][7].ToString(); cantidadReal = 0; dr[9] = 0; diasMora = 0; adeudo = 0;
                fechaReal = DateTime.Parse(strFechaReal);
            }
            /*
            if (fechaTeorica > DateTime.Now)
                break;
            */

            dr[7] = strFechaReal;
            dr[8] = parcialidad.ToString();
            //dr[9] = cantidadReal.ToString();
            TimeSpan fechaDif = new TimeSpan();
            if (fechaReal > fechaTeorica)
            {                
                fechaDif = fechaReal - fechaTeoricaValidada;
            }
            //double adeudo = parcialidad - cantidadReal;

            diasMora = fechaDif.Days;
            dr[10] = diasMora;
            dr[11] = adeudo;
            dr[12] = parcialidad;
            dr[13] = situacion;
            dr[14] = sucursal;
            dr[15] = grupo;
            dr[16] = ciclo;
            dr[18] = codigo;
            dr[19] = codregion;
            dr[20] = region;
            dr[21] = codgrupo;
            dr[22] = dtCredito.Rows[0][22];

            if (fechaTeoricaValidada < _fechaValidada)
            {
                if (totalAcumulado >= totalPagar)
                {                    
                    adeudo = totalPagar - totalAcumulado;
                    dr[11] = adeudo;
                    dtPlanPagos.Rows.Add(dr);
                    break;
                }
            }
            dtPlanPagos.Rows.Add(dr);
            count++;
        }
        dt = dtPlanPagos;
        dt = getDiferenciaPago_2(dtPagos, dtPlanPagos);
        return 1;
    }

    public DataTable getDiferenciaPago(DataTable dtPagos, DataTable dtPlanPagos)
    {
        double parcialidadPorcentaje = 0;
        DateTime startDate = DateTime.Parse(dtPlanPagos.Rows[0][0].ToString());
        DateTime endDate = DateTime.Now;
        double montoAcumulado = 0;
        double montoTotalAcumulado = 0;
        double parcialidad = 0;
        double parcialidadComparar = 0;
        parcialidad = Double.Parse(dtPlanPagos.Rows[0][12].ToString());
        parcialidadPorcentaje = parcialidad * .10;
        parcialidadComparar = parcialidadPorcentaje;
        
        if (parcialidadPorcentaje > 500)
        {
            parcialidadPorcentaje = parcialidad - 500;
            parcialidadComparar = 500;
        }
        else
        {
            parcialidadPorcentaje = parcialidad - parcialidadPorcentaje;
        }

        for (int i = 0; i < dtPlanPagos.Rows.Count; i++)
        {
            endDate = DateTime.Parse(dtPlanPagos.Rows[i][6].ToString());
            //endDate = DateTime.Parse(dtPlanPagos.Rows[i][7].ToString());
            //string fechaFin = dtPlanPagos.Rows[dtPlanPagos.Rows.Count-1][6].ToString();
            //dtPlanPagos.Rows[i][22] = fechaFin.Substring(0,10);
            string f = dtPlanPagos.Rows[0][0].ToString().Substring(0, 10);
            dtPlanPagos.Rows[i][0] = f;
            int iPago = -1;
            foreach (DataRow row in dtPagos.Rows)
            {
                iPago++;
                int pagado = int.Parse(row[2].ToString());
                if (pagado == 1)
                    continue;

                DateTime fechaPago = DateTime.Parse(row[0].ToString());
                double montoPago = double.Parse(row[1].ToString());

                if (fechaPago >= startDate && fechaPago <= endDate)
                {
                    montoAcumulado += montoPago;
                    row[2] = 1;
                }
            }
            bool setDiferencia = false;
            if (montoAcumulado < parcialidadPorcentaje)
            {
                //dtPlanPagos.Rows[i][10] = 0;
                double diferenciaReal = parcialidad - montoAcumulado;

                if ((montoTotalAcumulado+diferenciaReal) < parcialidad)
                {
                    montoTotalAcumulado += diferenciaReal;
                }
                else { montoTotalAcumulado = diferenciaReal; }


                dtPlanPagos.Rows[i][17] = montoTotalAcumulado;
                setDiferencia = true;
                //montoTotalAcumulado = 0;
            }
            double _diferenciaReal = 0;
            if (!setDiferencia)
            {
                _diferenciaReal = parcialidad - montoAcumulado;

                if ((montoTotalAcumulado+_diferenciaReal) < parcialidad)
                {
                    montoTotalAcumulado += _diferenciaReal;
                }
                else { montoTotalAcumulado = _diferenciaReal; }

                dtPlanPagos.Rows[i][17] = montoTotalAcumulado;
                if (montoTotalAcumulado < parcialidadComparar && montoTotalAcumulado!=0)
                    dtPlanPagos.Rows[i][10] = 0;

                //montoTotalAcumulado = 0;
            }

            startDate = endDate.AddDays(1);
            
            montoAcumulado = 0;
            //montoTotalAcumulado = 0;

        }

        return dtPlanPagos;
    }

    public DataTable getDiferenciaPago_2(DataTable dtPagos, DataTable dtPlanPagos)
    {
        double acumuladoTeorico = 0;
        double acumuladoReal = 0;
        int numRow = -1;
        int rowDelete = 0;
        bool hayPagos = false;
        DateTime fechaUltimoPago = new DateTime();
        ArrayList lista = new ArrayList();
        foreach (DataRow row in dtPlanPagos.Rows)
        {
            numRow++;
            hayPagos = false;
            DateTime fechaTeorica = DateTime.Parse(row[6].ToString()) ;
            double parcialidad = double.Parse(row[12].ToString());
            acumuladoTeorico += parcialidad;
            double plazo=int.Parse(row[3].ToString());

            double totalPagar = parcialidad * plazo;

            double parcialidadPorcentaje = 0;
            parcialidadPorcentaje = parcialidad * .10;

            for(int i=0;i<dtPagos.Rows.Count;i++)
            {
                
                int pagado = int.Parse(dtPagos.Rows[i][2].ToString());
                
                if (pagado == 1)
                    continue;

                hayPagos = true;
                DateTime fechaReal = DateTime.Parse(dtPagos.Rows[i][0].ToString());
                if (fechaReal <= fechaTeorica)
                {
                    acumuladoReal += double.Parse(dtPagos.Rows[i][1].ToString());
                    dtPagos.Rows[i][2] = 1;
                    fechaUltimoPago = DateTime.Parse(dtPagos.Rows[i][0].ToString());
                }

                if (acumuladoReal >= totalPagar)
                {
                    rowDelete = numRow;
                }

                if (fechaReal > fechaTeorica)
                    break;


            }

            if (hayPagos)
            {
                double diferencia = acumuladoTeorico - acumuladoReal;
                row[17] = diferencia;
                row[0] = (String)row[0].ToString().Substring(0, 10);
                DateTime vFechaFin;
                DateTime vFecha = DateTime.Parse(row[6].ToString());
                if (diferencia <= parcialidadPorcentaje)
                    row[10] = 0;
            }
            else
            {
                lista.Add(numRow);
            }
        }

        dtPlanPagos.Rows[rowDelete][7] = fechaUltimoPago.ToString("dd/MM/yyyy");
        for (int iDix = lista.Count; iDix >= 0; iDix--)
        {
            if ((iDix - 1) == -1)
                break;
            int idx = int.Parse(lista[iDix-1].ToString());
            dtPlanPagos.Rows.RemoveAt(idx);
        }
        return dtPlanPagos;
    }

    private double CalculaParcialidad(string periodicidad, double tasa, int plazo, double entregado)
    {
        double parcial = 0;
        switch (periodicidad)
        {
            case "S":
                parcial = (tasa * plazo * entregado) / (4 * 100);
                break;
            case "Q":
                parcial = ((tasa * plazo * entregado) * 15) / (30 * 100);
                break;
            case "C":
                parcial = (tasa * plazo * entregado) / (2 * 100);
                break;
            case "M":
                parcial = (tasa * plazo * entregado) / 100;
                break;
        }
        
        parcial = (entregado + parcial) / plazo;
        return parcial;
    }

    private DateTime ValidaDiaFestivo(DateTime fecha, DataTable inhabiles)
    {
        DateTime fechaIn = new DateTime();
        for (int i = 0; i < inhabiles.Rows.Count; i++)
        {
            fechaIn = DateTime.Parse(inhabiles.Rows[i][0].ToString());
            if (fecha == fechaIn)
            {
                fecha = fecha.AddDays(1);
                if (fecha.DayOfWeek == DayOfWeek.Saturday)
                {
                    fecha = fecha.AddDays(2);
                }
                else if (fecha.DayOfWeek == DayOfWeek.Sunday)
                {
                    fecha = fecha.AddDays(1);
                }
            }
        }
        return fecha;
    }
}
