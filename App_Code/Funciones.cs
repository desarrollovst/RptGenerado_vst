using System;
using System.Collections.Generic;
using System.Web;
using System.IO;
using Microsoft.VisualBasic;
using System.Drawing;

/// <summary>
/// Descripción breve de Funciones
/// </summary>
public static class Funciones
{
    public static byte[] ConversionImagen(string nombrearchivo)
    {
        //Declaramos fs para poder abrir la imagen.
        FileStream fs = new FileStream(nombrearchivo, FileMode.Open);
        // Declaramos un lector binario para pasar la imagen a bytes
        BinaryReader br = new BinaryReader(fs);
        byte[] imagen = new byte[(int)fs.Length];
        br.Read(imagen, 0, (int)fs.Length);
        br.Close();
        fs.Close();
        return imagen;
    }

    #region Convertir

    /**
 * Esta clase provee la funcionalidad de convertir un numero representado en
 * digitos a una representacion en letras. Mejorado para leer centavos
 * 
 */
    private static string[] UNIDADES = { "", "UN ", "DOS ", "TRES ",
			"CUATRO ", "CINCO ", "SEIS ", "SIETE ", "OCHO ", "NUEVE ", "DIEZ ",
			"ONCE ", "DOCE ", "TRECE ", "CATORCE ", "QUINCE ", "DIECISEIS ",
			"DIECISIETE ", "DIECIOCHO ", "DIECINUEVE ", "VEINTE " };

    private static string[] DECENAS = { "VEINTI", "TREINTA ", "CUARENTA ",
			"CINCUENTA ", "SESENTA ", "SETENTA ", "OCHENTA ", "NOVENTA ",
			"CIEN " };

    private static string[] CENTENAS = { "CIENTO ", "DOSCIENTOS ",
			"TRESCIENTOS ", "CUATROCIENTOS ", "QUINIENTOS ", "SEISCIENTOS ",
			"SETECIENTOS ", "OCHOCIENTOS ", "NOVECIENTOS " };

    /**
     * Convierte a letras un numero de la forma $123,456.32 (StoreMath)
     * <p>
     * Creation date 5/06/2006 - 10:20:52 AM
     * 
     * @param number
     *            Numero en representacion texto
     * @return Numero en letras
     * @since 1.0
     *
     *
     * Convierte un numero en representacion numerica a uno en representacion de
     * texto. El numero es valido si esta entre 0 y 999'999.999
     * <p>
     * Creation date 3/05/2006 - 05:37:47 PM
     * 
     * @param number
     *            Numero a convertir
     * @return Numero convertido a texto
     * @throws NumberFormatException
     *             Si el numero esta fuera del rango
     * @since 1.0
     */
    public static string convierteNumeroaLetra(decimal number)
    {
        string converted = string.Empty;

        // Validamos que sea un numero legal
        //double doubleNumber = StoreMath.round(number);
        if (number > 999999999)
            //throw new NumberFormatException(
            //		"El numero es mayor de 999'999.999, "
            //				+ "no es posible convertirlo");
            return "";

        string[] splitNumber = number.ToString().Replace('.', '#').Split('#');

        // Descompone el trio de millones - ¡SGT!
        int millon = Convert.ToInt32(getDigitAt(splitNumber[0], 8)
                + getDigitAt(splitNumber[0], 7)
                + getDigitAt(splitNumber[0], 6));
        if (millon == 1)
            converted = "UN MILLON ";
        if (millon > 1)
            converted = convierteNumero(millon.ToString()) + "MILLONES ";

        // Descompone el trio de miles - ¡SGT!
        int miles = Convert.ToInt32(getDigitAt(splitNumber[0], 5)
                + getDigitAt(splitNumber[0], 4)
                + getDigitAt(splitNumber[0], 3));
        if (miles == 1)
            converted += "MIL ";
        if (miles > 1)
            converted += convierteNumero(miles.ToString()) + "MIL ";

        // Descompone el ultimo trio de unidades - ¡SGT!
        int cientos = Convert.ToInt32(getDigitAt(splitNumber[0], 2)
                + getDigitAt(splitNumber[0], 1)
                + getDigitAt(splitNumber[0], 0));
        if (cientos == 1)
            converted += "UN ";

        if (millon + miles + cientos == 0)
            converted += "CERO ";
        if (cientos > 1)
            converted += convierteNumero(cientos.ToString());

        //Verifica la cantidad para escribir el texto en plural o singular
        if (number >= 1 && number < 2)
            converted += "PESO ";
        else
            converted += "PESOS ";

        // Descompone los centavos - Camilo
        string centavos = string.Empty;

        if (splitNumber.Length > 1)
            centavos = getDigitAt(splitNumber[1], 1) + getDigitAt(splitNumber[1], 0);
        else
            centavos = "00";

        converted += centavos + "/100 M.N.";

        /*if (centavos == 1)
			converted += " CON UN CENTAVO";
		if (centavos > 1)
			converted += " CON " + convertNumber(String.valueOf(centavos))
					+ "CENTAVOS";*/

        return converted;
    }
    /**
     * Convierte los trios de numeros que componen las unidades, las decenas y
     * las centenas del numero.
     * <p>
     * Creation date 3/05/2006 - 05:33:40 PM
     * 
     * @param number
     *            Numero a convetir en digitos
     * @return Numero convertido en letras
     * @since 1.0
     */

    private static string convierteNumero(string number)
    {
        string output = string.Empty;

        if (number.Length > 3)
            return "";

        if (getDigitAt(number, 2) != "0")
            output = CENTENAS[Convert.ToInt32(getDigitAt(number, 2)) - 1];

        int k = Convert.ToInt32(getDigitAt(number, 1)
                + getDigitAt(number, 0));

        if (k <= 20)
            output += UNIDADES[k];
        else
        {
            if (k > 30 && getDigitAt(number, 0) != "0")
                output += DECENAS[Convert.ToInt32(getDigitAt(number, 1)) - 2] + "Y "
                        + UNIDADES[Convert.ToInt32(getDigitAt(number, 0))];
            else
                output += DECENAS[Convert.ToInt32(getDigitAt(number, 1)) - 2]
                        + UNIDADES[Convert.ToInt32(getDigitAt(number, 0))];
        }

        // Caso especial con el 100
        if (getDigitAt(number, 2) == "1" && k == 0)
            output = "CIEN ";

        return output;
    }

    //Funcion que calcula el valor de Costo Anual Total
    public static string GeneraCAT(decimal cantEntre, decimal totalPagar, decimal pagoParc, int NoPagos, string Periodo)
    {
        int NoSemanas = 0;
        double[] valueArray;
        double CAT;
        valueArray = new double[NoPagos + 1];

        valueArray[0] = Convert.ToDouble(cantEntre * -1);

        for (int i = 0; i < NoPagos - 1; i++)
        {
            valueArray[i + 1] = Convert.ToDouble(pagoParc);
        }
        valueArray[NoPagos] = Convert.ToDouble((totalPagar - (pagoParc * (NoPagos - 1))));

        Periodo = Periodo.ToLower();
        if ((Periodo == "semanal") || (Periodo == "semanas"))
            NoSemanas = 52;
        else if (Periodo == "bisemanal")
            NoSemanas = 26;
        else if (Periodo == "quincenal")
            NoSemanas = 24;
        else if ((Periodo == "mensual") || (Periodo == "meses"))
            NoSemanas = 12;
        else if (Periodo == "trimestres")
            NoSemanas = 4;
        else if (Periodo == "semestres")
            NoSemanas = 2;

        if (cantEntre != totalPagar)
        {
            double guess = 0.1;
            double CalcRetRate = (Financial.IRR(ref valueArray, guess)) * 100;
            string IRR = CalcRetRate.ToString("#0.000");
            CAT = (Math.Pow(1 + (double.Parse(IRR) / 100), NoSemanas) - 1) * 100;
        }
        else
        {
            CAT = 0;
        }
        return CAT.ToString("##0.0");
    }

    /**
     * Retorna el digito numerico en la posicion indicada de derecha a izquierda
     * <p>
     * Creation date 3/05/2006 - 05:26:03 PM
     * 
     * @param origin
     *            Cadena en la cual se busca el digito
     * @param position
     *            Posicion de derecha a izquierda a retornar
     * @return Digito ubicado en la posicion indicada
     * @since 1.0
     */
    public static string getDigitAt(string origin, int position)
    {
        if (origin.Length > position && position >= 0)
            return origin.Substring(origin.Length - position - 1, 1);
        return "0";
    }

    public static byte[] ImagenToByte(System.Drawing.Image img)
    {
        ImageConverter converter = new ImageConverter();
        return (byte[])converter.ConvertTo(img, typeof(byte[]));
    }

    #endregion
}
