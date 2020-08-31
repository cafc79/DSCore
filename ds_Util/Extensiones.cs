using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace DeltaCore.Utilities
{
    public static class Extensiones
    {        
        public static string S(this XmlNode nodo, string atributo) //Atributo a String
        {
            if (!atributo.StartsWith("@")) throw new Exception(atributo + " sin @");
            return nodo.Attributes[atributo.Replace("@", "")].Value;
        }

        public static Guid G(this XmlNode nodo, string atributo) //Atributo a String
        {
            return new Guid(nodo.Attributes[atributo].Value);
        }

        public static int I(this XmlNode nodo, string atributo) //Atributo a String
        {
            return int.Parse(nodo.S(atributo));
        }

        public static decimal D(this XmlNode nodo, string atributo) //Atributo a String
        {
            return decimal.Parse(nodo.S(atributo));
        }

        public static decimal Porciento(this XmlNode nodo, string atributo) //Atributo a String
        {
            return decimal.Parse(nodo.S(atributo)) / 100m;
        }

        public static bool EsJson(this XmlNode nodo, string atributo) //Atributo a String
        {
            return nodo.S(atributo).Trim().StartsWith("{");
        }

        public static bool Tiene(this XmlNode nodo, string atributo) //Atributo a String
        {
            return nodo.Attributes[atributo.TrimStart('@')] != null;
        }


        public static string QuitarComa(this string numero) //request a String
        {
            return numero.Replace(",", "");
        }


        public static bool EstaVacio(this object valor)
        {
            if (valor.EsNulo()) return true;
            if (valor is string) return valor.Equals(string.Empty);
            if (valor is Guid) return valor.Equals(Guid.Empty);
            if (valor is DateTime) return ((DateTime?)valor).EsNulo();
            if (valor is int?) return valor.EsNulo();
            if (valor is decimal?) return valor.EsNulo();

            throw new Exception("No existe este tipo");
        }

        public static bool EstaLleno(this object valor)
        {
            return !valor.EstaVacio();
        }

        public static bool EsNulo(this object objeto)
        {
            if (objeto == null) return true;
            if (objeto.GetType().Name.Equals("DBNull")) return true;

            return false;
        }

        public static bool NoEsNulo(this object objeto)
        {
            if (objeto != null) return true;
            return !objeto.GetType().Name.Equals("DBNull");
        }

        public static bool NoEs(this object valor, object comparar)
        {
            return valor.ToString() != comparar.ToString();
        }

        public static bool Es(this object valor, object comparar)
        {
            return valor.ToString() == comparar.ToString();
        }

        public static int AEntero(this string valor)
        {
            return int.Parse(valor.EstaVacio() ? "0" : valor.Replace(",", ""));
        }

        public static int? ANullEntero(this string valor)
        {
            return valor.EstaVacio() ? new Nullable<int>() : new Nullable<int>(int.Parse(valor));
        }

        public static decimal ADecimal(this string valor)
        {
            return decimal.Parse(valor.EstaVacio() ? "0" : valor);
        }

        public static decimal? ANullDecimal(this string valor)
        {
            return valor.EstaVacio() ? new Nullable<decimal>() : new Nullable<decimal>(decimal.Parse(valor));
        }

        public static Regex EsGuid =
            new Regex(@"^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$",
                      RegexOptions.Compiled);

        public static Guid AGuid(this string valor)
        {
            return new Guid(valor);
        }

        public static Guid DameGuid(this string valor)
        {
            return EsGuid.IsMatch(valor) ? new Guid(valor) : Guid.NewGuid();
        }

        public static string ANumero(this int valor)
        {
            return ((decimal)valor).ANumero();
        }

        public static string ANumero(this decimal valor)
        {
            return valor - (Math.Floor(valor)) > 0 ? string.Format("{0:n}", valor) : string.Format("{0:n}", valor).Replace(".00", "");
        }

        public static string ANumero(this string valor)
        {
            return valor.ADecimal().ANumero();
        }

        public static string AMoneda(this decimal value)
        {
            return value >= 0 ? string.Format("{0:C}", value) : value.AMonedaNegativo();
        }

        public static string AMoneda(this decimal value, int precision)
        {
            return value >= 0 ? string.Format("{0:C" + precision + "}", value) : value.AMonedaNegativo(precision);
        }

        public static string AMoneda(this string value)
        {
            return value.ADecimal().AMoneda();
        }

        public static string AMonedaNegativo(this decimal value)
        {
            return string.Format((value != 0 ? "-" : "") + "{0:C}", value >= 0 ? value : value * -1);
        }

        public static string AMonedaNegativo(this decimal value, int precision)
        {
            return string.Format((value != 0 ? "-" : "") + "{0:C" + precision + "}", value >= 0 ? value : value * -1);
        }

        public static string ACamel(this string valor)
        {
            return valor.Substring(0, 1).ToLower() + valor.Substring(1);
        }

        public static string APascal(this string valor)
        {
            return valor.Substring(0, 1).ToUpper() + valor.Substring(1);
        }

        public static bool InicienCon(this string valor, params string[] compararCon)
        {
            return compararCon.Any(t => valor.StartsWith(t));
        }

        public static bool En(this string valor, params string[] compararCon)
        {
            return compararCon.Any(t => valor.Equals(t));
        }

        public static bool En(this int valor, params int[] compararCon)
        {
            return compararCon.Any(t => valor.Equals(t));
        }

        public static bool Entre(this decimal valor, decimal menorIncluido, decimal mayorIncluido)
        {
            return valor >= menorIncluido && valor <= mayorIncluido;
        }

        public static bool Entre(this int valor, decimal menorIncluido, decimal mayorIncluido)
        {
            return Entre((decimal)valor, menorIncluido, mayorIncluido);
        }

        public static string SiEs(this string valor, string comparador, string resultado)
        {
            return valor.Equals(comparador) ? resultado : "";
        }

        public static string SiEs(this string[] valor, string comparador, string resultado)
        {
            return valor.Contains(comparador) ? resultado : "";
        }

        public static string SiEs(this object valor, string comparador, string resultado)
        {
            if (valor is string)
                return ((string)valor).SiEs(comparador, resultado);
            if (valor is string[])
                return ((string[])valor).SiEs(comparador, resultado);
            return "";
        }

        public static string SiNoEs(this string valor, string comparador, string resultado)
        {
            return !valor.Equals(comparador) ? resultado : "";
        }

        public static bool EsNumero(this string valor)
        {
            double ReturnedNumber;
            return Double.TryParse(valor, NumberStyles.Any, NumberFormatInfo.InvariantInfo, out ReturnedNumber);
        }

        public static string Diferencia(this string foco, string comparador)
        {
            return foco.ADecimal().Diferencia(comparador.ADecimal());
        }

        public static string Diferencia(this int foco, int comparador)
        {
            return ((decimal)foco).Diferencia((decimal)comparador);
        }

        public static string Diferencia(this decimal foco, decimal comparador)
        {
            return (foco - comparador).ToString().ANumero();
        }

        public static string DMoneda(this string foco, string comparador)
        {
            return foco.ADecimal().DMoneda(comparador.ADecimal());
        }

        public static string DMoneda(this int foco, int comparador)
        {
            return ((decimal)foco).DMoneda((decimal)comparador);
        }

        public static string DMoneda(this decimal foco, decimal comparador)
        {
            return "$" + foco.Diferencia(comparador);
        }

        public static string Porcentaje(this string foco, string comparador)
        {
            return foco.ADecimal().Porcentaje(comparador.ADecimal());
        }

        public static string Porcentaje(this int foco, int comparador)
        {
            return ((decimal)foco).Porcentaje((decimal)comparador);
        }

        public static string Porcentaje(this decimal foco, decimal comparador)
        {
            if (foco == 0 && comparador == 0) return "0%";
            if (foco != 0 && comparador == 0) return "100%";
            return (Math.Round(foco * 100m / comparador, 2) - 100m) + "%";
        }

       

        public static string AcentosAHtml(this string texto)
        {
            return
                texto.Replace("á", "&aacute;").Replace("é", "&eacute;").Replace("í", "&iacute;").Replace("ó", "&oacute;").Replace(
                    "ú", "&uacute;").Replace("ñ", "&ntilde;");
        }

        public static string AListaSqlIn(this string elementos)
        {
            return
                elementos.Split(',').Aggregate("", (current, elemento) => current + string.Format("'{0}',", elemento.Trim())).
                    TrimEnd(',');
        }
        
    }
}