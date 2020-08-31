using System;
using System.Collections.Generic;
using System.Xml;


namespace DeltaCore.Utilities
{
  public class Utilidades
  {
    public string xMeses = "<xml><mes>Ene</mes><mes>Feb</mes><mes>Mar</mes><mes>Abr</mes><mes>May</mes><mes>Jun</mes><mes>Jul</mes><mes>Ago</mes><mes>Sep</mes><mes>Oct</mes><mes>Nov</mes><mes>Dic</mes></xml>";
    public string xMesesComp = "<xml><mes>Enero</mes><mes>Febrero</mes><mes>Marzo</mes><mes>Abril</mes><mes>Mayo</mes><mes>Junio</mes><mes>Julio</mes><mes>Agosto</mes><mes>Septiembre</mes><mes>Octubre</mes><mes>Noviembre</mes><mes>Diciembre</mes></xml>";

    public XmlDocument Meses() {
      var xDocs = new XmlDocument();
      xDocs.LoadXml(xMeses);
      return xDocs;
    }

    public String Mes_MMM(int Index) {
      var xDocs = new XmlDocument();
      xDocs.LoadXml(xMeses);
      XmlNodeList xNode = xDocs.DocumentElement.SelectNodes("mes");
      return xNode.Item(Index - 1).InnerText;
    }

    public string Mes(int Index) {
      var xDocs = new XmlDocument();
      xDocs.LoadXml(xMesesComp);
      XmlNodeList xNode = xDocs.DocumentElement.SelectNodes("mes");
      return xNode.Item(Index - 1).InnerText;
    }

    public int Mes_Int(string Mes, bool format) {
      var xDocs = new XmlDocument();
    	xDocs.LoadXml(format ? xMeses : xMesesComp);
    	XmlNodeList xNode = xDocs.DocumentElement.SelectNodes("mes");
      bool Status = false;
      int index = 0;
      while (!Status) {
        if (xNode.Item(index).InnerText == Mes) {
          Status = true;
        }
        index++;
      	if (index <= 12) continue;
      	index = 0;
      	Status = true;
      }
      return index;
    }

    public int LastDayM(int Anio, int Mes) {
      var ldm = new DateTime(Anio, Mes, 1).AddMonths(1).AddDays(-1);      
      return ldm.Day;
    }

    public bool ValidaNumero(string strCadena) {
      bool bOk = true;
      bool bCarOk;
      const string checkOK = "0123456789.";
      string ch = "";

      bOk = true;
      for (int i = 0; i < strCadena.Length; i++) {
        ch = strCadena.Substring(i, 1);
        bCarOk = false;
        for (int j = 0; j < checkOK.Length; j++) {
          if (ch == checkOK.Substring(j, 1)) {
            bCarOk = true;
            j = checkOK.Length;
            break;
          }
        }
        if (bCarOk == false) {
          i = strCadena.Length;
          bOk = false;
        }
      }
      return bOk;
    }

    public List<String> CsvLine2Array(string inputLine) {
      char[] chr = {','};    
      char str = '"';
      List<String> value;
      string[] spliteado = null;
      if (inputLine.Contains("\"")) {
        int strLength = 0;
        int comaPos = 0;
        int comillaPos = 0;
        value = new List<String>();
        while (strLength <= inputLine.Length) {
          comaPos = (inputLine.IndexOf(chr[0], strLength) < 0) ? inputLine.Length : inputLine.IndexOf(chr[0], strLength);
          comillaPos = (inputLine.IndexOf(str, strLength) < 0) ? inputLine.Length : inputLine.IndexOf(str, strLength);
          if (comaPos <= comillaPos) {
            value.Add(inputLine.Substring(strLength, comaPos - strLength));
            strLength = comaPos + 1;
          }
          else {
            comillaPos = inputLine.IndexOf(str, strLength + 1);
            string strField = inputLine.Substring(strLength + 1, comillaPos - strLength - 1);
            value.Add(strField);
            strLength = comillaPos + 2;
          }
        }        
      }
      else {
        spliteado = inputLine.Split(chr);
        value = new List<String>(spliteado);
      }      
      return value;
    }
  }
}