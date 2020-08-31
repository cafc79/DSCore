using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DeltaCore.Utilities;

namespace DeltaCore.FileAccess
{
	public class XmlManager
	{
		public XmlDocument xDoc;
		private Dictionary<string, Dictionary<string, object>> JsonObjectsList = new Dictionary<string, Dictionary<string, object>>();

		public XmlManager()
		{
			xDoc = new XmlDocument();
		}

		public XmlManager(string xml)
		{
			xDoc = new XmlDocument();
			xDoc.LoadXml(xml);
		}

		public static XmlManager CargarXml(string xml)
		{
			return new XmlManager(xml);
		}

		public static XmlManager CargarRuta(string ruta)
		{
			return new XmlManager(File.ReadAllText(ruta));
		}

		public void CrearNodo(string ruta, string nombre, params object[] atributos)
		{
			XmlNode NuevoNodo = xDoc.CreateNode(XmlNodeType.Element, nombre, "");
			for (int i = 0; i < atributos.Length; i += 2)
			{
				XmlAttribute NuevoAtributo = xDoc.CreateAttribute(atributos[i].ToString().TrimStart('@'));
				NuevoAtributo.Value = (atributos[i + 1] ?? "").ToString();
				NuevoNodo.Attributes.Append(NuevoAtributo);
			}
			if (ruta.EstaVacio()) xDoc.AppendChild(NuevoNodo);
			else xDoc.SelectSingleNode(ruta).AppendChild(NuevoNodo);
		}

		public void CrearNodo(string ruta, string nombre)
		{
			CrearNodo(ruta, nombre, new string[0]);
		}

		public void CrearAtributos(string ruta, params object[] atributos)
		{
			XmlNode Nodo = DameNodo(ruta);
			for (int i = 0; i < atributos.Length; i += 2)
			{
				XmlAttribute NuevoAtributo = xDoc.CreateAttribute(atributos[i].ToString().TrimStart('@'));
				NuevoAtributo.Value = (atributos[i + 1] ?? "").ToString();
				Nodo.Attributes.Append(NuevoAtributo);
			}
		}

		public bool TieneAtributo(string ruta)
		{
			string[] PartesDeRuta = new string[]
			                        	{
			                        		ruta.Substring(0, ruta.LastIndexOf("@")),
			                        		ruta.Substring(ruta.LastIndexOf("@"))
			                        	};
			XmlNode Nodo = xDoc.SelectSingleNode(PartesDeRuta[0].EstaVacio() ? "/*" : PartesDeRuta[0]);
			if (Nodo == null) return false;
			return Nodo.Tiene(PartesDeRuta[1]);
		}

		public void InsertarNodo(string ruta, XmlNode nuevoNodo)
		{
			XmlNode NodoPadre = DameNodo(ruta);
			XmlNode NuevoNodoImportado = xDoc.ImportNode(nuevoNodo, true);
			NodoPadre.AppendChild(NuevoNodoImportado);
		}

		public void RemplazarNodo(string ruta, XmlNode nuevoNodo)
		{
			XmlNode viejoNodo = DameNodo(ruta);
			XmlNode NuevoNodoImportado = xDoc.ImportNode(nuevoNodo, true);
			viejoNodo.ParentNode.ReplaceChild(NuevoNodoImportado, viejoNodo);
		}

		public void EliminarNodo(string ruta)
		{
			XmlNode NodoARemover = DameNodo(ruta);
			NodoARemover.ParentNode.RemoveChild(NodoARemover);
		}

		public void EliminarAtributo(string ruta)
		{
			string[] PartesDeRuta = new string[]
			                        	{
			                        		ruta.Substring(0, ruta.LastIndexOf("@")),
			                        		ruta.Substring(ruta.LastIndexOf("@") + 1)
			                        	};

			XmlNode NodoARemover = DameNodo(PartesDeRuta[0]);
			NodoARemover.Attributes.RemoveNamedItem(PartesDeRuta[1]);
		}

		public void Modificar(string ruta, object valor)
		{
			string[] PartesDeRuta = new string[]
			                        	{
			                        		ruta.Substring(0, ruta.LastIndexOf("@")),
			                        		ruta.Substring(ruta.LastIndexOf("@") + 1)
			                        	};

			XmlNode Nodo = xDoc.SelectSingleNode(PartesDeRuta[0].EstaVacio() ? "/*" : PartesDeRuta[0]);
			if (Nodo.EsNulo()) throw new Exception("No existe \n" + PartesDeRuta[0]);

			if (PartesDeRuta[1].IndexOf(".") > 0)
			{
				List<string> AtributoYRutaJSON = new List<string>(PartesDeRuta[1].Split('.'));

				ModificaValorDeJson(AgregarJSonObject(PartesDeRuta[0], AtributoYRutaJSON[0], Nodo), AtributoYRutaJSON,
				                    valor.ToString());
				return;
			}

			Nodo.Attributes[PartesDeRuta[1]].Value = valor.ToString();
		}

		public XmlNode DameNodo(string ruta)
		{
			return xDoc.SelectSingleNode(ruta);
		}

		public XmlNodeList DameNodos(string ruta)
		{
			return xDoc.SelectNodes(ruta);
		}

		private string RutaDeNodo(XmlNode nodo)
		{
			return
				((nodo.ParentNode != null && nodo.ParentNode.NodeType == XmlNodeType.Element ? RutaDeNodo(nodo.ParentNode) : "") +
				 "/" + nodo.Name).TrimStart('/');
		}

		public void GrabarJsonModificados()
		{
			foreach (KeyValuePair<string, Dictionary<string, object>> JsonObject in JsonObjectsList)
			{
				var JsonStringBuilder = new StringBuilder();
				ConstruyeJson(JsonObject.Value, JsonStringBuilder);
				Modificar(JsonObject.Key, JsonStringBuilder.ToString());
			}
		}

		private void ModificaValorDeJson(Dictionary<string, object> json, List<string> AtributoYRutaJSON, string valor)
		{
			AtributoYRutaJSON.RemoveAt(0);
			if (AtributoYRutaJSON.Count == 1)
			{
				json[AtributoYRutaJSON[0]] = valor;
				return;
			}
			ModificaValorDeJson((Dictionary<string, object>) json[AtributoYRutaJSON[0]], AtributoYRutaJSON, valor);
		}

		private string DameValorDeJson(Dictionary<string, object> json, List<string> AtributoYRutaJson)
		{
			AtributoYRutaJson.RemoveAt(0);
			return AtributoYRutaJson.Count == 1 ? json[AtributoYRutaJson[0]].ToString().Trim('\'') : DameValorDeJson((Dictionary<string, object>) json[AtributoYRutaJson[0]], AtributoYRutaJson);
		}

		private Dictionary<string, object> AgregarJSonObject(string ruta, string atributo, XmlNode nodo)
		{
			string RutaDeJson = ruta + "@" + atributo;
			if (!JsonObjectsList.ContainsKey(RutaDeJson))
				JsonObjectsList.Add(RutaDeJson, JsonADictionary(nodo.S("@" + atributo)));

			return JsonObjectsList[RutaDeJson];
		}

		private void ConstruyeJson(Dictionary<string, object> jsonObject, StringBuilder jsonStringBuilder)
		{
			jsonStringBuilder.Append("{");
			foreach (KeyValuePair<string, object> JasonObjectPair in jsonObject)
			{
				if (JasonObjectPair.Value is Dictionary<string, object>)
				{
					jsonStringBuilder.AppendFormat("{0}:", JasonObjectPair.Key);
					ConstruyeJson((Dictionary<string, object>) JasonObjectPair.Value, jsonStringBuilder);
				}
				else
				{
					jsonStringBuilder.AppendFormat("{0}:'{1}'", JasonObjectPair.Key, JasonObjectPair.Value);
				}
				if (jsonObject.Keys.Last() != JasonObjectPair.Key) jsonStringBuilder.Append(",");
			}
			jsonStringBuilder.Append("}");
		}

		private Dictionary<string, object> JsonADictionary(string json)
		{
			var Elemento = new Dictionary<string, object>();
			json = json.Trim();
			if (json.StartsWith("{"))
			{
				json = json.Remove(0, 1);
				if (json.EndsWith("}"))
					json = json.Substring(0, json.Length - 1);
			}
			json = json.Trim();

			while (json.Length > 0)
			{
				var InicioDeValor = json.Substring(0, json.IndexOf(':'));
				json = json.Substring(InicioDeValor.Length);
				var SiguienteComa = json.IndexOf(',');
				string FinalDeValor;

				if (SiguienteComa > -1)
				{
					FinalDeValor = json.Substring(0, SiguienteComa);
					json = json.Substring(FinalDeValor.Length);
				}
				else
				{
					FinalDeValor = json;
					json = string.Empty;
				}

				var SiguienteLlave = FinalDeValor.IndexOf('{');
				if (SiguienteLlave > -1)
				{
					var NumeroDeLlaves = 1;
					while (FinalDeValor.Substring(SiguienteLlave + 1).IndexOf("{") > -1)
					{
						NumeroDeLlaves++;
						SiguienteLlave = FinalDeValor.Substring(SiguienteLlave + 1).IndexOf("{");
					}
					while (NumeroDeLlaves > 0)
					{
						FinalDeValor += json.Substring(0, json.IndexOf('}') + 1);
						json = json.Remove(0, json.IndexOf('}') + 1);
						NumeroDeLlaves--;
					}
				}

				json = json.Trim();
				if (json.StartsWith(",")) json = json.Remove(0, 1);
				json.Trim();

				var ParNombreValor = (InicioDeValor + FinalDeValor).Trim();

				var Nombre = ParNombreValor.Substring(0, ParNombreValor.IndexOf(":")).Trim();
				var Valor = ParNombreValor.Substring(Nombre.Length + 1).Trim();
				if (Nombre.StartsWith("\"") && Nombre.EndsWith("\""))
				{
					Nombre = Nombre.Substring(1, Nombre.Length - 2);
				}
				decimal VerificadorDeDecimal;
				if (Valor.StartsWith("\"") && Valor.StartsWith("\""))
				{
					Elemento.Add(Nombre, Valor.Substring(1, Valor.Length - 2));
				}
				else if (Valor.StartsWith("{") && Valor.EndsWith("}"))
				{
					Elemento.Add(Nombre, JsonADictionary(Valor));
				}
				else if (decimal.TryParse(Valor, out VerificadorDeDecimal))
				{
					Elemento.Add(Nombre, VerificadorDeDecimal);
				}
				else
				{
					Elemento.Add(Nombre, Valor);
				}
			}
			return Elemento;
		}

		public bool EstaVacio(string ruta)
		{
			return S(ruta).EstaVacio();
		}

		public string S(string ruta) //Ruta a string
		{
			return S(ruta, "");
		}

		public string S(string ruta, string valorPorDefecto)
		{
			string[] PartesDeRuta = new string[]
			                        	{
			                        		ruta.Substring(0, ruta.LastIndexOf("@")),
			                        		ruta.Substring(ruta.LastIndexOf("@"))
			                        	};
			XmlNode Nodo = xDoc.SelectSingleNode(PartesDeRuta[0].EstaVacio() ? "/*" : PartesDeRuta[0]);
			if (Nodo.EsNulo()) return valorPorDefecto;
			if (PartesDeRuta[1].IndexOf(".") > 0)
			{
				var atributoYRutaJson = new List<string>(PartesDeRuta[1].Split('.'));
				return DameValorDeJson(AgregarJSonObject(PartesDeRuta[0], atributoYRutaJson[0], Nodo), atributoYRutaJson);
			}
			return Nodo.Tiene(PartesDeRuta[1]) ? Nodo.S(PartesDeRuta[1]) : valorPorDefecto;
		}

		public override string ToString()
		{
			return xDoc.OuterXml;
		}
	}
}