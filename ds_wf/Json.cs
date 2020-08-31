using System;
using System.Data;
using System.IO;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Formatting = Newtonsoft.Json.Formatting;

namespace DeltaCore.WorkFlow
{
	public static class Json
	{
		public static string Serialize(object value)
		{
			Type type = value.GetType();
			var json = new Newtonsoft.Json.JsonSerializer
			           	{
			           		NullValueHandling = NullValueHandling.Ignore,
			           		ObjectCreationHandling = Newtonsoft.Json.ObjectCreationHandling.Replace,
			           		MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore,
			           		ReferenceLoopHandling = ReferenceLoopHandling.Ignore
			           	};
			if (type == typeof (DataTable))
				json.Converters.Add(new DataTableConverter());
			else if (type == typeof (DataSet))
				json.Converters.Add(new DataSetConverter());

			var sw = new StringWriter();
			var writer = new JsonTextWriter(sw) {Formatting = Formatting.Indented, QuoteChar = '"'};
			json.Serialize(writer, value);
			string output = sw.ToString();
			writer.Close();
			sw.Close();
			return output;
		}

		public static object Deserialize(string jsonText, Type valueType)
		{
			var json = new JsonSerializer
			           	{
			           		NullValueHandling = NullValueHandling.Ignore,
			           		ObjectCreationHandling = ObjectCreationHandling.Replace,
			           		MissingMemberHandling = MissingMemberHandling.Ignore,
			           		ReferenceLoopHandling = ReferenceLoopHandling.Ignore
			           	};
			var sr = new StringReader(jsonText);
			var reader = new JsonTextReader(sr);
			object result = json.Deserialize(reader, valueType);
			reader.Close();

			return result;
		}

		public static string DataTable2Json(DataTable table)
		{
			return JsonConvert.SerializeObject(table, Formatting.Indented);
		}

		public static XmlDocument JsonToXml(string json)
		{
			return (XmlDocument)JsonConvert.DeserializeXmlNode(json);
		}

		public static string Xml2Json(XmlDocument xDoc)
		{
			return JsonConvert.SerializeXmlNode(xDoc);
		}
	}
	
	/*
	Public Class PropiedadJSON
    Private _NO_PROPIEDAD As String
    Private _NO_VALOR As String
    Private _NO_CLASE As String
    Private _NO_TIPO_DATO As String
    Public Property NO_TIPO_DATO() As String
        Get
            Return _NO_TIPO_DATO
        End Get
        Set(ByVal value As String)
            _NO_TIPO_DATO = value
        End Set
    End Property
    Public Property NO_CLASE() As String
        Get
            Return _NO_CLASE
        End Get
        Set(ByVal value As String)
            _NO_CLASE = value
        End Set
    End Property
    Public Property NO_PROPIEDAD() As String
        Get
            Return _NO_PROPIEDAD
        End Get
        Set(ByVal value As String)
            _NO_PROPIEDAD = value
        End Set
    End Property
    Public Property NO_VALOR() As String
        Get
            Return _NO_VALOR
        End Get
        Set(ByVal value As String)
            _NO_VALOR = value
        End Set
    End Property
    Public ReadOnly Property GetPropiedad() As String
        Get
            If _NO_TIPO_DATO.Equals("System.String") Then
                Return String.Concat("""", _NO_PROPIEDAD, """:""", _NO_VALOR, """,")
            ElseIf _NO_TIPO_DATO.Equals("System.Int32") Then
                If CType(_NO_VALOR, Integer) > 0 Then
                    Return String.Concat("""", _NO_PROPIEDAD, """:", _NO_VALOR, ",")
                End If
            ElseIf _NO_TIPO_DATO.Equals("System.Decimal") Then
                If CType(_NO_VALOR, Decimal) > 0 Then
                    Return String.Concat("""", _NO_PROPIEDAD, """:", _NO_VALOR, ",")
                End If
            End If
            Return ""
        End Get
    End Property
End Class*/
}