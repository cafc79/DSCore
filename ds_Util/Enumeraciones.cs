using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace DeltaCore.Utilities
{
	[AttributeUsage(AttributeTargets.Enum | AttributeTargets.Field, AllowMultiple = false)]
	public sealed class EnumDescriptionAttribute : Attribute
	{
		public string Description { get; private set; }

		public EnumDescriptionAttribute(string description)
		{
			this.Description = description;
		}
	}
	
	public class Enumeraciones
	{
		public static string GetEnumDescription(Enum value)
		{
			FieldInfo fi = value.GetType().GetField(value.ToString());
			var attributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

			if (attributes != null && attributes.Length > 0)
				return attributes[0].Description;
			return value.ToString();
		}

		public static string GetDescription(Enum value)
		{
			if (value == null)
			{
				throw new ArgumentNullException("value");
			}
			string description = value.ToString();
			FieldInfo fieldInfo = value.GetType().GetField(description);
			EnumDescriptionAttribute[] attributes = (EnumDescriptionAttribute[])
																							fieldInfo.GetCustomAttributes(typeof(EnumDescriptionAttribute), false);
			if (attributes != null && attributes.Length > 0)
			{
				description = attributes[0].Description;
			}
			return description;
		}

		public static IList ToList(Type type)
		{
			if (type == null)
			{
				throw new ArgumentNullException("type");
			}
			var list = new ArrayList();
			Array enumValues = Enum.GetValues(type);
			foreach (Enum value in enumValues)
			{
				list.Add(new KeyValuePair<Enum, string>(value, GetDescription(value)));
			}
			return list;
		}

		public static IEnumerable<T> EnumToList<T>()
		{
			Type enumType = typeof(T);
			if (enumType.BaseType != typeof(Enum))
				throw new ArgumentException("T debe ser de tipo System.Enum");

			Array enumValArray = Enum.GetValues(enumType);
			List<T> enumValList = new List<T>(enumValArray.Length);
			enumValList.AddRange(from int val in enumValArray select (T) Enum.Parse(enumType, val.ToString()));
			return enumValList;
		}

		public static T GetAttributeOfType<T>(Enum enumVal) where T : System.Attribute
		{
			var type = enumVal.GetType();
			var memInfo = type.GetMember(enumVal.ToString());
			var attributes = memInfo[0].GetCustomAttributes(typeof(T), false);
			return (T)attributes[0];
		}
	}

	public enum ValidCert
	{
		[EnumDescription("Valido")]
		cC,
		[EnumDescription("Vencido")]
		cV,
		[EnumDescription("Información Inexistente")]
		cII,
		[EnumDescription("Host Invalido")]
		cHI,
		[EnumDescription("Información Alterada")]
		cIA
	}
	
	public enum NetProtocol
	{
		[Description("WebSocket")]
		ws,
		[Description("HTTP")]
		http,
		[Description("FTP")]
		ftp,
		[Description("FILE")]
		file
	}
}