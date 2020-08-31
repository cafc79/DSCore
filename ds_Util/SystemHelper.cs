using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.VisualBasic;
using System;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Web;

namespace DeltaCore.Utilities
{
	public class SystemHelper
	{
		// Methods
		public static Dictionary<string, string> GetPropiedadesValores(object o)
		{
			return o.GetType().GetMembers().Where(mi => mi.MemberType == MemberTypes.Property).OfType<PropertyInfo>().ToDictionary(pi => pi.Name, pi => pi.GetValue(o, null).ToString());
		}

		public static List<string> GetPropiedades(object o)
		{
			return o.GetType().GetMembers().Where(mi => mi.MemberType == MemberTypes.Property).OfType<PropertyInfo>().Select(pi => pi.Name).ToList();
		}

		public static object GetPropertyValue(object src, string propName)
		{
			return src.GetType().GetProperty(propName).GetValue(src, null);
		}

		public static Object GetPropertyValue(String name, Object obj)
		{
			foreach (String part in name.Split('.'))
			{
				if (obj == null) { return null; }

				Type type = obj.GetType();
				PropertyInfo info = type.GetProperty(part);
				if (info == null) { return null; }

				obj = info.GetValue(obj, null);
			}
			return obj;
		}

		public static string WhoCalledMe()
		{
			var st = new StackTrace();
			var sf = st.GetFrame(1); //StackFrame 
			var mb = sf.GetMethod(); //MethodBase
			return mb.Name;
		}

		public static List<string> GetMethodName(object o)
		{
			return o.GetType().GetMembers().Where(mi => mi.MemberType == MemberTypes.Property).OfType<PropertyInfo>().Select(pi => pi.Name).ToList();
		}

		public static byte[] Serialize(object elObjeto)
		{
			try
			{
				var MemoryStream = new MemoryStream();
				var BinaryFormatter = new BinaryFormatter();
				BinaryFormatter.Serialize(MemoryStream, elObjeto);
				MemoryStream.Position = 0;
				byte[] ObjetoSerializado = new byte[MemoryStream.Length];
				MemoryStream.Read(ObjetoSerializado, 0, (int)MemoryStream.Length);
				MemoryStream.Close();
				return ObjetoSerializado;
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}

		public static object Deserializar(object elObjeto)
		{
			try
			{
				byte[] Bytes = (byte[])elObjeto;
				var MemoryStream = new MemoryStream();
				MemoryStream.Write(Bytes, 0, Bytes.Length);
				MemoryStream.Position = 0;
				var BinaryFormatter = new BinaryFormatter();
				return BinaryFormatter.Deserialize(MemoryStream);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}

		public static string AppTitle()
		{
			//'From http://www.dotnet247.com/247reference/msgs/15/78041.aspx
			if (SystemHelper.IsWebApplication())
			{

				char[] chArray1 = new char[1] {'/'};
				return HttpContext.Current.Request.ApplicationPath.TrimStart(chArray1); //'returns name like "/vLearning"
			}
			//'or sProductName = Process.GetCurrentProcess().ProcessName
			//' or Path.GetFileNameWithoutExtension(AppDomain.CurrentDomain.FriendlyName)
			Assembly asm = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
			return AssemblyTitle(asm);
		}

		public static string AssemblyTitle(Assembly asm)
		{
			//'From http://www.dotnet247.com/247reference/msgs/15/78041.aspx
			AssemblyTitleAttribute attribute1 = (AssemblyTitleAttribute) Attribute.GetCustomAttribute(asm, typeof (AssemblyTitleAttribute));
			//'Application.ProductName is not working in Dll
			//'Assembly.GetEntryAssembly().FullName not reliable ?? see http://weblogs.asp.net/asanto/posts/26710.aspx
			string sRet = attribute1.Title;
			//if (DataHelper.IsNullOrEmpty(sRet))
			//	sRet = asm.GetName().Name;
			return sRet;
		}

		public static string GetAssemblyVersion(Assembly asm)
		{
			FileVersionInfo info1 = FileVersionInfo.GetVersionInfo(asm.Location);
			return info1.ProductVersion;
		}

		public static bool IsWebApplication()
		{
			//'http://groups.google.com.au/groups?hl=en&lr=&ie=UTF-8&oe=UTF-8&selm=ebVausWRBHA.1408%40tkmsftngp05
			//issue: it will return wrong answer if called from async thread running from Web
			//try to check some System.AppDomain.CurrentDomain properties
		    bool bIsASP = !Information.IsNothing(HttpContext.Current);
		    return bIsASP;
		}

		public static string GetWebRootFolder(Assembly ExecutableAssembly)
		{
			//Debug.Assert(IsWebApplication()); can be called frm installer or tester
			string sPath = ExecutableAssembly.Location;
			//remove assembly name
			sPath = sPath.Remove(sPath.LastIndexOf(@"\"), Strings.Len(sPath) - sPath.LastIndexOf(@"\"));
			//remove BIN folder name
			if ((sPath.ToUpper().EndsWith(@"\BIN"))) //15/6/2006
			{
				sPath = sPath.Remove(sPath.LastIndexOf(@"\"), Strings.Len(sPath) - sPath.LastIndexOf(@"\"));
			}
			Debug.Assert(Directory.Exists(sPath));
			//Debug.Assert(sPath==System.AppDomain.CurrentDomain.BaseDirectory);//can be called frm installer or tester
			return sPath;
		}

		//See DateTime.Compare
		public static DateTime Min(DateTime t1, DateTime t2)
		{
			return DateTime.Compare(t1, t2) > 0 ? t2 : t1;
		}

        public string DoubleToHex(double value, int maxDecimals)
        {
            string result = string.Empty;
            if (value < 0)
            {
                result += "-";
                value = -value;
            }
            if (value > ulong.MaxValue)
            {
                result += double.PositiveInfinity.ToString();
                return result;
            }
            ulong trunc = (ulong)value;
            result += trunc.ToString("X");
            value -= trunc;
            if (value == 0)
            {
                return result;
            }
            result += ".";
            byte hexdigit;
            while ((value != 0) && (maxDecimals != 0))
            {
                value *= 16;
                hexdigit = (byte)value;
                result += hexdigit.ToString("X");
                value -= hexdigit;
                maxDecimals--;
            }
            return result;
        }

        public ulong HexToUInt64(string hex)
        {
            ulong result;
            if (ulong.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            throw new ArgumentException("Cannot parse hex string.", "hex");
        }
    }
}