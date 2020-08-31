using System;
using System.Linq;
using System.Management;

namespace DeltaCore.WinApp
{
	/*
C# Network Programming 
by Richard Blum

Publisher: Sybex 
ISBN: 0782141765
*/

	/// <summary>
	/// Generates a 16 byte Unique Identification code of a computer
	/// Example: 4876-8DB5-EE85-69D3-FE52-8CF7-395D-2EA9
	/// </summary>
	public class FingerPrint
	{
		#region Original Device ID Getting Code

		//Return a hardware identifier
		private static string identifier(string wmiClass, string wmiProperty, string wmiMustBeTrue)
		{
			string result = string.Empty;
			var mc = new ManagementClass(wmiClass);
			ManagementObjectCollection moc = mc.GetInstances();
			foreach (ManagementObject mo in from ManagementObject mo in moc where mo[wmiMustBeTrue].ToString() == "True" where string.IsNullOrEmpty(result) select mo)
			{
				try
				{
					result = mo[wmiProperty].ToString();
					break;
				}
				catch
				{
				}
			}
			return result;
		}

		//Return a hardware identifier
		private static string identifier(string wmiClass, string wmiProperty)
		{
			string result = "";
			var mc = new ManagementClass(wmiClass);
			ManagementObjectCollection moc = mc.GetInstances();
			foreach (ManagementObject mo in moc.Cast<ManagementObject>().Where(mo => result == ""))
			{
				try
				{
					result = mo[wmiProperty].ToString();
					break;
				}
				catch
				{
				}
			}
			return result;
		}

		public static string cpuId()
		{
			//Uses first CPU identifier available in order of preference
			//Don't get all identifiers, as it is very time consuming
			string retVal = identifier("Win32_Processor", "UniqueId");
			if (retVal == "") //If no UniqueID, use ProcessorID
			{
				retVal = identifier("Win32_Processor", "ProcessorId");
				if (retVal == "") //If no ProcessorId, use Name
				{
					retVal = identifier("Win32_Processor", "Name");
					if (retVal == "") //If no Name, use Manufacturer
					{
						retVal = identifier("Win32_Processor", "Manufacturer");
					}
					//Add clock speed for extra security
					retVal += identifier("Win32_Processor", "MaxClockSpeed");
				}
			}
			return retVal;
		}

		//BIOS Identifier
		public static string biosId()
		{
			return identifier("Win32_BIOS", "Manufacturer")
						 + identifier("Win32_BIOS", "SMBIOSBIOSVersion")
						 + identifier("Win32_BIOS", "IdentificationCode")
						 + identifier("Win32_BIOS", "SerialNumber")
						 + identifier("Win32_BIOS", "ReleaseDate")
						 + identifier("Win32_BIOS", "Version");
		}

		//Main physical hard drive ID
		public static string diskId()
		{
			return identifier("Win32_DiskDrive", "Model")
						 + identifier("Win32_DiskDrive", "Manufacturer")
						 + identifier("Win32_DiskDrive", "Signature")
						 + identifier("Win32_DiskDrive", "TotalHeads");
		}

		//Motherboard ID
		public static string baseId()
		{
			return identifier("Win32_BaseBoard", "Model")
						 + identifier("Win32_BaseBoard", "Manufacturer")
						 + identifier("Win32_BaseBoard", "Name")
						 + identifier("Win32_BaseBoard", "SerialNumber");
		}

		//Primary video controller ID
		public static string videoId()
		{
			return identifier("Win32_VideoController", "DriverVersion")
						 + identifier("Win32_VideoController", "Name");
		}

		//First enabled network card ID
		public static string macId()
		{
			return identifier("Win32_NetworkAdapterConfiguration",
												"MACAddress", "IPEnabled");
		}
		#endregion

		public ManagementObjectSearcher GetInternalInformation(bool logica)
		{
			if (logica)
				return new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
			return new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia");
		}

		public static string GetCPUId()
		{
			string cpuInfo = String.Empty;
			string temp = String.Empty;
			ManagementClass mc = new ManagementClass("Win32_Processor");
			ManagementObjectCollection moc = mc.GetInstances();
			foreach (ManagementObject mo in moc)
			{
				if ((cpuInfo == String.Empty))
				{
					cpuInfo = mo.Properties["ProcessorId"].Value.ToString();
				}
			}
			return cpuInfo;
		}

		public static string GetHDSerial()
		{
			ManagementObject disk = new ManagementObject("Win32_LogicalDisk.DeviceID=\"C:\"");
			PropertyData diskPropertyA = disk.Properties["VolumeSerialNumber"];
			return diskPropertyA.Value.ToString();
		}

		/*
		public static string GetProcessorID()
		{
			string sProcessorID = "";
			string sQuery = "SELECT ProcessorId FROM Win32_Processor";
			ManagementObjectSearcher oManagementObjectSearcher = new ManagementObjectSearcher(sQuery);
			ManagementObjectCollection oCollection = oManagementObjectSearcher.Get();
			foreach (ManagementObject oManagementObject in oCollection)
			{
				sProcessorID = (string)oManagementObject["ProcessorId"];
			}
			return (sProcessorID);
		}
		*/
	}
}