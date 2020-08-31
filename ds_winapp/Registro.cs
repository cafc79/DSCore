using System;
using Microsoft.Win32;

namespace DeltaCore.WinApp
{
	public static class Registro
	{
		public static bool ponerEnInicio(string nombreClave, string nombreApp)
		{
			// Resgistrará en Inicio del registro la aplicación indicada
			// Devuelve True si todo fue bien, False en caso contrario
			// Guardar la clave en el registro
			// HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run
			try
			{
				RegistryKey runK = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
				// añadirlo al registro
				// Si el path contiene espacios se debería incluir entre comillas dobles
				if (nombreApp.StartsWith("\"") == false && nombreApp.IndexOf(" ") > -1)
				{
					nombreApp = "\"" + nombreApp + "\"";
				}
				runK.SetValue(nombreClave, nombreApp);
				return true;
			}
			catch (Exception ex)
			{
				Console.WriteLine("ERROR al guardar en el registro.{0}Seguramente no tienes privilegios suficientes.{0}{1}{0}---xxx---{2}", '\n', ex.Message, ex.StackTrace);
				return false;
			}
		}

		public static bool quitarDeInicio(string nombreClave)
		{
			// Quitará de Inicio la aplicación indicada
			// Devuelve True si todo fue bien, False en caso contrario
			// Si la aplicación no estaba en Inicio, devuelve True salvo que se produzca un error
			//
			try
			{
				RegistryKey runK = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
				// quitar la clave indicada del registo
				runK.DeleteValue(nombreClave, false);
				return true;
			}
			catch (Exception ex)
			{
				Console.WriteLine("ERROR al eliminar la clave del registro.{0}Seguramente no tienes privilegios suficientes.{0}{1}{0}---xxx---{2}", '\n', ex.Message, ex.StackTrace);
				return false;
			}
		}

		public static string comprobarEnInicio(string nombreClave)
		{
			// Comprobará si la clave indicada está asignada en Inicio
			// en caso de ser así devolverá el contenido,
			// en caso contrario devolverá una cadena vacia
			// Si se produce un error, se devolverá la cadena de error
			try
			{
				RegistryKey runK = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false);
				// comprobar si está
				return runK.GetValue(nombreClave, "").ToString();
			}
			catch (Exception ex)
			{
				return String.Format("ERROR al leer el valor de la clave del registro.{0}Seguramente no tienes privilegios suficientes.{0}{1}{0}---xxx---{2}", '\n', ex.Message, ex.StackTrace);
			}
		}
	}
}