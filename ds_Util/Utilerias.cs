using System;
using System.Text.RegularExpressions;

namespace DeltaCore.Utilities
{
	public static class Validaciones
	{
		public static string ValidateFieldNonAvailable(string str)
		{
			return string.IsNullOrEmpty(str) ? "No Disponible" : str;
		}

		public static string ValidateFieldNonavailableZero(string str)
		{
			return string.IsNullOrEmpty(str) ? "0" : str;
		}

		public static int SeparaID_AutoComplete(string Texto)
		{
			int iID = 0;
			string[] sPartes = Texto.Split(Char.Parse("-"));
			if (sPartes.Length > 0)
				int.TryParse(sPartes[0], out iID);
			return iID;
		}
		
		/// <summary>
		/// Revisa el formato de un string con una expresion regular.
		/// </summary>
		/// <param name="entidadARevisar"></param>
		/// <param name="expresionRegular"></param>
		/// <returns>Si el formato es correcto envia true, de lo contrario enviara false</returns>
		public static bool validateFieldExpresion(string entidadARevisar, string expresionRegular)
		{
			return Regex.IsMatch(entidadARevisar, expresionRegular);
		}

		/// <summary>
		/// Obtiene una Fecha con hora y regresa solo la fecha.
		/// Por ejemplo: De la fecha 01/01/2009 12:00 a.m., se regresara la cadena 01/01/2009.
		/// </summary>
		/// <returns>Regresa la fecha sin la hora.</returns>
		public static string getDateWhitOutHour(string fecha)
		{
			try
			{
				char[] delimiterChars = { ' ' };
				string[] words = fecha.Split(delimiterChars);
				return words[0];
			}
			catch (IndexOutOfRangeException)
			{
				throw new ArgumentException("La fecha no tiene el formato requerido.");
			}
		}

		/// <summary>
		/// Valida que un campo exista o sea diferente de "";
		/// </summary>
		/// <param name="valorCampo">La variable que se desea verificar</param>
		/// <param name="nombreCampo">El string que se desea imprimir en el mensaje para decir que
		/// el campo es requerido</param>
		public static void isRequired(string valorCampo, string nombreCampo)
		{
			if (string.IsNullOrEmpty(valorCampo))
			{
				throw new ArgumentException("El campo " + nombreCampo + " es requerido.");
			}
		}

		public static void isComboRequired(int index, string nombreCampo)
		{
			if (index < 0)
			{
				throw new ArgumentException("Se debe seleccionar un valor: " + nombreCampo);
			}
		}
		/// <summary>
		/// Valida que una fecha tenga el formato dd/mm/aaaa
		/// </summary>
		/// <param name="fecha">La fecha que se desea revisar</param>
		public static void validateDateDDMMAAA(string fecha)
		{
			const string expresionRegular = @"^(0?[1-9]|1[0-9]|2|2[0-9]|3[0-1])(/|-)(0?[1-9]|1[0-2])(/|-)(\d{2}|\d{4})$";

			if (!validateFieldExpresion(fecha, expresionRegular))
			{
				throw new ArgumentException("La fecha " + fecha + " es incorrecta.");
			}
		}

		/// <summary>
		/// Valida que una fecha tenga el formato aaaa/mm/dd
		/// </summary>
		/// <param name="fecha">La fecha que se desea revisar</param>
		public static void validateDateAAAAMMDD(string fecha)
		{
			const string expresionRegular = @"^(\d{2}|\d{4})(/|-)(0?[1-9]|1[0-2])(/|-)(0?[1-9]|1[0-9]|2|2[0-9]|3[0-1])$";

			if (!validateFieldExpresion(fecha, expresionRegular))
			{
				throw new ArgumentException("La fecha " + fecha + " es incorrecta.");
			}
		}

		/// <summary>
		/// convierte una fecha dd/mm/aaaa en aaaa/mm/dd
		/// </summary>
		/// <param name="fecha">La fecha que se desea revisar</param> 
		public static string convertDate(string fecha)
		{
			fecha = getDateWhitOutHour(fecha);
			validateDateAAAAMMDD(fecha);
			char[] delimiterChars = { '/', '-' };
			string[] words = fecha.Split(delimiterChars);
			string fechaYearMonthDay = words[2] + "-" + words[1] + "-" + words[0];
			return fechaYearMonthDay;
		}

		/// <summary>
		/// convierte una fecha aaaa/mm/dd en dd/mm/aaaa
		/// </summary>
		/// <param name="fecha">La fecha que se desea revisar</param> 
		public static string convertDateDayMonthYear(string fecha)
		{
			validateDateAAAAMMDD(fecha);
			char[] delimiterChars = { '-' };
			string[] words = fecha.Split(delimiterChars);
			string fechaYearMonthDay = words[2] + "/" + words[1] + "/" + words[0];
			return fechaYearMonthDay;
		}
		
		/// <summary>
		/// Valida si un string es numerico.
		/// </summary>
		/// <param name="valor">El string a revisar.</param>
		/// <param name="nombre">El nombre del string a revisar.</param>
		public static void isNumeric(string valor, string nombre)
		{
			const string expresionRegular = @"^(\d+)\.(\d+)$|^(\d+)$";

			if (!validateFieldExpresion(valor, expresionRegular))
			{
				throw new ArgumentException("El dato " + nombre + " debe ser numérico.");
			}
		}

		/// <summary>
		/// Revisa que un string sea alfabético, que no incluya números ni simbolos.
		/// </summary>
		/// <param name="campo"></param>
		/// <param name="nombreCampo"></param>
		public static void isNotNumeric(string campo, string nombreCampo)
		{
			const string expresionRegular = @"^(([a-zA-Z])||ñ|Ñ|[{á-úÁ-Ú]|ü|\s)+$";

			if (!validateFieldExpresion(campo, expresionRegular))
			{
				throw new ArgumentException("El " + nombreCampo + " no puede contener numeros");
			}
		}

		public static string ToNumeric(string strIn)
		{
			const string expresionRegular = @"^(([a-zA-Z])||ñ|Ñ|[{á-úÁ-Ú]|ü|\s)+$";
			return Regex.Replace(strIn, expresionRegular, "");
		}

		/// <summary>
		/// Valida que un string tenga al menos sierto tamaño.
		/// </summary>
		/// <param name="campo"></param>
		/// <param name="tamaño"></param>
		/// <param name="nombreCampo"></param>
		public static void valitaeMinSize(string campo, string tamaño, string nombreCampo)
		{
			string expresionRegular = @".{" + tamaño + ",}";

			if (!validateFieldExpresion(campo, expresionRegular))
			{
				throw new ArgumentException("El " + nombreCampo + " debe contener al menos " + tamaño + " caracteres (Sin contar los espacios)");
			}
		}

		/// <summary>
		/// Valida que un e-mail tenga formato correcto(user@domain.com)
		/// </summary>
		/// <param name="eMail">El e-mail a revisar</param>
		public static void validateEmail(string eMail)
		{
			const string expresionRegular = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})$";

			if (!validateFieldExpresion(eMail, expresionRegular))
			{
				throw new ArgumentException("El Formato del E-mail es incorrecto.");
			}
		}

		public static void validatePartList(string sPartList, string menssage)
		{
			const string expresionRegular = @"^(\w+((/\w+)?(-\w+)?(\b\w+)?(_\w+)?(.\w+)?)*)+$";

			if (!validateFieldExpresion(sPartList, expresionRegular))
			{
				throw new ArgumentException("El Formato de " + menssage + " es incorrecto.");
			}
		}

		/// <summary>
		/// Valida un nuevo Password el nuevo password solo deve contener Letras y numeros
		/// y deve contener al menos una letra minuscula, una mauscula  y al menos un numero.
		/// </summary>
		/// <param name="eMail">El Password a Validar</param>
		public static void validatePassword(string password)
		{
			const string expresionRegular = @"^\w*(([a-z]+[A-Z]+\d+)|([a-z]+\d+[A-Z]+)|([A-Z]+[a-z]+\d+)|([A-Z]+\d+[a-z]+)|(\d+[A-Z]+[a-z]+)|(\d+[a-z]+[A-Z]+))\w*$";

			if (!validateFieldExpresion(password, expresionRegular))
			{
				throw new ArgumentException("El Password solo puede contener letras y números, y debe incluir al menos una letra mayúscula, una letra minúscula y un dígito.");
			}
			//if (Util.validarSrringCarateresSeguidosRepetidos(password))
			//{
			//    throw new ArgumentException("El Nuevo Password no puede contener dos caracteres identicos consecutivos.");
			//}
		}

		/// <summary>
		/// Obtiene la fecha actual con el formato aaaa-mm-dd
		/// </summary>
		/// <returns>la fecha actual con el formato aaaa-mm-dd</returns>
		public static string getDate()
		{
			DateTime today = DateTime.Today;
			string fechaProceso = today.Year + "-" + today.Month + "-" + today.Day;
			return fechaProceso;
		}

		/// <summary>
		/// Hace la comparacion entre dos fechas, la fecha inicial no debe ser porsterior a la fecha final.
		/// </summary>
		/// <param name="fechaInicio">La fecha inicial</param>
		/// <param name="fechaFin">La fecha final</param>
		public static void validateDatesHomeEnd(string fechaInicio, string title)
		{
			if ((DateTime.Parse(fechaInicio)) <= (DateTime.Now))
			{
				throw new ArgumentException("La Fecha " + title + " debe ser posterior a la Fecha Actual.");
			}
		}


	}
}