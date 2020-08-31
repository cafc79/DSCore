using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.Security;
using Telerik.Web.UI;

/// <summary>
/// Summary description for Utilerias
/// </summary>
public class Utilerias
{
	public Utilerias() { }
	
	public FormsAuthenticationTicket ValoresAcceso()
	{
		HttpContext context = HttpContext.Current;
		HttpCookie cookieBase = context.Request.Cookies[FormsAuthentication.FormsCookieName];
		if (cookieBase != null)
		{
			FormsAuthenticationTicket ticketDeSesion = FormsAuthentication.Decrypt(cookieBase.Value);
			return ticketDeSesion;
		}
		else return null;
	}

	public int GetUserId()
	{
		HttpContext contexto = HttpContext.Current;
		return 0;
		//return BOCAR.SACI.Business.SecurityManager.ObtenerUsuario(contexto).IdUsuario; // int.Parse(Session["IdUsr"].ToString());
	}

	public int GetEmpleadoId()
	{
		HttpContext contexto = HttpContext.Current;
		return 0;
		//return BOCAR.SACI.Business.SecurityManager.ObtenerUsuario(contexto).IdEmpleado; // int.Parse(Session["IdUsr"].ToString());
	}

	public void LoadMenu(string sMenu, int iRol, ref Menu menuVertical)
	{
		/*var swap = new BOCAR.SACI.Business.MenuRolManager();
		var menu = (List<BOCAR.SACI.Data.menuRolCompositeType>)HttpContext.Current.Session["MenuH"];
		var iMenu = (from p in menu where p.sModulo.Equals(sMenu) select p).FirstOrDefault();
		if (iMenu == null)
		{
			throw new ArgumentNullException();
		}
		var menuInfor = swap.GetMenuPagina(iMenu.iIdMenu, iRol);
		foreach (BOCAR.SACI.Data.menuCompositeType mt in menuInfor)
		{
			var miNuevo = new MenuItem { Text = mt.sDescription, Value = mt.sURL + "|" + mt.sPermission, Selectable = true };
			menuVertical.Items.Add(miNuevo);
		}
		menuVertical.Items[0].Selected = true;*/
	}

	public void GetValidStage(string sCode, Menu menuVertical)
	{
		char[] splitC = { '|' };
		bool exito = false;
		for (int i = 0; i < menuVertical.Items.Count; i++)
			if (sCode.Split(splitC)[0] == menuVertical.Items[i].Value.Split(splitC)[0])
			{
				exito = true;
				break;
			}
		if (!exito)
			throw new ArgumentNullException();
	}

	public string GetActionMenu(string sCode, bool bPermision)
	{
		char[] splitC = {'|'};
		return bPermision ? sCode.Split(splitC)[1] : sCode.Split(splitC)[0];
	}

	public void GetActionMenu(string sCode)
	{
		char[] splitC = { '|' };
		if (sCode.Split(splitC)[1] == "100")
			throw new AccessViolationException();
	}

	public void ErroDisplay(byte bError, string sMsg, ref Label lblMessage)
	{
		lblMessage.Visible = true;
		lblMessage.ForeColor = Color.Red;
		switch (bError)
		{
			case 1:
				lblMessage.Text = "Error al guardar el registro, comuníquese con el administrador: " + sMsg;
				break;
			case 2:
				lblMessage.Text = "La descripción o número ya existe, no es posible duplicar registros.";
				break;
			case 3:
				lblMessage.Text = sMsg;
				break;
			case 4:
				lblMessage.Text = "Error al eliminar el registro, comuniquese con el administrador";
				break;
			case 5:
				lblMessage.ForeColor = Color.Gray;
				lblMessage.Text = "Registro introducido con éxito";
				break;
			case 6:
				lblMessage.ForeColor = Color.Gray;
				lblMessage.Text = "Registro modificado con éxito";
				break;
			case 7:
				lblMessage.ForeColor = Color.Gray;
				lblMessage.Text = "Registro borrado con éxito";
				break;
			case 8:
				lblMessage.ForeColor = Color.DarkViolet;
				lblMessage.Text = "No se permite campos vacíos";
				break;
		}
	}

	public int IndexCombo(RadComboBox combo, string compare)
	{
		int i = 0;
		for (i = 0; i <= combo.Items.Count - 1; i++)
			if (combo.Items[i].Value == compare)
				break;
		return i;
	}

	public void AccessDocument(string path)
	{
		try
		{
			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);
		}
		catch(IOException ioe)
		{
			throw new IOException(ioe.Message);
		}
	}

	public bool DownloadFile(string iExchange, byte iSegmento, string sArchivo, ref string sDownloadPath)
	{
		try
		{
			byte iPath = 0;
			string[] documento = sArchivo.Split('\\');
			switch (iSegmento)
			{
				case 1:
					if (File.Exists(ConfigurationManager.AppSettings["URLArchivos"] + "\\" + iExchange + "\\" + documento[documento.Count() - 1]))
						iPath = 1;
					break;
				case 2:
					if (File.Exists(ConfigurationManager.AppSettings["rutaSeguimiento"] + "\\" + iExchange + "\\" + documento[documento.Count() - 1]))
						iPath = 2;
					break;
				case 3:
					if (File.Exists(ConfigurationManager.AppSettings["rutaLiberacion"] + "\\" + iExchange + "\\" + documento[documento.Count() - 1]))
						iPath = 3;
					break;
			}
			if (iPath > 0) {
				sDownloadPath = "Services/Downloader.ashx?path="+ iPath.ToString() + "&ex=" +iExchange + "&doc=" + documento[documento.Count() - 1];
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			return false;
		}
	}

	public bool DownloadRepOp(int iExchange, string sArchivo, ref string sDownloadPath)
	{
		try
		{
			byte iPath = 0;
			string[] documento = sArchivo.Split('\\');
			if (File.Exists(ConfigurationManager.AppSettings["URLArchivos"] + "\\" + iExchange + "\\" + documento[documento.Count() - 1]))
				iPath = 1;
			else if (File.Exists(ConfigurationManager.AppSettings["rutaSeguimiento"] + "\\" + iExchange + "\\" + documento[documento.Count() - 1]))
				iPath = 2;
			else if (File.Exists(ConfigurationManager.AppSettings["rutaLiberacion"] + "\\" + iExchange + "\\" + documento[documento.Count() - 1]))
				iPath = 3;
			if (iPath > 0)
			{
				sDownloadPath = "Services/Downloader.ashx?path=" + iPath.ToString() + "&ex=" + iExchange + "&doc=" + documento[documento.Count() - 1];
				return true;
			}
			if (File.Exists(sArchivo))
			{
				sDownloadPath = "Services/Downloader.ashx?path=4&ex=" + iExchange + "&doc=" + sArchivo;
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			return false;
		}
	}

	public string DefaultCombo(RadComboBox combo, string fieldCombo)
	{
		if (combo.Items.Count == 0)
			Util.isRequired(string.Empty, fieldCombo);
		else
		{
			if (String.IsNullOrEmpty(combo.SelectedValue))
				return combo.Items[0].Value;
			else
				return combo.SelectedValue;
		}
		return string.Empty;
	}

	public string Default_SimpleCombo(RadComboBox combo, string fieldCombo)
	{
		if (combo.Items.Count == 0)
			return string.Empty;
		else
		{
			if (String.IsNullOrEmpty(combo.SelectedValue))
				return combo.Items[0].Value;
			else
				return combo.SelectedValue;
		}
		return string.Empty;
	}

	public void CleanControls(Panel contenedor)
	{
		for (int i = 0; i <= contenedor.Controls.Count - 1; i++)
		{
			if (contenedor.Controls[i].GetType().Name == "TextBox")
				((TextBox)contenedor.Controls[i]).Text = string.Empty;
			if (contenedor.Controls[i].GetType().Name == "RadTextBox")
				((RadTextBox)contenedor.Controls[i]).Text = string.Empty;
			if (contenedor.Controls[i].GetType().Name == "RadMaskedTextBox")
				((RadMaskedTextBox)contenedor.Controls[i]).Text = string.Empty;
			if (contenedor.Controls[i].GetType().Name == "RadComboBox")
			{
				string ds = ((RadComboBox)contenedor.Controls[i]).DataSourceID;
				((RadComboBox)contenedor.Controls[i]).DataSourceID = null;
				((RadComboBox)contenedor.Controls[i]).DataSource = null;
				((RadComboBox)contenedor.Controls[i]).DataSourceID = ds;
				((RadComboBox)contenedor.Controls[i]).DataBind();
			}
		}
	}

	public void CleanControls(RadAjaxPanel contenedor)
	{
		for (int i = 0; i <= contenedor.Controls.Count - 1; i++)
		{
			if (contenedor.Controls[i].GetType().Name == "TextBox")
				((TextBox)contenedor.Controls[i]).Text = string.Empty;
			if (contenedor.Controls[i].GetType().Name == "RadTextBox")
				((RadTextBox)contenedor.Controls[i]).Text = string.Empty;
			if (contenedor.Controls[i].GetType().Name == "RadMaskedTextBox")
				((RadMaskedTextBox)contenedor.Controls[i]).Text = string.Empty;
			if (contenedor.Controls[i].GetType().Name == "RadComboBox")
			{
				((RadComboBox)contenedor.Controls[i]).SelectedValue = null;
				((RadComboBox)contenedor.Controls[i]).DataSourceID = null;
				((RadComboBox)contenedor.Controls[i]).DataSource = null;
				((RadComboBox)contenedor.Controls[i]).DataBind();
				//((RadComboBox)contenedor.Controls[i]).DataBind();
			}
		}
	}
}