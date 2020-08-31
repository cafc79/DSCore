using System.Web.UI;

/// <summary>
/// Summary description for Alert
/// </summary>
public static class Alert
{
	/// <summary> 
	/// Shows a client-side JavaScript alert in the browser. 
	/// </summary> 
	/// <param name="message">The message to appear in the alert.</param> 
	public static void Show(string message, Page pagina)
	{
		string cleanMessage = message.Replace("'", "\\'");
		var sb = new System.Text.StringBuilder();
		sb.Append(@"<script language='javascript'>");
		sb.Append(@"alert('" + cleanMessage + "');");
		sb.Append(@"</script>");
		ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), "alert", sb.ToString(), false);
	}

	public static void Permisos(Page pagina)
	{
		var sb = new System.Text.StringBuilder();
		sb.Append(@"<script language='javascript'>");
		sb.Append(@"alert('No tienes permiso para acceder a la función solicitada');");
		sb.Append(@"</script>");
		ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), "alert", sb.ToString(), false);
	}

	public static void Denegado(Page pagina)
	{
		var sb = new System.Text.StringBuilder();
		sb.Append(@"<script language='javascript'>");
		sb.Append(@"alert('No tienes permiso para acceder a la página solicitada');");
		sb.Append(@"history.back(-1);");
		sb.Append(@"</script>");
		ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), "alert", sb.ToString(), false);
	}

	public static void Refresh(Page pagina)
	{
		var sb = new System.Text.StringBuilder();
		sb.Append(@"<script language='javascript'>");
		sb.Append(@"setup();");
		sb.Append(@"</script>");
		ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), "", sb.ToString(), false);
	}
}