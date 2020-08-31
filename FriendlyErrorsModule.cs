using System;
using System.Web;
using System.Web.Configuration;

/// <summary>
/// Handles all unhanlded exceptions in an application
/// </summary>
public class FriendlyErrorsModule : IHttpModule
{
	public void Init(HttpApplication app)
	{
		app.Error += new EventHandler(app_Error);
	}

	/// <summary>
	/// When debug mode is enabled, display
	/// friendly error page
	/// </summary>
	private void app_Error(object sender, EventArgs e)
	{
		HttpApplication app = (HttpApplication) sender;
		HttpContext context = app.Context;
		Exception error = context.Server.GetLastError().GetBaseException();
		context.Response.Clear();
		/*CompilationSection compilationConfig = (CompilationSection) WebConfigurationManager.GetWebApplicationSection("system.web/compilation");
		if (compilationConfig.Debug)
			context.Server.Transfer("~/FriendlyErrors/Debug.aspx");
		else
			context.Server.Transfer("~/FriendlyErrors/Public.aspx");*/
	}

	public void Dispose()
	{
	}
}