using System;
using System.Text;
using System.Web;
using System.IO;
using System.Diagnostics;

/// <summary>
/// This class represents an error raised when requesting an ASP.NET page
/// </summary>
public class FriendlyError
{
	private Exception _lastError;
	private StackFrame _frame;
	private string _fileName;

	/// <summary>
	/// Returns the human readable error message
	/// </summary>
	public string GetErrorMessage()
	{
		return HttpContext.Current.Server.HtmlEncode(_lastError.GetBaseException().Message);
	}

	/// <summary>
	/// Gets the file associated with the error
	/// </summary>
	public string GetFileName()
	{
		string appRoot = HttpContext.Current.Request.PhysicalApplicationPath;
		string path = _fileName.Remove(0, appRoot.Length);
		return path.Replace(@"\", @"/");
	}

	/// <summary>
	/// Returns the line and column associated with an error
	/// </summary>
	public string GetErrorCoordinates()
	{
		return String.Format("line {0}, column {1}", _frame.GetFileLineNumber(), _frame.GetFileColumnNumber());
	}

	/// <summary>
	/// Returns the source code associated with an error starting a specified number of lines before and after the error line
	/// </summary>
	public string GetSourceCode(int startOffset, int endOffset)
	{
		// Get line and column number
		int lineNumber = _frame.GetFileLineNumber();
		int colNumber = _frame.GetFileColumnNumber();

		// If no line number
		if (lineNumber == 0)
			return "No relevant source code";

		// Load source code file
		string[] sourceLines = File.ReadAllLines(_fileName);

		// Calculate start and end line
		int startLine = lineNumber + startOffset;
		int endLine = lineNumber + endOffset;
		if (startLine < 0)
			startLine = 0;
		if (endLine > sourceLines.Length)
			endLine = sourceLines.Length;

		// Format source code
		var builder = new StringBuilder();
		for (int k = startLine; k < endLine; k++)
		{
			string sourceLine = HttpContext.Current.Server.HtmlEncode(sourceLines[k]);
			builder.AppendFormat(k + 1 == lineNumber ? "<span class='error'>{0}: {1}</span>\n" : "{0}: {1}\n", k + 1, sourceLine);
		}
		return builder.ToString();
	}

	/// <summary>
	/// Returns the url for a picture from the Pictures folder
	/// </summary>
	public string GetRandomPicture()
	{
		HttpContext context = HttpContext.Current;
		var rnd = new Random();
		string[] pictures = Directory.GetFiles(context.Server.MapPath("~/Error/Images"));
		return "~/Error/Images/" + Path.GetFileName(pictures[rnd.Next(pictures.Length)]);
	}

	public FriendlyError()
	{
		// Get the last error message
		HttpContext context = HttpContext.Current;
		_lastError = context.Server.GetLastError();
		// Get the Stack Frame
		var stack = new StackTrace(_lastError.GetBaseException(), true);
		for (int i = 0; i < stack.FrameCount; i++)
			if (!String.IsNullOrEmpty(stack.GetFrame(i).GetFileName()))
				_frame = stack.GetFrame(i);
		if (_frame == null)
			_frame = stack.GetFrame(0);
		// Get file name
		_fileName = _frame.GetFileName();
		if (String.IsNullOrEmpty(_fileName))
			_fileName = context.Request.PhysicalPath;
	}
}