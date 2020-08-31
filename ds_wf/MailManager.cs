using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Runtime.Remoting.Messaging;

namespace DeltaCore.WorkFlow
{
	public class MailManager
	{
		private SmtpClient smtp;
		private MailMessage message;
		public delegate bool SendEmailDelegate(string msgBody, string msgSubject);

		/// <summary>
		/// Constructor alterno el cual recibe la configuración personalizada
		/// </summary>
		public MailManager(string server, int puerto, string usuario, string contraseña, string sendFrom, string sendTo)
		{
			smtp = new SmtpClient(server)
			       	{Credentials = new System.Net.NetworkCredential(usuario, contraseña), EnableSsl = false, Timeout = 10000};
			if (puerto != 0)
				smtp.Port = puerto;
			message = new MailMessage(sendFrom, sendTo) { IsBodyHtml = true };
		}

		public MailManager(string address)
		{
			smtp = new SmtpClient();
			message = new MailMessage();
			message.To.Add(address);
			message.IsBodyHtml = true;
		}
		
		public void Message_AddOn(List<string> lstBCC, Stream stream, string attName, string attMIME)
		{
			if (stream != null && stream.Length > 0 && !string.IsNullOrEmpty(attName))
				message.Attachments.Add(new Attachment(stream, attName, attMIME));
			foreach (var bcc in lstBCC)
				message.Bcc.Add(bcc);
		}

		/// <summary>
		/// Método que se encargá de enviar el mail, basado en la configuracion del contructor
		/// </summary>
		public bool SendMail(string msgBody, string msgSubject)
		{
			message.Body = msgBody;
			message.Subject = msgSubject;
			//ServicePointManager.ServerCertificateValidationCallback = (s, certificate, chain, sslPolicyErrors) => true;
			try
			{
				smtp.Send(message);
				return true;
			}
			catch (SmtpException se)
			{
				throw new Exception(se.StatusCode + " " + se.Message + "; SendMail_SMTP - No se puede enviar el mail", se);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message + "; SendMail_SMTP - No se puede enviar el mail", ex);
			}
		}

		public void GetResultsOnCallback(IAsyncResult ar)
		{
			var del = (SendEmailDelegate)((AsyncResult)ar).AsyncDelegate;
			try
			{
				bool result = del.EndInvoke(ar);
			}
			catch (Exception ex)
			{
				bool result = false;
			}
		}

		public bool SendEmailAsync(string msgBody, string msgSubject)
		{
			try
			{
				var dc = new SendEmailDelegate(SendMail);
				var cb = new AsyncCallback(this.GetResultsOnCallback);
				IAsyncResult ar = dc.BeginInvoke(msgBody, msgSubject, cb, null);
			}
			catch (Exception ex)
			{
				return false;
			}
			return true;
		}
	}
}