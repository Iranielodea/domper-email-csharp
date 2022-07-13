using System;
using System.Net.Mail;
using System.Text;

namespace DomperEmail
{
    public class ServicoEmail : IServicoEmail
    {
        public void EnviarEmail(Email email)
        {
            String userName = email.UserName;
            String password = email.Password;
            //MailMessage msg = new MailMessage(email.MeuEmail, email.Destinatario);
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(email.MeuEmail);
            msg.Subject = email.Assunto;
            StringBuilder sb = new StringBuilder();
            msg.IsBodyHtml = true;
            sb.AppendLine(email.Texto);
            msg.Body = sb.ToString();
            //msg.BodyEncoding = Encoding.UTF8;

            if (!string.IsNullOrEmpty(email.Copia))
            {
                var copias = email.Copia.Split(';');
                foreach (var copia in copias)
                {
                    msg.CC.Add(copia);
                }

                //MailAddress cc = new MailAddress(email.Copia);
                //msg.CC.Add(cc);
            }

            if (!string.IsNullOrEmpty(email.CopiaOculta))
            {
                var copiasOcultas = email.CopiaOculta.Split(';');
                foreach (var copiaOculta in copiasOcultas)
                {
                    msg.Bcc.Add(copiaOculta);
                }

                //MailAddress bcc = new MailAddress(email.CopiaOculta);
                //msg.Bcc.Add(bcc);
            }

            var destinatarios = email.Destinatario.Split(';');
            foreach (var destinatario in destinatarios)
            {
                msg.To.Add(destinatario);
            }

            if (!string.IsNullOrEmpty(email.ArquivoTexto))
            {
                Attachment attach = new Attachment(email.ArquivoTexto);
                msg.Attachments.Add(attach);
            }

            SmtpClient SmtpClient = new SmtpClient();
            SmtpClient.Credentials = new System.Net.NetworkCredential(userName, password);
            SmtpClient.Host = email.Host; // "smtp.office365.com";
            SmtpClient.Port = email.Porta; // 587;
            SmtpClient.EnableSsl = true;
            SmtpClient.Send(msg);

            //String userName = "irani@Domper.com.br";
            //String password = "E@584369";
            //MailMessage msg = new MailMessage("irani@Domper.com.br", "iranielodea@hotmail.com");
            //msg.Subject = "Assunto";
            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine("Name: ");
            //sb.AppendLine("Teste envio de email ");
            //sb.AppendLine("Email: teste");
            //msg.Body = sb.ToString();
            ////Attachment attach = new Attachment(Server.MapPath("folder/" + ImgName));
            ////msg.Attachments.Add(attach);
            //SmtpClient SmtpClient = new SmtpClient();
            //SmtpClient.Credentials = new System.Net.NetworkCredential(userName, password);
            //SmtpClient.Host = "smtp.office365.com";
            //SmtpClient.Port = 587;
            //SmtpClient.EnableSsl = true;
            //SmtpClient.Send(msg);
        }
    }
}
