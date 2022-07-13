using System;
using System.Net.Mail;
using System.Text;

namespace DomperEmail
{
    public class ServicoEmailOutros : IServicoEmail
    {
        public void EnviarEmail(Email email)
        {
            // https://www.youtube.com/watch?v=CboOxdiK7bs - fonte de materia

            MailMessage mail = new MailMessage();

            mail.From = new MailAddress(email.MeuEmail);
            mail.To.Add(email.Destinatario);
            mail.Subject = email.Assunto;
            StringBuilder sb = new StringBuilder();
            mail.IsBodyHtml = true;
            sb.AppendLine(email.Texto);
            mail.Body = sb.ToString();

            if (!string.IsNullOrEmpty(email.Copia))
            {
                var copias = email.Copia.Split(';');
                foreach (var copia in copias)
                {
                    mail.CC.Add(copia);
                }
            }

            if (!string.IsNullOrEmpty(email.CopiaOculta))
            {
                var copiasOcultas = email.CopiaOculta.Split(';');
                foreach (var copiaOculta in copiasOcultas)
                {
                    mail.Bcc.Add(copiaOculta);
                }
            }

            var destinatarios = email.Destinatario.Split(';');
            foreach (var destinatario in destinatarios)
            {
                mail.To.Add(destinatario);
            }

            if (!string.IsNullOrEmpty(email.ArquivoTexto))
            {
                Attachment attach = new Attachment(email.ArquivoTexto);
                mail.Attachments.Add(attach);
            }

            using (var smtp = new SmtpClient(email.Host))
            {
                smtp.UseDefaultCredentials = false; // vamos utilizar credencias especificas

                // seu usuário e senha para autenticação
                smtp.Credentials = new System.Net.NetworkCredential(email.UserName, email.Password);
                smtp.EnableSsl = false;
                smtp.Port = email.Porta;
                try
                {
                    smtp.Send(mail);
                } 
                catch(Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                
            }
        }
    }
}
