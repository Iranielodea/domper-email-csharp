using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomperEmail
{
    public class Email
    {
        public int Porta { get; set; }
        public string Host { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string MeuEmail { get; set; }
        public string MeuNome { get; set; }
        public string Destinatario { get; set; }
        public string Copia { get; set; }
        public string CopiaOculta { get; set; }
        public string Assunto { get; set; }
        public string Texto { get; set; }
        public string ArquivoTexto { get; set; }
    }
}
