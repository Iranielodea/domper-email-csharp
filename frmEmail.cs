using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DomperEmail
{
    public partial class frmEmail : Form
    {
        ServicoEmail _servico;
        ServicoEmailOutros _servicoOutros;
        ServicoEmailGmail _servicoEmailGmail;

        IServicoEmail _servicoEmail;

        List<Email> _emails;
        public frmEmail()
        {
            InitializeComponent();
            _servico = new ServicoEmail();
            _servicoOutros = new ServicoEmailOutros();
            _servicoEmailGmail = new ServicoEmailGmail();
            
            _emails = new List<Email>();
        }

        private void EnviarEmail()
        {
            string arquivo = "email.txt";
            if (File.Exists(arquivo))
            {
                Email email = null;
                
                int opcao;
                bool opcao11 = false;
                string linha;
                foreach (string line in File.ReadLines("Email.txt", Encoding.GetEncoding("iso-8859-1")))
                {
                    opcao = 0;
                    if (! string.IsNullOrEmpty(line))
                    {
                        try
                        {
                            opcao = int.Parse(line.Substring(0, 3));
                        }
                        catch
                        {
                            if (opcao11)
                            {
                                email.Texto += linha = line.Substring(0);
                                continue;
                            }
                        }

                        linha = line.Substring(4);
                        switch (opcao)
                        {
                            case 1:
                                email = new Email();
                                email.Porta = int.Parse(linha);
                                break;
                            case 2:
                                email.Host = linha;
                                break;
                            case 3:
                                email.UserName = linha;
                                break;
                            case 4:
                                email.Password = linha;
                                break;
                            case 5:
                                email.MeuEmail = linha;
                                break;
                            case 6:
                                email.MeuNome = linha;
                                break;
                            case 7:
                                email.Destinatario = linha;
                                break;
                            case 8:
                                email.Copia = linha;
                                break;
                            case 9:
                                email.CopiaOculta = linha;
                                break;
                            case 10:
                                email.Assunto = linha;
                                break;
                            case 11:
                                email.Texto = linha;
                                opcao11 = true;
                                break;
                            case 12:
                                email.ArquivoTexto = linha;
                                opcao11 = false;
                                break;
                        }

                        if (opcao == 12)
                        {
                            _emails.Add(email);
                        }
                    }
                }

                pgbBarra.Maximum = _emails.Count;
                int contaEmail = 1;

                foreach (var item in _emails)
                {
                    try
                    {
                        if (item.Host.Trim() == "smtp.office365.com")
                            _servicoEmail = new ServicoEmail();
                        else if (item.Host.Trim() == "smtp.gmail.com")
                            _servicoEmail = new ServicoEmailGmail();
                        else
                            _servicoEmail = new ServicoEmailOutros();

                        _servicoEmail.EnviarEmail(item);

                        pgbBarra.Value = contaEmail;
                        contaEmail++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

            /*
        Writeln(arq, '001 ' + email.Porta.ToString());
        Writeln(arq, '002 ' + email.Host);
        Writeln(arq, '003 ' + email.UserName);
        Writeln(arq, '004 ' + email.Password);
        Writeln(arq, '005 ' + email.MeuEmail);
        Writeln(arq, '006 ' + email.MeuNome);
        Writeln(arq, '007 ' + email.Destinatario);
        Writeln(arq, '008 ' + email.Copia);
        Writeln(arq, '009 ' + email.CopiaOculta);
        Writeln(arq, '010 ' + email.Assunto);
        Writeln(arq, '011 ' + email.Texto);
        Writeln(arq, '012 ' + email.ArquivoAnexo);
             */

        }

        private void frmEmail_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            EnviarEmail();
            Application.Exit();
        }
    }
}
