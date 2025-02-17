using System;
using System.Configuration;
using System.Net.Mail;

namespace Data2Sheet
{
    /// <summary>
    /// Email class for the application.
    /// To run this set config flag ActivateEmail = yes
    /// Note set smtp up in config
    /// </summary>
    internal class Email
    {
        private readonly string _sendTo;
        private readonly string _message;
        private readonly string _subject;
        private readonly string _sendCC;

        public Email(string sendTo, string message, string subject, string sendCC) 
        {
            _sendTo = sendTo;
            _message = message;
            _subject = subject;
            _sendCC = sendCC;
        }

        public void SendMail()
        {
            var appSettings = ConfigurationManager.AppSettings;

            if (appSettings["ActivateEmail"] == "yes")
            {
                var smtpClient = new SmtpClient();
                var mailMessage = new MailMessage();
                foreach (var emailAddress in this._sendTo.Split(','))
                {
                    mailMessage.To.Add(emailAddress);
                }

                if (!String.IsNullOrWhiteSpace(this._sendCC))
                {
                    foreach (var emailAddress in this._sendCC.Split(','))
                    {
                        mailMessage.CC.Add(emailAddress);
                    }
                }

                //Send Mail
                mailMessage.Subject = this._subject;
                mailMessage.Body = this._message;

                smtpClient.Send(mailMessage);
                smtpClient.Dispose();
            }
        }
    }
}
