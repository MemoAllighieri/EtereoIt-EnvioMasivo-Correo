﻿using envioMasivoCorreos.Models;
using Serilog;
using System.Net;
using System.Net.Mail;

namespace envioMasivoCorreos.Utils
{
    public class SendEmail
    {
        public SmtpClient InicializaConfiguracion(ConfigurationSmtpClient configurationSmtpClient)
        {
            return new SmtpClient(configurationSmtpClient.Host, configurationSmtpClient.Port) 
            {
                EnableSsl = configurationSmtpClient.EnableSsl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(configurationSmtpClient.User, configurationSmtpClient.Password)
            };
        }

        public bool Send(SmtpClient client, string from, string to)
        {
            Log.Information("Sending Email...");
            Console.WriteLine("Sending Email...");
            var response = false;
            try
            {                
                var email = new MailMessage(from, to);
                email.Subject = "Asunto: " + "Prueba de envío de correo";
                email.Body = "PRUEBA PRUEBA PRUEBA PRUEBA";
                client.Send(email);

                response = true;
                Log.Information("Email Send!!!");
                Console.WriteLine("Email Send!!!");
            }
            catch(Exception ex)
            {
                Log.Error($"Error in method SendEmail.Send : {ex.Message}");
                Console.WriteLine($"Error in method SendEmail.Send : {ex.Message}");
                throw ex;
            }
            return response;    
        }
    }
}
