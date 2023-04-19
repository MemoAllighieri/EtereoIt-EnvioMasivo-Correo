using envioMasivoCorreos.Dtos;
using envioMasivoCorreos.Models;
using envioMasivoCorreos.Utils;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using Serilog;
using System.Data;
using System.Net;
using System.Net.Mail;

internal class Program
{
    static SendEmail _sendEmail = new SendEmail();
    static Util _util = new Util();
    static ConfigurationSmtpClient _configurationSmtpClient;

    private static void Main(string[] args)
    {
        //Config Logger
        Log.Logger = new LoggerConfiguration()
                        .WriteTo.File(@$"C:\Logs\{DateTime.Now:yyyy-MM-dd}-log.txt", rollingInterval: RollingInterval.Day)
                        .CreateLogger();

        Log.Information("Initialize serilog");

        //Create a builder and config
        Log.Information("Create a builder and config");
        var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
        var config = builder.Build();        

        //initialize smtp credentials
        Log.Information("initialize smtp credentials");        
        _configurationSmtpClient = new ConfigurationSmtpClient()
        {
            Host = config["ConfigurationSmtpClient:Host"],
            Port = int.Parse(config["ConfigurationSmtpClient:Port"]),
            EnableSsl = bool.Parse(config["ConfigurationSmtpClient:EnableSsl"]),
            User = config["ConfigurationSmtpClient:User"],
            Password = config["ConfigurationSmtpClient:Password"]
        };

        //Execute EmailMassive
        Log.Information("Execute EmailMassive");
        Console.WriteLine("Se ejecutara el metodo EmailMassive");
        EmailMassive();
    }
    private static void EmailMassive()
    {
        try
        {
            //initialize SmtpClient
            Log.Information("Initialize SmtpClient");
            SmtpClient smtpClient = _sendEmail.InicializaConfiguracion(_configurationSmtpClient);

            var listaUtil = _util.GetData().ToList();

            for (int i = 1; i < listaUtil.Count(); i++)
            {
                var emp = listaUtil[i].Split('|').ToArray();
                if (emp.Length < 3) return;
                Log.Information("_____________________________________________________________________");
                Log.Information($"EMAIL NRO:{i} | fecha: {DateTime.Now}");
                Console.WriteLine("_____________________________________________________________________");
                Console.WriteLine($"EMAIL NRO:{i} | fecha: {DateTime.Now}");
                EnterpriseDataDto data = new EnterpriseDataDto()
                {
                    RUC = emp[0],
                    Email = emp[1],
                    EmailADM = emp[2] == "" ? emp[1] : emp[2],
                    Phone = emp[3],
                    Departament = emp[4],
                    Province = emp[5],
                    District = emp[6],
                    BusinessName = emp[7],
                    Heading = emp[8],
                    WebPage = emp[9],
                };
                Log.Information($"User :{data.RUC} | Email :{data.Email} | Phone :{data.BusinessName}");
                Console.WriteLine($"User :{data.RUC} | Email :{data.Email} | Phone :{data.BusinessName}");
                //Send Email

                MailMessage email = PrepareMail(data);

                bool rpta = _sendEmail.Send(smtpClient, email);
            }
        }
        catch(Exception ex)
        {
            Log.Error($"Error in method Program.EmailMassive : {ex.Message}");
            Console.WriteLine($"Error in method Program.EmailMassive : {ex.Message}");
            throw ex;
        }
        
    }

    private static MailMessage PrepareMail(EnterpriseDataDto data)
    {
        var email = new MailMessage(_configurationSmtpClient.User.ToString(), "test-r54h6t4om@srv1.mail-tester.com");
        //var email = new MailMessage(_configurationSmtpClient.User.ToString(), data.EmailADM);

        email.Subject = "Asunto: " + "Prueba de envío de correo";

        string body = File.ReadAllText(@$"C:\etereo-correo\index.html");

        Attachment oAttachmentBanner = new Attachment(@$"C:\etereo-correo\images\BANNER.jpg");
        oAttachmentBanner.ContentId = "bannerid";
        Attachment oAttachmentFacebook = new Attachment(@$"C:\etereo-correo\images\facebook-icon.png");
        oAttachmentFacebook.ContentId = "facebookid";
        Attachment oAttachmentInstagram = new Attachment(@$"C:\etereo-correo\images\instagram-icon.png");
        oAttachmentInstagram.ContentId = "instagramid";
        Attachment oAttachmentLinkd = new Attachment(@$"C:\etereo-correo\images\linkedin-icon.png");
        oAttachmentLinkd.ContentId = "linkdid";
        Attachment oAttachmentLogo = new Attachment(@$"C:\etereo-correo\images\logo-etereo-nav.png");
        oAttachmentLogo.ContentId = "logoid";
        Attachment oAttachmentTwitter = new Attachment(@$"C:\etereo-correo\images\twitter-icon.png");
        oAttachmentTwitter.ContentId = "twitterid";
        Attachment oAttachmentIconoWeb = new Attachment(@$"C:\etereo-correo\images\etereo.jpg");
        oAttachmentIconoWeb.ContentId = "iconowebid";
        Attachment oAttachmentIconoWsp = new Attachment(@$"C:\etereo-correo\images\contacto.jpg");
        oAttachmentIconoWsp.ContentId = "iconowspid";

        body = body.Replace("XXX_ETEREO_LOGO_XXX", oAttachmentLogo.ContentId)
                   .Replace("XXX_ETEREO_BANNER_XXX", oAttachmentBanner.ContentId)
                   .Replace("XXX_FACEBOOK_XXX", oAttachmentFacebook.ContentId)
                   .Replace("XXX_INSTAGRAM_XXX", oAttachmentInstagram.ContentId)
                   .Replace("XXX_LINKD_XXX", oAttachmentLinkd.ContentId)
                   .Replace("XXX_TWITTER_XXX", oAttachmentTwitter.ContentId)
                   .Replace("XXX_ICONO_WEB_XXX", oAttachmentIconoWeb.ContentId)
                   .Replace("XXX_ICONO_WSP_XXX", oAttachmentIconoWsp.ContentId)
                   .Replace("XXX_EMPRESA_NOMBRE_XXX", data.BusinessName);

        email.Body = body;
        email.IsBodyHtml = true;
        email.Attachments.Add(oAttachmentLogo);
        email.Attachments.Add(oAttachmentBanner);
        email.Attachments.Add(oAttachmentFacebook);
        email.Attachments.Add(oAttachmentInstagram);
        email.Attachments.Add(oAttachmentLinkd);
        email.Attachments.Add(oAttachmentTwitter);

        return email;
    }
}