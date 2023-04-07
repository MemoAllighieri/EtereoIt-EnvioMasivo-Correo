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

            for (int i = 0; i < listaUtil.Count(); i++)
            {
                var emp = listaUtil[i].Split(';').ToArray();
                if (emp.Length < 3) return;
                Log.Information("_____________________________________________________________________");
                Console.WriteLine("_____________________________________________________________________");
                EnterpriseDataDto emails = new EnterpriseDataDto()
                {
                    Name = emp[0],
                    Email = emp[1],
                    Phone = emp[2],
                };
                Log.Information($"User :{emails.Name} | Email :{emails.Email} | Phone :{emails.Phone}");
                Console.WriteLine($"User :{emails.Name} | Email :{emails.Email} | Phone :{emails.Phone}");
                //Send Email                
                bool rpta = _sendEmail.Send(smtpClient, _configurationSmtpClient.User, emails.Email);
            }
        }
        catch(Exception ex)
        {
            Log.Error($"Error in method Program.EmailMassive : {ex.Message}");
            Console.WriteLine($"Error in method Program.EmailMassive : {ex.Message}");
            throw ex;
        }
        
    }
}