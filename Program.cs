using System;
using System.Net.Mail;
using System.Data.Odbc;

namespace Correos
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Ingrese el destinatario del correo: ");
            //string correo = Console.ReadLine();
            EnviarCorreo();
        }

        public static bool EnviarCorreo()
        {
            var fromAddress = new MailAddress("brandsan15@gmail.com", "Bradley Sánchez");
            //var toAddress = new MailAddress(destinatario, "");
            const string fromPassword = "bradleysanchez";
            string subject = "Correo de verificación";
            string body = "<h2>Codigo de Verficacion Bot Unitec:<h2>\n <h1>000000<h1>";

            var cliente = new SmtpClient()
            {
                Host = "smtp.gmail.com",
                Port = 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = true,
                EnableSsl = true,
                Credentials = new System.Net.NetworkCredential(fromAddress.Address, fromPassword)

            };


            string selectquery = "SELECT CorreoElectronico FROM Alumnos;";
            OdbcConnection odbcCon = new OdbcConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/BRADLEYJOHELSANCHEZA/Downloads/Bot_Base.accdb");

            OdbcCommand cmd = new OdbcCommand(selectquery, odbcCon);
            odbcCon.Open();
            OdbcDataReader dataReader = cmd.ExecuteReader();
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader.GetString(0));
                try
                {
                    var toAddress = new MailAddress(dataReader.GetString(0));
                    Console.WriteLine(dataReader.GetString(0));
                    var message = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = subject,
                        Body = body,
                        IsBodyHtml = true
                    };
                    cliente.Send(message);
                    Console.WriteLine("Se envio el correo :)");
                    return true;
                }
                catch (Exception e)
                {
                    odbcCon.Close();
                    Console.WriteLine("Error al enviar el correo :(" + e.Message);
                    return false;
                }
            }
            odbcCon.Close();
            return false;
        }
    }
}