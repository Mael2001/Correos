using System;
using System.Net.Mail;
using System.Data.Odbc;
using System.Data;

namespace Correos
{
    class Program
    {
        static OdbcConnection odbcCon = new OdbcConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/BotVinculacion/db/Bot_Base.accdb");
        static OdbcConnection odbcC = new OdbcConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/BotVinculacion/db/Vinculacion_Base.accdb");
        static void Main(string[] args)
        {
            Console.WriteLine("Ingrese el destinatario del correo: ");
            //string correo = Console.ReadLine();
            EnviarCorreo();
        }

        public static void EnviarCorreo()
        {
            var fromAddress = new MailAddress("sanchezabj07@unitec.edu", "Bradley Sánchez");
            //var toAddress = new MailAddress(destinatario, "");
            const string fromPassword = "Rekjel07";
            string subject = "Detalle de Horas de Vinculación";
            string body = "<h2>Este es tu detalle de horas:<h2>\n";

            var cliente = new SmtpClient()
            {
                Host = "smtp.office365.com",
                Port = 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = true,
                EnableSsl = true,
                Credentials = new System.Net.NetworkCredential(fromAddress.Address, fromPassword)

            };


            string selectquery = "SELECT NumeroCuenta, CorreoElectronico FROM Alumnos;";
            

            OdbcCommand cmd = new OdbcCommand(selectquery, odbcCon);
            var datatable = GetDataTable(cmd);
            foreach (DataRow dr in datatable.Rows)
            {
                string correo = dr["CorreoElectronico"].ToString();
                try
                {
                    var toAddress = new MailAddress(correo);
                    Console.WriteLine(dr["NumeroCuenta"].ToString());
                    string dethoras = HorasDetalle2(dr["NumeroCuenta"].ToString());
                    string cuerpo = body + "<h1>" + dethoras + "<h1>";
                    var message = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = subject,
                        Body = cuerpo,
                        IsBodyHtml = true
                    };
                    cliente.Send(message);
                    Console.WriteLine("Se envio el correo a "+correo);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error al enviar el correo a "+ correo+ " " + e.Message);
                }
            }
            
        }
        private static DataTable GetDataTable(OdbcCommand cmd)
        {
            OdbcDataAdapter da1 = new OdbcDataAdapter(cmd);
            var datatable = new DataTable();
            da1.Fill(datatable);
            return datatable;
        }
        public static string HorasDetalle2(string nCuenta)
        {
            string detalles = "";
            string selectQuery = "SELECT id_proyecto,Periodo,Beneficiario,Horas_Acum FROM [Tabla General] where No_Cuenta = ? ";
            try
            {
                var cmd = new OdbcCommand(selectQuery, odbcC);
                cmd.Parameters.Add("@Cuenta", OdbcType.VarChar).Value = nCuenta;
                var datatable = GetDataTable(cmd);
                foreach (DataRow dr in datatable.Rows)
                {
                    detalles += "Nombre de Proyecto: " + dr["id_proyecto"].ToString() + "\n";
                    detalles += "Periodo: " + dr["Periodo"].ToString() + "\n";
                    detalles += "Beneficiaro: " + dr["Beneficiario"].ToString() + "\n";                   //HORAS DE PROYECTO 
                    detalles += "Horas Trabajadas:" + dr["Horas_Acum"].ToString() + "\n\n";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine();
            }
            Console.WriteLine(detalles);
            return detalles;
        }
    }
}