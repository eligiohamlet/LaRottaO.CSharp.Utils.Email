using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading.Tasks;

namespace LaRottaO.CSharp.Utils.Email
{
    public class SendEmail
    {
        public Task<Tuple<Boolean, String>> sendWithoutAttachments(String argServer, String argPort, String argDestinationAddress, 
            String argSubject, String argBody, Boolean argIsBodyHtml, String fromEmail, String fromPassword, String fromNickname, 
            String argDestinationCc = null, String argDestinationBcc = null, bool agregaFirma = false, String nombreFirma = "") //,bool addAnimation = false, String animationPath =""
        {
            return Task.Run(() =>
            {
                fromNickname = "<" + fromNickname + "> " + fromEmail;

                Console.WriteLine("Server: " + argServer);
                Console.WriteLine("Port: " + argPort);
                Console.WriteLine("DestinationAddress: " + argDestinationAddress);
                Console.WriteLine("CCAddress: " + argDestinationCc);
                Console.WriteLine("BCCAddress: " + argDestinationBcc);
                Console.WriteLine("Subject: " + argSubject);
                Console.WriteLine("Body: ..." + argBody);
                Console.WriteLine("IsBodyHtml: " + argIsBodyHtml);
                Console.WriteLine("FromEmail: " + fromEmail);
                Console.WriteLine("Password: " + fromPassword);
                Console.WriteLine("FromNickname: " + fromNickname);

                try
                {
                    var client = new SmtpClient(argServer, Convert.ToInt32(argPort))
                    {
                        Credentials = new NetworkCredential(fromEmail, fromPassword),
                        EnableSsl = true,
                    };

                    MailMessage msg = new MailMessage(fromNickname, argDestinationAddress.Trim().Replace(" ", ""));
                    msg.Subject = argSubject;

                    String rutaFirma = @"C:\Firmas\"; // Ruta local en el equipo que ejecuta el proceso

                    if(String.IsNullOrEmpty(nombreFirma))
                    { rutaFirma += "YobiBasica.jpg"; }
                    else
                    { rutaFirma += nombreFirma.Trim() + ".jpg"; }

                    if (!agregaFirma || !File.Exists(rutaFirma))
                    { msg.Body = argBody.Replace("<img src=\"cid:Firma\">", ""); }

                    if (agregaFirma && File.Exists(rutaFirma))
                    {
                        LinkedResource LinkedImage = new LinkedResource(rutaFirma);

                        LinkedImage.ContentId = "Firma";
                        //Added the patch for Thunderbird as suggested by Jorge
                        
                        LinkedImage.ContentType = new ContentType(MediaTypeNames.Image.Jpeg);

                        AlternateView htmlView = AlternateView.CreateAlternateViewFromString(
                          argBody, null, MediaTypeNames.Text.Html);

                        htmlView.LinkedResources.Add(LinkedImage);
                        msg.AlternateViews.Add(htmlView);
                    }
                    
                    msg.IsBodyHtml = argIsBodyHtml;

                    if (argDestinationCc != null && !String.IsNullOrWhiteSpace(argDestinationCc))
                    {
                        List<string> listaCc = new List<string>(argDestinationCc.Split(';'));
                        foreach (string cc in listaCc)
                        { msg.CC.Add(cc.Trim().Replace(" ", "")); }
                        //msg.CC.Add(argDestinationCc.Trim().Replace(" ", ""));
                    }

                    if (argDestinationBcc != null && !String.IsNullOrWhiteSpace(argDestinationBcc))
                    {
                        List<string> listaBcc = new List<string>(argDestinationBcc.Split(';'));
                        foreach (string bcc in listaBcc)
                        { msg.Bcc.Add(bcc.Trim().Replace(" ", "")); }
                        //msg.Bcc.Add(argDestinationBcc.Trim().Replace(" ", ""));
                    }

                    client.Send(msg);

                    return new Tuple<Boolean, String>(true, "E-mail sent successfuly");
                }
                catch (Exception ex)
                {
                    return new Tuple<Boolean, String>(false, "Error while sending e-mail: " + ex);
                }
            });
        }

        public Task<Tuple<Boolean, String>> sendWithAttachments(String argServer, String argPort, String argDestinationAddress, String argSubject, String argBody, Boolean argIsBodyHtml, String attachedFilePath, String fromEmail, String fromPassword, String fromNickname, String argDestinationCc = null, String argDestinationBcc = null)
        {
            return Task.Run(() =>
            {
                fromNickname = "<" + fromNickname + "> " + fromEmail;

                Console.WriteLine("Server: " + argServer);
                Console.WriteLine("Port: " + argPort);
                Console.WriteLine("DestinationAddress: " + argDestinationAddress);
                Console.WriteLine("CCAddress: " + argDestinationCc);
                Console.WriteLine("BCCAddress: " + argDestinationBcc);
                Console.WriteLine("Subject: " + argSubject);
                Console.WriteLine("Body: ..." + argBody);
                Console.WriteLine("IsBodyHtml: " + argIsBodyHtml);
                Console.WriteLine("FromEmail: " + fromEmail);
                Console.WriteLine("Password: " + fromPassword);
                Console.WriteLine("FromNickname: " + fromNickname);

                try
                {
                    var client = new SmtpClient(argServer, Convert.ToInt32(argPort))
                    {
                        Credentials = new NetworkCredential(fromEmail, fromPassword),
                        EnableSsl = true
                    };

                    MailMessage msg = new MailMessage(fromNickname, argDestinationAddress.Trim().Replace(" ", ""), argSubject, argBody);

                    msg.IsBodyHtml = argIsBodyHtml;

                    if (argDestinationCc != null && !String.IsNullOrWhiteSpace(argDestinationCc))
                    {
                        List<string> listaCc = new List<string>(argDestinationCc.Split(';'));
                        foreach (string cc in listaCc)
                        { msg.CC.Add(cc.Trim().Replace(" ", "")); }
                        //msg.CC.Add(argDestinationCc.Trim().Replace(" ", ""));
                    }

                    if (argDestinationBcc != null && !String.IsNullOrWhiteSpace(argDestinationBcc))
                    {
                        List<string> listaBcc = new List<string>(argDestinationBcc.Split(';'));
                        foreach (string bcc in listaBcc)
                        { msg.Bcc.Add(bcc.Trim().Replace(" ", "")); }
                        //msg.Bcc.Add(argDestinationBcc.Trim().Replace(" ", ""));
                    }

                    Attachment data = new Attachment(attachedFilePath, MediaTypeNames.Application.Octet);

                    ContentDisposition disposition = data.ContentDisposition;
                    disposition.CreationDate = System.IO.File.GetCreationTime(attachedFilePath);
                    disposition.ModificationDate = System.IO.File.GetLastWriteTime(attachedFilePath);
                    disposition.ReadDate = System.IO.File.GetLastAccessTime(attachedFilePath);

                    msg.Attachments.Add(data);

                    client.Send(msg);

                    String successMessage = "E-mail with attachments sent successfuly";
                    Console.WriteLine(successMessage);
                    return new Tuple<Boolean, String>(true, successMessage);
                }
                catch (Exception ex)
                {
                    String failureMessage = "Error while sending e-mail: " + ex;
                    Console.WriteLine(failureMessage);
                    return new Tuple<Boolean, String>(false, failureMessage);
                }
            });
        }
    }
}