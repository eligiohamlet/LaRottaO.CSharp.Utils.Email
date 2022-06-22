using System;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading.Tasks;

namespace LaRottaO.CSharp.Utils.Email
{
    public class SendEmail
    {
        public Task<Tuple<Boolean, String>> sendWithoutAttachments(String argServer, String argPort, String argDestinationAddress, String argSubject, String argBody, Boolean argIsBodyHtml, String fromEmail, String fromPassword, String fromNickname, String argDestinationCc = null, String argDestinationBcc = null)
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

                    MailMessage msg = new MailMessage(fromNickname, argDestinationAddress, argSubject, argBody);

                    msg.IsBodyHtml = argIsBodyHtml;

                    if (argDestinationCc != null)
                    {
                        msg.CC.Add(argDestinationCc);
                    }

                    if (argDestinationBcc != null)
                    {
                        msg.Bcc.Add(argDestinationBcc);
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

                    MailMessage msg = new MailMessage(fromNickname, argDestinationAddress, argSubject, argBody);

                    msg.IsBodyHtml = argIsBodyHtml;

                    if (argDestinationCc != null)
                    {
                        msg.CC.Add(argDestinationCc);
                    }

                    if (argDestinationBcc != null)
                    {
                        msg.Bcc.Add(argDestinationBcc);
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