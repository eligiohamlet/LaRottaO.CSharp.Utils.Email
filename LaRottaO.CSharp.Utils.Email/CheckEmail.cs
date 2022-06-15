using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace LaRottaO.CSharp.Utils.Email
{
    public class CheckEmail
    {
        /// <summary>
        ///
        /// Downloads all unread messages from an email account
        /// and their attachments
        /// Wrapping the nice Mailkit library
        ///
        /// </summary>

        public Task<Tuple<Boolean, String, List<EmailMessage>>> checkUnread(String argUsername, String argPassword, String argSmptpServer, int argPort, Boolean argMarkAsRead, String arfgromAddressContains = null, String argSubjectContains = null, String argBodyContains = null, Boolean argDownloadAttachments = false)
        {
            return Task.Run(() =>
            {
                try
                {
                    List<EmailMessage> listaSalida = new List<EmailMessage>();

                    using (var client = new ImapClient())
                    {
                        using (var cancel = new CancellationTokenSource())
                        {
                            client.Connect(argSmptpServer, argPort, true, cancel.Token);

                            client.AuthenticationMechanisms.Remove("XOAUTH");

                            client.Authenticate(argSmptpServer, argPassword, cancel.Token);

                            var inbox = client.Inbox;
                            inbox.Open(FolderAccess.ReadWrite, cancel.Token);

                            var query = SearchQuery.NotSeen;

                            IList<UniqueId> listaCorreosNoLeidos = inbox.Search(query, cancel.Token);

                            Console.WriteLine("Unread e-mails: " + listaCorreosNoLeidos.Count);

                            foreach (var uid in listaCorreosNoLeidos)
                            {
                                var message = inbox.GetMessage(uid, cancel.Token);

                                EmailMessage emailMessage = new EmailMessage();

                                emailMessage.uniqueId = uid;
                                emailMessage.from = message.From.First().ToString();
                                emailMessage.subject = message.Subject;
                                emailMessage.receivedDate = message.Date;
                                emailMessage.body = message.GetTextBody(MimeKit.Text.TextFormat.Plain);

                                Boolean isAValidMatch = false;

                                if ((arfgromAddressContains == null && argSubjectContains == null && argBodyContains == null) ||
                                    (arfgromAddressContains != null && emailMessage.from.ToLower().Contains(arfgromAddressContains.ToLower())) ||
                                    (argSubjectContains != null && emailMessage.subject.ToLower().Contains(argSubjectContains.ToLower())) ||
                                    (argBodyContains != null && emailMessage.body.ToLower().Contains(argBodyContains.ToLower())))
                                {
                                    isAValidMatch = true;
                                }

                                if (!isAValidMatch)
                                {
                                    continue;
                                }

                                if (argMarkAsRead)
                                {
                                    inbox.AddFlags(emailMessage.uniqueId, MessageFlags.Seen, true);
                                }

                                if (argDownloadAttachments)
                                {
                                    foreach (MimeEntity attachment in message.Attachments)
                                    {
                                        var fileName = emailMessage.uniqueId + "_" + attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                                        String rutaArchivo = AppContext.BaseDirectory + @"\Downloaded Attachments\";

                                        if (!Directory.Exists(rutaArchivo))
                                        {
                                            Directory.CreateDirectory(rutaArchivo);
                                        }

                                        Console.WriteLine("Downloading " + fileName + "...");

                                        using (var stream = File.Create(rutaArchivo + fileName))
                                        {
                                            if (attachment is MessagePart)
                                            {
                                                var rfc822 = (MessagePart)attachment;

                                                rfc822.Message.WriteTo(stream);
                                            }
                                            else
                                            {
                                                var part = (MimePart)attachment;

                                                part.Content.DecodeTo(stream);
                                            }

                                            emailMessage.downloadedAttachments.Add(rutaArchivo + fileName);
                                        }
                                    }
                                }

                                listaSalida.Add(emailMessage);
                            }

                            client.Disconnect(true, cancel.Token);
                        }
                    }

                    return new Tuple<Boolean, String, List<EmailMessage>>(true, "", listaSalida);
                }
                catch (Exception ex)
                {
                    return new Tuple<Boolean, String, List<EmailMessage>>(false, ex.ToString(), new List<EmailMessage>());
                }
            });
        }
    }
}