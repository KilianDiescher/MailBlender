using System;
using System.Collections;
using System.Text;
using AE.Net.Mail;

class Program
{
    static List<MailMessage> mails = new List<MailMessage>();
    static List<Attachment> printStack = new List<Attachment>();
    static HashSet<string> processedMessages = new HashSet<string>();
    static string processedMessagesFile = "processedMessages.txt";

    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        LoadProcessedMessages();

        Timer timer = new Timer(RunScript, null, TimeSpan.Zero, TimeSpan.FromMinutes(0.5));

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static void RunScript(object state)
    {
        using (var client = new ImapClient("imap.gmx.com", "RenderForMe@gmx.net", "RFM@GMX.NET", AuthMethods.Login, 993, true))
        {
            int messageCount = client.GetMessageCount();
            for (int i = 0; i < messageCount; i++)
            {
                var message = client.GetMessage(i);
                Console.WriteLine("Received " + $"Subject: {message.Subject}");
                mails.Add(message);
            }

            client.Disconnect();
        }

        getBlends();
        saveBlends();
        SaveProcessedMessages();
    }

    static void getBlends()
    {
        foreach (var mail in mails)
        {
            if (mail.Attachments.Count > 0)
            {
                foreach (var attachment in mail.Attachments)
                {
                    if (attachment.Filename.EndsWith(".blend") && !processedMessages.Contains(mail.MessageID))
                    {
                        printStack.Add(attachment);
                        processedMessages.Add(mail.MessageID);
                    }
                }
            }
        }
    }

    static void LoadProcessedMessages()
    {
        if (File.Exists(processedMessagesFile))
        {
            var lines = File.ReadAllLines(processedMessagesFile);
            foreach (var line in lines)
            {
                processedMessages.Add(line);
            }
        }
    }

    static void SaveProcessedMessages()
    {
        File.WriteAllLines(processedMessagesFile, processedMessages);
    }

    static void saveBlends()
    {
        foreach (var attachment in printStack)
        {
            File.WriteAllBytes(attachment.Filename, attachment.GetData());
        }
    }
}
