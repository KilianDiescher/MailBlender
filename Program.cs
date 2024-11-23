using System;
using System.Collections;
using System.IO.Compression;
using System.Text;
using AE.Net.Mail;

class Program
{
    static List<MailMessage> mails = new List<MailMessage>();
    static List<Attachment> printStack = new List<Attachment>();
    static HashSet<string> processedMessages = new HashSet<string>();
    static string processedMessagesFile = "processedMessages.txt";
    static string blendsDirectory = "ToDo";
    static string blendsOutput = "Finished";

    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        if (!Directory.Exists(blendsOutput))
        {
            Directory.CreateDirectory(blendsOutput);
        }

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
                    if ((attachment.Filename.EndsWith(".blend") || attachment.Filename.EndsWith(".zip")) && !processedMessages.Contains(mail.MessageID))
                    {
                        Console.WriteLine("Received: " + mail.Subject);
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
        // Ensure the directory exists
        if (!Directory.Exists(blendsDirectory))
        {
            Directory.CreateDirectory(blendsDirectory);
        }

        foreach (var attachment in printStack)
        {
            if (attachment.Filename.EndsWith(".blend"))
            {
                string filePath = Path.Combine(blendsDirectory, attachment.Filename);
                File.WriteAllBytes(filePath, attachment.GetData());
            }
            else if (attachment.Filename.EndsWith(".zip"))
            {
                using (var memoryStream = new MemoryStream(attachment.GetData()))
                {
                    using (var archive = new ZipArchive(memoryStream))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            if (entry.FullName.EndsWith(".blend", StringComparison.OrdinalIgnoreCase))
                            {
                                string filePath = Path.Combine(blendsDirectory, entry.FullName);
                                entry.ExtractToFile(filePath, true);
                            }
                        }
                    }
                }
            }
        }
    }
}
