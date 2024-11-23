using System.Diagnostics;
using System.IO.Compression;
using System.Net.Mail;
using System.Text;
using AE.Net.Mail;
using System.Data;
using System.Data.SQLite;

class MailBlender
{
    static List<AE.Net.Mail.MailMessage> mails = new List<AE.Net.Mail.MailMessage>();
    static List<AE.Net.Mail.Attachment> printStack = new List<AE.Net.Mail.Attachment>();
    static HashSet<string> processedMessages = new HashSet<string>();
    static string processedMessagesFile = "processedMessages.txt";
    static string addresssDb = "addresses.db";
    static string blendsDirectory = "ToDo";
    static string blendsOutput = "Finished";
    static string blendsDone = "Done";

    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


        if (!Directory.Exists(blendsOutput))
        {
            Directory.CreateDirectory(blendsOutput);
        }

        if (!Directory.Exists(blendsDone))
        {
            Directory.CreateDirectory(blendsDone);
        }

        if (!File.Exists(addresssDb))
        {
            using (FileStream fs = File.Create(addresssDb))
            {
                fs.Close();
            }
        }

        SQLiteConnection connection = new SQLiteConnection("Data Source=" + addresssDb);
        
        connection.Open();
        

        using (var command = new SQLiteCommand(connection))
        {
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS Contacts (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    MessageID TEXT NOT NULL,
                    Email TEXT NOT NULL
                )";
            command.ExecuteNonQuery();
        }

        connection.Close();

        LoadProcessedMessages();

        Timer timer = new Timer(RunScript, null, TimeSpan.Zero, TimeSpan.FromMinutes(0.5));

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }


    static void saveToDb(string messageID, string email)
    {
        using (SQLiteConnection connection = new SQLiteConnection("Data Source=" + addresssDb))
        {
            connection.Open();
            using (var command = new SQLiteCommand(connection))
            {
                // Check if the email already exists
                command.CommandText = "SELECT COUNT(*) FROM Contacts WHERE MessageID = @MessageID";
                command.Parameters.AddWithValue("@MessageID", messageID);
                long count = (long)command.ExecuteScalar();

                if (count == 0)
                {
                    // Insert the email if it does not exist
                    command.CommandText = "INSERT INTO Contacts (MessageID, Email) VALUES (@MessageID, @Email)";
                    command.Parameters.AddWithValue("@MessageID", messageID);
                    command.Parameters.AddWithValue("@Email", email);
                    command.ExecuteNonQuery();
                }
            }
        }
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
                saveToDb(message.MessageID, message.From.Address);
            }

            client.Disconnect();
        }

        getBlends();
        saveBlends();
        SaveProcessedMessages();
        foreach (var f in getToDo())
        {
            renderAnimation(f);
        }
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
        else
        {
            File.Create(processedMessagesFile);
            LoadProcessedMessages();
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

    static string[] getToDo()
    {
        string[] files = Directory.GetFiles(blendsDirectory);
        return files;
    }

    static void renderAnimation(string file)
    {
        string blendFile = file;
        string blendFileName = Path.GetFileName(blendFile);
        string blendOutput = Path.Combine(Path.GetFullPath(blendsOutput), blendFileName);


        if (Directory.Exists(blendOutput))
        {
            Directory.Delete(blendOutput, true);
        }

        Directory.CreateDirectory(blendOutput);

        // Find Blender executable in PATH
        string blenderPath = FindExecutableInPath("blender.exe");
        if (blenderPath == null)

        {
            Console.WriteLine("Blender executable not found in PATH.");
            return;
        }

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-Command \"& '{blenderPath}' -b '{blendFile}' -o '{blendOutput}/' -a\"",
            RedirectStandardOutput = false,
            UseShellExecute = true,
            CreateNoWindow = true,
        };

        Process process = new Process
        {
            StartInfo = startInfo
        };
        process.Start();
        process.WaitForExit();

        string doneFilePath = Path.Combine(Path.GetFullPath(blendsDone), Path.GetFileName(file));
        File.Move(file, doneFilePath);
        file.PadLeft(10);

        sendBack(file);
    }

    static string FindExecutableInPath(string executableName)
    {
        var paths = Environment.GetEnvironmentVariable("Path").Split(';');
        foreach (var path in paths)
        {
            var fullPath = Path.Combine(path, executableName);
            
            if (File.Exists(fullPath))
            {
                return fullPath;
            }
        }
        return null;
    }

    static void sendBack(string file)
    {
        string finishedFileDirectory = Path.Combine(blendsOutput, Path.GetFileName(file));
        string zipFilePath = Path.Combine(blendsOutput, Path.GetFileNameWithoutExtension(file) + ".zip");

        if (!Directory.Exists(finishedFileDirectory))
        {
            Console.WriteLine("Directory not found in the Finished folder.");
            return;
        }

        if (File.Exists(zipFilePath))
        {
            File.Delete(zipFilePath);
        }
        ZipFile.CreateFromDirectory(finishedFileDirectory, zipFilePath);

        Console.WriteLine("Sending back: " + zipFilePath);

        using (var client = new ImapClient("imap.gmx.com", "RenderForMe@gmx.net", "RFM@GMX.NET", AuthMethods.Login, 993, true))
        {
            SmtpClient smtp = new SmtpClient("mail.gmx.com")
            {
                Port = 587,
                EnableSsl = true,
                Credentials = new System.Net.NetworkCredential("RenderForMe@gmx.net", "RFM@GMX.NET")
            };

            var message = new System.Net.Mail.MailMessage
            {
                From = new MailAddress("RenderForMe@Gmx.net"),
                Subject = "Rendered Animation",
                Body = "Your animation has been rendered.",
                Attachments = { new System.Net.Mail.Attachment(zipFilePath) }
            };

            //message.To.Add(new MailAddress())
        }
    }

}
