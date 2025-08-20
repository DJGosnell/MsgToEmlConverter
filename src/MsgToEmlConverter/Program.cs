using System;
using System.IO;
using MsgReader.Outlook;
using MimeKit;
using System.Text;

namespace MsgToEmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            
            string inputPath;
            string outputPath;
            
            if (args.Length == 0)
            {
                var exeDirectory = AppContext.BaseDirectory;
                inputPath = exeDirectory;
                outputPath = Path.Combine(exeDirectory, "converted");
                
                Console.WriteLine($"Converting all .msg files from: {exeDirectory}");
                Console.WriteLine($"Output directory: {outputPath}");
            }
            else if (args.Length == 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }
            else
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("  MsgToEmlConverter                             - Convert all .msg files from exe directory to 'converted' subfolder");
                Console.WriteLine("  MsgToEmlConverter <input.msg> <output.eml>   - Convert single file");
                Console.WriteLine("  MsgToEmlConverter <input_dir> <output_dir>   - Convert directory");
                return;
            }

            try
            {
                if (File.Exists(inputPath))
                {
                    ConvertSingleFile(inputPath, outputPath);
                }
                else if (Directory.Exists(inputPath))
                {
                    ConvertDirectory(inputPath, outputPath);
                }
                else
                {
                    Console.WriteLine($"Error: Input path '{inputPath}' does not exist.");
                    return;
                }

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }

        static void ConvertSingleFile(string msgFilePath, string emlFilePath, int currentFile = 1, int totalFiles = 1)
        {
            var fileName = Path.GetFileName(msgFilePath);
            Console.Write($"[{currentFile}/{totalFiles}] Converting: {fileName}... ");
            
            var startTime = DateTime.Now;

            using (var msg = new Storage.Message(msgFilePath))
            {
                var mimeMessage = new MimeMessage();

                if (!string.IsNullOrEmpty(msg.Sender?.Email))
                {
                    mimeMessage.From.Add(new MailboxAddress(msg.Sender.DisplayName, msg.Sender.Email));
                }

                if (msg.Recipients != null)
                {
                    foreach (var recipient in msg.Recipients)
                    {
                        var mailboxAddress = new MailboxAddress(recipient.DisplayName, recipient.Email);
                        
                        switch (recipient.Type)
                        {
                            case RecipientType.To:
                                mimeMessage.To.Add(mailboxAddress);
                                break;
                            case RecipientType.Cc:
                                mimeMessage.Cc.Add(mailboxAddress);
                                break;
                            case RecipientType.Bcc:
                                mimeMessage.Bcc.Add(mailboxAddress);
                                break;
                        }
                    }
                }

                mimeMessage.Subject = msg.Subject ?? string.Empty;
                mimeMessage.Date = msg.SentOn ?? DateTimeOffset.Now;

                var bodyBuilder = new BodyBuilder();

                if (!string.IsNullOrEmpty(msg.BodyText))
                {
                    bodyBuilder.TextBody = msg.BodyText;
                }

                if (!string.IsNullOrEmpty(msg.BodyHtml))
                {
                    bodyBuilder.HtmlBody = msg.BodyHtml;
                }

                if (msg.Attachments != null)
                {
                    foreach (var attachment in msg.Attachments)
                    {
                        if (attachment is Storage.Attachment fileAttachment)
                        {
                            var attachmentPart = new MimePart(MimeTypes.GetMimeType(fileAttachment.FileName))
                            {
                                Content = new MimeContent(new MemoryStream(fileAttachment.Data)),
                                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                                ContentTransferEncoding = ContentEncoding.Base64,
                                FileName = fileAttachment.FileName
                            };
                            bodyBuilder.Attachments.Add(attachmentPart);
                        }
                    }
                }

                mimeMessage.Body = bodyBuilder.ToMessageBody();

                var directory = Path.GetDirectoryName(emlFilePath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                mimeMessage.WriteTo(emlFilePath);
            }
            
            var elapsed = DateTime.Now - startTime;
            Console.WriteLine($"✓ Done ({elapsed.TotalMilliseconds:F0}ms)");
        }

        static void ConvertDirectory(string inputDir, string outputDir)
        {
            Directory.CreateDirectory(outputDir);

            var msgFiles = Directory.GetFiles(inputDir, "*.msg", SearchOption.AllDirectories);
            
            if (msgFiles.Length == 0)
            {
                Console.WriteLine("No .msg files found to convert.");
                return;
            }
            
            Console.WriteLine($"Found {msgFiles.Length} .msg file{(msgFiles.Length == 1 ? "" : "s")} to convert.");
            Console.WriteLine();
            
            var overallStartTime = DateTime.Now;
            var successCount = 0;
            var failureCount = 0;

            for (int i = 0; i < msgFiles.Length; i++)
            {
                var msgFile = msgFiles[i];
                var relativePath = Path.GetRelativePath(inputDir, msgFile);
                var emlFile = Path.Combine(outputDir, Path.ChangeExtension(relativePath, ".eml"));
                
                try
                {
                    ConvertSingleFile(msgFile, emlFile, i + 1, msgFiles.Length);
                    successCount++;
                }
                catch (Exception ex)
                {
                    var fileName = Path.GetFileName(msgFile);
                    Console.WriteLine($"[{i + 1}/{msgFiles.Length}] Converting: {fileName}... ✗ Failed ({ex.Message})");
                    failureCount++;
                }
                
                // Show progress bar
                if (msgFiles.Length > 1)
                {
                    var progress = (double)(i + 1) / msgFiles.Length;
                    var progressBar = CreateProgressBar(progress, 40);
                    var percentage = progress * 100;
                    Console.Write($"\rProgress: {progressBar} {percentage:F1}% ({i + 1}/{msgFiles.Length})");
                    
                    if (i == msgFiles.Length - 1)
                    {
                        Console.WriteLine();
                    }
                }
            }
            
            var totalElapsed = DateTime.Now - overallStartTime;
            Console.WriteLine();
            Console.WriteLine($"Conversion completed in {totalElapsed.TotalSeconds:F1} seconds");
            Console.WriteLine($"✓ Success: {successCount} files");
            if (failureCount > 0)
            {
                Console.WriteLine($"✗ Failed: {failureCount} files");
            }
        }
        
        static string CreateProgressBar(double progress, int width)
        {
            var filled = (int)(progress * width);
            var empty = width - filled;
            return "[" + new string('█', filled) + new string('░', empty) + "]";
        }
    }
}
