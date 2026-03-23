using System;
using System.IO;
using MimeKit;

namespace EmlPdfExtractor {
    class Program {
        static void Main(string[] args) {
            var baseDir = AppContext.BaseDirectory;
            var inputDirectory = Path.Combine(baseDir, "EMLFiles");
            var outputDirectory = Path.Combine(baseDir, "ExtractedPDFs");

            if (!Directory.Exists(outputDirectory)) {
                Directory.CreateDirectory(outputDirectory);
            }

            var emlFiles = Directory.GetFiles(inputDirectory, "*.eml", SearchOption.TopDirectoryOnly);

            Console.WriteLine($"Found {emlFiles.Length} .eml files");

            foreach (var file in emlFiles) {
                try {
                    ProcessEml(file, outputDirectory);
                }
                catch (Exception ex) {
                    Console.WriteLine($"Error processing {file}: {ex.Message}");
                }
            }

            Console.WriteLine("Finished. Press any key to exit...");
            Console.ReadKey();
        }

        static void ProcessEml(string emlPath, string outputDirectory) {
            using var stream = File.OpenRead(emlPath);
            var message = MimeMessage.Load(stream);

            foreach (var attachment in message.Attachments) {
                if (attachment is MimePart part) {
                    if (part.ContentType.MimeType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase)) {
                        /*var fileName = part.FileName;

                        if (string.IsNullOrWhiteSpace(fileName)) {
                            fileName = Path.GetFileNameWithoutExtension(emlPath) + ".pdf";
                        }*/

                        var emlBaseName = Path.GetFileNameWithoutExtension(emlPath) + ".pdf";
                        var outputPath = Path.Combine(outputDirectory, emlBaseName);

                        using var outputStream = File.Create(outputPath);
                        part.Content!.DecodeTo(outputStream);

                        Console.WriteLine($"Extracted PDF: {emlBaseName}");
                    }
                }
            }
        }
    }
}