using Case_File_Compilation.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Security;
using System.Diagnostics;

namespace Case_File_Compilation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult GenerateDocument(
            IFormFile destinationFile,
            bool useDefaultFiles,
            IFormFileCollection customFiles,
            string mergeFormat,
            bool includeTOC,
            bool includeBookmarks,
            int? tocLowerHeadingLevel,
            int? tocUpperHeadingLevel,
            string outputFormat,
            string docxSecurityType,
            string docxEncryptionPassword,
            string docxProtectionType,
            string docxProtectionPassword,
            bool enablePdfSecurity,
            string pdfUserPassword,
            string pdfOwnerPassword,
            bool pdfAllowPrint,
            bool pdfAllowCopy,
            bool pdfAllowEditAnnotations,
            bool pdfAllowEditContent)
        {
            try
            {
                // Set defaults
                mergeFormat = mergeFormat ?? "KeepSourceFormatting";
                outputFormat = outputFormat?.ToLower() ?? "pdf";
                int lowerLevel = tocLowerHeadingLevel ?? 1;
                int upperLevel = tocUpperHeadingLevel ?? 3;

                List<Stream> sourceStreams = new List<Stream>();
                List<string> sourceNames = new List<string>();

                // Load source documents
                if (useDefaultFiles)
                {
                    LoadDefaultDocuments(sourceStreams, sourceNames);
                }
                else
                {
                    LoadCustomDocuments(customFiles, sourceStreams, sourceNames);
                }

                // Validate source documents
                if (sourceStreams.Count == 0)
                {
                    return Json(new { success = false, message = "❌ No source documents found." });
                }

                // Create or load destination document
                WordDocument destinationDocument = CreateDestinationDocument(destinationFile);

                try
                {
                    // Add bookmark to destination document ONLY if user provided it
                    if (destinationFile != null && includeBookmarks)
                    {
                        string destinationName = Path.GetFileNameWithoutExtension(destinationFile.FileName);
                        AddBookmark(destinationDocument, destinationName);
                    }
                    // Merge source documents into destination
                    MergeDocuments(destinationDocument, sourceStreams, sourceNames, mergeFormat, includeBookmarks);
                    // Add TOC at the beginning if requested
                    if (includeTOC)
                    {
                        CreateTableOfContents(destinationDocument, lowerLevel, upperLevel);
                        destinationDocument.UpdateTableOfContents();
                    }

                    // Generate output based on format
                    byte[] resultBytes;
                    string contentType;
                    string fileName;

                    if (outputFormat == "pdf")
                    {
                        resultBytes = ConvertToPDF(
                            destinationDocument,
                            enablePdfSecurity,
                            pdfUserPassword,
                            pdfOwnerPassword,
                            pdfAllowPrint,
                            pdfAllowCopy,
                            pdfAllowEditAnnotations,
                            pdfAllowEditContent);
                        contentType = "application/pdf";
                        fileName = $"CompiledCaseFile_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                    }
                    else // docx
                    {
                        // Apply security if requested
                        ApplyDocumentSecurity(destinationDocument, docxSecurityType, docxEncryptionPassword, docxProtectionType, docxProtectionPassword);

                        resultBytes = SaveToDocx(destinationDocument);
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        fileName = $"CompiledCaseFile_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                    }

                    // Cleanup
                    CleanupStreams(sourceStreams);
                    destinationDocument.Dispose();

                    return File(resultBytes, contentType, fileName);
                }
                catch
                {
                    destinationDocument?.Dispose();
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Document generation failed: {ex.Message}");
                return Json(new { success = false, message = $"❌ Error: {ex.Message}" });
            }
        }

        /// <summary>
        /// Creates a new destination document or loads an existing one.
        /// If no destination file is provided, creates a blank document.
        /// </summary>
        private WordDocument CreateDestinationDocument(IFormFile destinationFile)
        {
            WordDocument document;

            if (destinationFile != null && destinationFile.Length > 0)
            {
                // Load existing destination document
                var stream = new MemoryStream();
                destinationFile.CopyTo(stream);
                stream.Position = 0;
                document = new WordDocument(stream, FormatType.Automatic);
            }
            else
            {
                // Create new blank document
                document = new WordDocument();
                // Add a blank section to maintain consistency
                IWSection section = document.AddSection();
                section.PageSetup.Margins.All = 72; // 1 inch margins
            }
            return document;
        }

        /// <summary>
        /// Loads predefined default Word documents from the application's Data folder.
        /// </summary>
        private void LoadDefaultDocuments(List<Stream> streams, List<string> names)
        {
            string[] defaultFiles = new string[]
            {
                "Contract.docx",
                "Demand_Letter.docx",
                "Complaint.docx",
                "Motion_for_Summary_Judgment.docx"
            };

            string defaultDocsPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data");
            foreach (var fileName in defaultFiles)
            {
                string filePath = Path.Combine(defaultDocsPath, fileName);
                if (System.IO.File.Exists(filePath))
                {
                    streams.Add(System.IO.File.OpenRead(filePath));
                    names.Add(Path.GetFileNameWithoutExtension(fileName));
                }
            }
        }

        /// <summary>
        /// Loads user-uploaded custom documents (maintains user-specified order).
        /// </summary>
        private void LoadCustomDocuments(IFormFileCollection files, List<Stream> streams, List<string> names)
        {
            if (files == null) return;

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    var stream = new MemoryStream();
                    file.CopyTo(stream);
                    stream.Position = 0;
                    streams.Add(stream);
                    names.Add(Path.GetFileNameWithoutExtension(file.FileName));
                }
            }
        }
        /// <summary>
        /// Creates and inserts a Table of Contents (TOC) at the beginning of the Word document.
        /// The TOC is generated based on heading styles and outline levels.
        /// </summary>
        private void CreateTableOfContents(WordDocument document, int lowerHeadingLevel, int upperHeadingLevel)
        {
            // Clone the first section or create new one
            WSection tocSection;
            if (document.Sections.Count > 0)
            {
                tocSection = document.Sections[0];
            }
            else
            {
                tocSection = document.AddSection() as WSection;
            }

            // TOC Field
            IWParagraph tocPara = new WParagraph(document);
            tocSection.Body.ChildEntities.Insert(0, tocPara);
            TableOfContent toc = tocPara.AppendTOC(lowerHeadingLevel, upperHeadingLevel);
            toc.UseHeadingStyles = true;
            toc.UseOutlineLevels = true;
            tocPara.ParagraphFormat.AfterSpacing = 12;
            tocPara.AppendBreak(BreakType.PageBreak);

            // TOC Heading
            IWParagraph tocHeadingPara = new WParagraph(document);
            tocSection.Body.ChildEntities.Insert(0, tocHeadingPara);
            IWTextRange tocHeading = tocHeadingPara.AppendText("Table of Contents");
            tocHeading.CharacterFormat.FontSize = 16;
            tocHeading.CharacterFormat.Bold = true;
            tocHeadingPara.ParagraphFormat.AfterSpacing = 12;
            tocHeadingPara.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;           
        }
        /// <summary>
        /// Adds a bookmark to the specified section of the Word document.
        /// The bookmark is created by inserting a start and end marker.
        /// </summary>
        private void AddBookmark(WordDocument document, string bookmarkName)
        {
            // Safety check: Ensure document has sections
            if (document == null || document.Sections.Count == 0)
            {
                return;
            }
            // Get first section
            IWSection firstSection = document.Sections[0];
            // Safety check: Ensure section has at least one paragraph, if not add one
            if (firstSection.Paragraphs.Count == 0)
            {
                firstSection.AddParagraph();
            }
            // Create bookmark start paragraph
            IWParagraph bookmarkStartPara = new WParagraph(document);
            bookmarkStartPara.AppendBookmarkStart(bookmarkName);

            // Insert at the beginning of FIRST section
            firstSection.Paragraphs.Insert(0, bookmarkStartPara);

            // Create bookmark end paragraph
            IWParagraph bookmarkEndPara = new WParagraph(document);
            bookmarkEndPara.AppendBookmarkEnd(bookmarkName);

            // Add at the end of LAST section
            document.LastSection.Paragraphs.Add(bookmarkEndPara);
        }
        /// <summary>
        /// Merges multiple Word documents into a single document,
        /// inserting page breaks and bookmarks for each merged document.
        /// </summary>
        private void MergeDocuments(
            WordDocument destinationDocument,
            List<Stream> sourceStreams,
            List<string> sourceNames,
            string mergeFormat,
            bool includeBookmarks)
        {
            // Parse merge format string to ImportOptions enum
            ImportOptions importOption = ParseMergeFormat(mergeFormat);

            for (int i = 0; i < sourceStreams.Count; i++)
            {
                using (WordDocument sourceDoc = new WordDocument(sourceStreams[i], FormatType.Automatic))
                {
                    // Add bookmark if requested
                    if (includeBookmarks)
                    {
                        IWParagraph bookmarkStartPara = new WParagraph(destinationDocument);
                        bookmarkStartPara.AppendBookmarkStart(sourceNames[i]);
                        sourceDoc.Sections[0].Paragraphs.Insert(0, bookmarkStartPara);
                        AddBookmark(sourceDoc,sourceNames[i]);
                    }

                    // Import content
                    destinationDocument.ImportContent(sourceDoc, importOption);

                    // Add bookmark if requested
                    if (includeBookmarks)
                    {
                        IWParagraph bookmarkEndPara = new WParagraph(destinationDocument);
                        bookmarkEndPara.AppendBookmarkEnd(sourceNames[i]);
                        destinationDocument.LastSection.Paragraphs.Add(bookmarkEndPara);
                    }
                }
            }
        }
        /// <summary>
        /// Converts merge format string to ImportOptions enum.
        /// </summary>
        private ImportOptions ParseMergeFormat(string mergeFormat)
        {
            return mergeFormat switch
            {
                "KeepSourceFormatting" => ImportOptions.KeepSourceFormatting,
                "MergeFormatting" => ImportOptions.MergeFormatting,
                "KeepTextOnly" => ImportOptions.KeepTextOnly,
                "ListContinueNumbering" => ImportOptions.ListContinueNumbering,
                "ListRestartNumbering" => ImportOptions.ListRestartNumbering,
                "UseDestinationStyles" => ImportOptions.UseDestinationStyles,
                _ => ImportOptions.KeepSourceFormatting
            };
        }

        /// <summary>
        /// Applies security settings to the document (encryption or protection).
        /// </summary>
        private void ApplyDocumentSecurity(
            WordDocument document,
            string securityType,
            string encryptionPassword,
            string protectionType,
            string protectionPassword)
        {
            if (string.IsNullOrEmpty(securityType)) return;

            if (securityType.ToLower() == "encryption" && !string.IsNullOrEmpty(encryptionPassword))
            {
                // Apply encryption
                document.EncryptDocument(encryptionPassword);
            }
            else if (securityType.ToLower() == "protection" && !string.IsNullOrEmpty(protectionPassword))
            {
                // Parse protection type
                ProtectionType protType = ParseProtectionType(protectionType);

                // Apply protection
                document.Protect(protType, protectionPassword);
            }
        }

        /// <summary>
        /// Parses protection type string to protectionType enum.
        /// </summary>
        private ProtectionType ParseProtectionType(string protectionType)
        {
            return protectionType switch
            {
                "ReadOnly" => ProtectionType.AllowOnlyReading,
                "AllowOnlyComments" => ProtectionType.AllowOnlyComments,
                "AllowOnlyFormFields" => ProtectionType.AllowOnlyFormFields,
                "AllowOnlyRevisions" => ProtectionType.AllowOnlyRevisions,
                _ => ProtectionType.AllowOnlyReading
            };
        }
        /// <summary>
        /// Converts Word document to PDF with optional security settings.
        /// </summary>
        private byte[] ConvertToPDF(
            WordDocument document,
            bool enableSecurity,
            string userPassword,
            string ownerPassword,
            bool allowPrint,
            bool allowCopy,
            bool allowEditAnnotations,
            bool allowEditContent)
        {
            using (DocIORenderer renderer = new DocIORenderer())
            {
                using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
                {
                    // Apply security if enabled
                    if (enableSecurity && (!string.IsNullOrEmpty(userPassword) || !string.IsNullOrEmpty(ownerPassword)))
                    {
                        PdfSecurity security = pdfDocument.Security;

                        // Set passwords
                        if (!string.IsNullOrEmpty(userPassword))
                            security.UserPassword = userPassword;

                        if (!string.IsNullOrEmpty(ownerPassword))
                            security.OwnerPassword = ownerPassword;

                        // Permissions are only set if owner password is provided
                        if (!string.IsNullOrEmpty(ownerPassword))
                        {
                            PdfPermissionsFlags permissions = PdfPermissionsFlags.Default;

                            if (allowPrint)
                                permissions |= PdfPermissionsFlags.Print;

                            if (allowCopy)
                                permissions |= PdfPermissionsFlags.CopyContent;

                            if (allowEditAnnotations)
                                permissions |= PdfPermissionsFlags.EditAnnotations;

                            if (allowEditContent)
                                permissions |= PdfPermissionsFlags.EditContent | PdfPermissionsFlags.AssembleDocument | PdfPermissionsFlags.FillFields;

                            if (permissions != PdfPermissionsFlags.Default || !string.IsNullOrEmpty(ownerPassword))
                            {
                                security.Permissions = permissions;
                            }
                        }
                    }
                    using (MemoryStream pdfStream = new MemoryStream())
                    {
                        pdfDocument.Save(pdfStream);
                        return pdfStream.ToArray();
                    }
                }
            }
        }
        /// <summary>
        /// Saves Word document to byte array.
        /// </summary>
        private byte[] SaveToDocx(WordDocument document)
        {
            using (MemoryStream docxStream = new MemoryStream())
            {
                document.Save(docxStream, FormatType.Docx);
                return docxStream.ToArray();
            }
        }
        /// <summary>
        /// Disposes all streams in the provided collection to free resources.
        /// </summary>
        private void CleanupStreams(List<Stream> streams)
        {
            // Dispose each stream safely
            foreach (var stream in streams)
            {
                stream?.Dispose();
            }
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
