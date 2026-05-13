# Syncfusion ASP.NET Core – Legal Case File Compilation Demo

This repository contains a complete showcase sample demonstrating how to build an **Automated Legal Case File Compilation System** using **Syncfusion DocIO** and **Syncfusion PDF** libraries in an ASP.NET Core MVC application. The sample illustrates how legal professionals can streamline case preparation by merging multiple legal documents into a single, professionally formatted case file with automatic table of contents, bookmarks, and security features.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── ErrorViewModel.cs
├── Views/
│   ├── Home/
│   │   ├── Index.cshtml
│   │   └── Privacy.cshtml
│   └── Shared/
├── wwwroot/
│   └── Data/
│       ├── Contract.docx
│       ├── Demand_Letter.docx
│       ├── Complaint.docx
│       └── Motion_for_Summary_Judgment.docx
└── README.md
```

---

## ✨ Features

### Document Management
- **Optional Destination Document** – Merge into an existing case file or create a new one automatically
- **Flexible Source Selection** – Choose from default case files or upload custom documents
- **Document Ordering** – Reorder source documents using intuitive up/down arrow controls
- **Drag & Drop Support** – Modern drag-and-drop interface for file uploads
- **Multiple File Formats** – Supports .doc, .docx Word document formats as input

### Merge Configuration
- **Six Import Formatting Options**:
  - Keep Source Formatting
  - Merge Formatting
  - Keep Text Only
  - List Continue Numbering
  - List Restart Numbering
  - Use Destination Styles
- **Automatic Table of Contents** – Generate TOC with customizable heading levels (1-9)
- **Smart Bookmarks** – Add navigational bookmarks for each merged document

### Output Formats
- **Word Document (.docx)** – Editable format with encryption and protection options
- **PDF Document (.pdf)** – Read-only format with password protection and permissions

### Security Features

#### Word Document Security
- **Encryption** – Password-protect documents with AES encryption
- **Document Protection** – Four protection types:
  - Read Only
  - Allow Only Comments
  - Allow Only Form Fields
  - Allow Only Revisions

#### PDF Security
- **User Password** – Require password to open document
- **Owner Password** – Require password to modify permissions
- **Granular Permissions**:
  - Allow/Deny Printing
  - Allow/Deny Content Copying
  - Allow/Deny Annotation Editing
  - Allow/Deny Content Editing

---

## 🚀 Getting Started

### Prerequisites

- [.NET 7.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

---

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Case-File-Compilation
cd Case-File-Compilation
```

---

### 2. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

---

### 3. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

---

### 4. Run the Application

```bash
dotnet run
```

Open your browser and navigate to:
```
https://localhost:5001
```

---

## 📋 How to Use

### Basic Workflow

1. **Select Destination Document (Optional)**
   - Upload an existing case file to merge documents into
   - Or leave empty to create a new compiled document automatically
   - Drag & drop supported

2. **Choose Source Documents**
   - **Option A: Default Files** – Use pre-loaded case files (Contract, Demand Letter, Complaint, Motion)
   - **Option B: Custom Upload** – Upload your own documents (single or multiple)
   - Reorder documents using ▲ ▼ arrow buttons

3. **Configure Merge Settings**
   - Select import formatting behavior from dropdown (6 options)
   - Enable **Table of Contents** (optional)
     - Set lower and upper heading levels (e.g., 1-3 includes Heading 1, 2, and 3)
   - Enable **Bookmarks** to add navigation markers for each document

4. **Choose Output Format**
   - **Word (.docx)** – Select for editable documents
     - Enable security: Encryption OR Protection
     - Enter password if security enabled
   - **PDF (.pdf)** – Select for final distribution
     - Enable security: User/Owner passwords
     - Configure permissions (print, copy, edit)

5. **Generate Case File**
   - Click "Generate Case File" button
   - Download automatically starts
   - File name format: `CompiledCaseFile_YYYYMMDD_HHMMSS.docx/pdf`

---

## 🎯 Use Cases

### Legal Practice Scenarios

- **Case File Compilation** – Merge contracts, correspondence, pleadings, and evidence into comprehensive case files
- **Discovery Document Bundles** – Compile discovery documents with automatic indexing and bookmarks
- **Court Filing Preparation** – Create formatted submissions with TOC for court filings
- **Client Document Packages** – Assemble complete case information packets for client review
- **Appellate Record Preparation** – Compile trial records with proper organization and indexing
- **Settlement Agreement Bundles** – Merge related settlement documents into single packages
- **Due Diligence Reports** – Compile transaction documents with table of contents
- **Contract Management** – Merge master agreements with amendments and exhibits

### Document Types Supported

- Contracts and Agreements
- Legal Briefs and Memoranda
- Pleadings and Motions
- Correspondence and Letters
- Discovery Documents
- Evidence Documentation
- Affidavits and Declarations
- Court Orders and Judgments

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Importing Word Documents](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-word-document#importing-word-documents)
- [Table of Contents](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-table-of-contents)
- [Bookmarks](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-bookmarks)
- [Word Document Security](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-security)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)
- [PDF Security](https://help.syncfusion.com/document-processing/pdf/pdf-library/net/working-with-security)

---

## ✅ Benefits

- **Time Savings** – Automate hours of manual document compilation work
- **Consistency** – Ensure standardized formatting across all case files
- **Professional Output** – Generate court-ready, properly formatted documents
- **Secure Distribution** – Protect confidential legal information with encryption
- **Easy Navigation** – Automatic TOC and bookmarks for quick reference
- **Flexible Workflow** – Adapt to various legal document compilation needs
- **Batch Processing** – Handle multiple documents simultaneously
- **Quality Control** – Reduce manual errors in document assembly

---

## 📊 Workflow Example

### Scenario: Preparing Court Filing Package

1. **Upload Destination**: Firm's standard court filing template
2. **Select Sources**: 
   - Motion brief (Order #1)
   - Supporting affidavit (Order #2)
   - Exhibits A-C (Order #3-5)
   - Proposed order (Order #6)
3. **Configure**:
   - Format: Keep Source Formatting
   - Enable TOC (Levels 1-3)
   - Enable Bookmarks
4. **Output**: PDF with password protection
5. **Result**: Professional court filing package with navigation aids, ready for electronic filing

---

## 📣 Try It Out

Clone the repository, run the sample, and discover how **Syncfusion DocIO** can revolutionize legal document workflows in your practice.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific legal practice requirements.

You can modify the document templates, merge logic, security settings, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our [Syncfusion Support Team](https://support.syncfusion.com/support/tickets/create) for guidance.

---

## 🔐 Security & Compliance

### Best Practices

- **Encryption Passwords**: Use strong passwords with mixed case, numbers, and special characters
- **Protection Types**: Choose appropriate protection levels based on document sensitivity
- **PDF Permissions**: Set restrictive permissions for confidential documents
- **Temporary Files**: Application automatically cleans up temporary files after processing
- **No Storage**: Documents are processed in memory without server-side storage

### Compliance Considerations

- Suitable for attorney-client privileged documents
- Compatible with electronic court filing systems
- Supports legal document retention policies
- Enables secure client communication

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)

---
