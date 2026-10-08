# Xceed Words for .NET – Word Processing & PDF Manipulation API for .NET

![Xceed Banner](https://raw.githubusercontent.com/xceedsoftware/media/refs/heads/main/nuget/xceed.words.net_banner.png)

---

[![NuGet](https://img.shields.io/nuget/v/Xceed.Words.NET)](https://www.nuget.org/packages/Xceed.Words.NET/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Xceed.Words.NET)](https://www.nuget.org/packages/Xceed.Words.NET/)
[![GitHub License](https://img.shields.io/github/license/xceedsoftware/DocX)](https://github.com/xceedsoftware/DocX/blob/master/LICENSE)

---

## Xceed Words for .NET – Word Processing & PDF Manipulation API for .NET

[![Product](https://img.shields.io/badge/Product-FF6F00?style=for-the-badge&logo=googlechrome&logoColor=white)](https://xceed.com/products/net/words-for-net/)
[![Docs](https://img.shields.io/badge/Docs-FF6F00?style=for-the-badge&logo=readthedocs&logoColor=white)](https://xceed.com/products/net/words-for-net/documentation/)
[![API Reference](https://img.shields.io/badge/API_Reference-FF6F00?style=for-the-badge&logo=html5&logoColor=white)](https://xceed.com/products/net/words-for-net/documentation/)
[![Examples](https://img.shields.io/badge/Examples-FF6F00?style=for-the-badge&logo=github&logoColor=white)](https://github.com/xceedsoftware/DocX/tree/master/Examples)
[![Blog](https://img.shields.io/badge/Blog-FF6F00?style=for-the-badge&logo=wordpress&logoColor=white)](https://xceed.com/blog/)
[![Releases](https://img.shields.io/badge/Releases-FF6F00?style=for-the-badge&logo=rocket&logoColor=white)](https://xceed.com/blog/)
[![Support](https://img.shields.io/badge/Support-FF6F00?style=for-the-badge&logo=googlemessages&logoColor=white)](https://xceed.com/support/)
[![License](https://img.shields.io/badge/License-FF6F00?style=for-the-badge&logo=rocket&logoColor=white)](https://xceed.com/licensing/)

---

## Overview

Xceed Words for .NET is an enterprise-grade word processing and PDF manipulation library for creating, editing, converting, and managing Word documents and PDF files programmatically. Create invoices, reports, contracts, and business documents with full control over formatting, styles, and layout.

Convert DOCX documents to PDF, merge and split files, apply templates, protect documents, and automate document workflows. No Microsoft Office installation required. Works seamlessly with ASP.NET, cloud platforms, and cross-platform .NET applications.

Trusted by Fortune 500 companies and .NET developers worldwide.

## Key Features

### Word Document Processing
- Create DOCX documents from scratch
- Open, edit, and modify existing Word files
- Insert paragraphs, tables, images, and shapes
- Insert HTML content with HTML5 parsing and CSS support
- Add headers, footers, and watermarks
- Apply styles, fonts, colors, and formatting
- Insert hyperlinks and bookmarks
- Generate dynamic reports and templates

### PDF Generation & Conversion
- Convert DOCX to PDF with high fidelity
- Generate PDF documents programmatically
- Add text, images, and graphics to PDFs
- Create searchable PDF files
- Apply password protection and encryption
- Add digital signatures
- Merge and split PDF documents

### Document Manipulation
- Merge multiple documents into one
- Split large documents by configurable criteria (SplitOptions / SplitCriteria)
- Clone and copy document elements
- Manage document properties and metadata
- Protect documents with passwords
- Set document permissions

### Charting
- Area, bar, bubble, line, pie, doughnut, scatter, radar, sunburst, surface, waterfall, and combo charts
- Data labels with canonical ordering
- Per-data-point coloring
- Configurable gridlines and chart styling

### Advanced Features
- Template-based document generation
- Batch document processing
- Parallel processing support
- Mail merge functionality
- Automatic numbering and bullets
- Table of contents generation
- Custom styles and formatting

### Format Support
- Read and write DOCX files
- PDF generation and manipulation
- Support for DOC, RTF, ODT formats
- Insert HTML content into documents (HTML5-compliant parsing)
- Export to HTML (with embedded font handling and configurable CSS prefixes)
- Image format support (PNG, JPEG, BMP, TIFF)
- SVG graphics support

### Cross-Platform
- Windows, Linux, macOS support
- ASP.NET and ASP.NET Core compatible
- Blazor support
- Cloud platform ready
- No Office interop required
- Fully managed .NET code

## Installation

### NuGet Package Manager
```bash
dotnet add package Xceed.Words.NET
```

### Package Manager Console
```powershell
Install-Package Xceed.Words.NET
```

## Quick Start

### Create a Word Document

```csharp
using Xceed.Words.NET;

var document = DocX.Create("Report.docx");

document.InsertParagraph("Hello from Xceed.Words.NET")
        .Bold()
        .FontSize(18);

document.InsertParagraph("This is a test document with formatting.")
        .Italic()
        .FontSize(12);

document.Save();
```

### Edit an Existing Document

```csharp
using Xceed.Words.NET;

var document = DocX.Load("existing.docx");

document.InsertParagraph("New content added to document.")
        .Color(System.Drawing.Color.Blue);

document.ReplaceText("{CustomerName}", "John Smith");

document.Save();
```

### Convert DOCX to PDF

```csharp
using Xceed.Words.NET;

var document = DocX.Load("Report.docx");

// Convert to PDF
document.SaveAsPdf("Report.pdf");
```

### Create Table in Word Document

```csharp
using Xceed.Words.NET;

var document = DocX.Create("TableDocument.docx");

document.InsertParagraph("Sales Report").Bold().FontSize(16);

var table = document.InsertTable(4, 3);
table.Rows[0].Cells[0].Paragraphs.First().InsertText("Product");
table.Rows[0].Cells[1].Paragraphs.First().InsertText("Q1");
table.Rows[0].Cells[2].Paragraphs.First().InsertText("Q2");

document.Save();
```

## Core Capabilities

### Document Creation & Editing
| Feature | Supported |
|---------|-----------|
| Create DOCX from scratch | ✓ |
| Open and edit documents | ✓ |
| Insert paragraphs | ✓ |
| Insert tables | ✓ |
| Insert images | ✓ |
| Insert HTML content (HTML5 parsing) | ✓ |
| Add headers/footers | ✓ |
| Add hyperlinks | ✓ |
| Add bookmarks | ✓ |

### Text & Formatting
| Feature | Supported |
|---------|-----------|
| Font styling (bold, italic, underline) | ✓ |
| Font size and color | ✓ |
| Paragraph alignment | ✓ |
| Spacing and indentation | ✓ |
| Styles and themes | ✓ |
| Watermarks | ✓ |
| Numbered and bullet lists | ✓ |

### Charting
| Feature | Supported |
|---------|-----------|
| Bar, line, area, scatter, bubble charts | ✓ |
| Pie, doughnut, radar, sunburst charts | ✓ |
| Surface, waterfall, combo charts | ✓ |
| Data labels | ✓ |
| Per-data-point coloring | ✓ |
| Gridline configuration | ✓ |

### PDF Features
| Feature | Supported |
|---------|-----------|
| DOCX to PDF conversion | ✓ |
| PDF generation | ✓ |
| Password protection | ✓ |
| Encryption | ✓ |
| Digital signatures | ✓ |
| Merge PDFs | ✓ |
| Split PDFs | ✓ |

### Advanced Features
| Feature | Supported |
|---------|-----------|
| Document merging | ✓ |
| Document splitting (SplitOptions / SplitCriteria) | ✓ |
| Mail merge | ✓ |
| Find and replace | ✓ |
| Content cloning | ✓ |
| Batch processing | ✓ |
| Parallel processing | ✓ |

## Supported File Formats

### Read & Write
| Format | Support |
|--------|---------|
| DOCX | ✓ |
| PDF | ✓ |
| DOC | ✓ |
| RTF | ✓ |
| ODT | ✓ |

### Import / Insert
| Format | Support |
|--------|---------|
| HTML (HTML5-compliant parsing) | ✓ |

### Export To
| Format | Support |
|--------|---------|
| PDF | ✓ |
| HTML (embedded fonts, configurable CSS prefix) | ✓ |
| Markdown | ✓ |
| Images (PNG, JPEG, TIFF) | ✓ |
| SVG | ✓ |

## Supported Frameworks

- .NET Framework 4.5+
- .NET Core
- .NET 5+
- .NET 6+
- .NET 7+
- .NET 8+
- ASP.NET
- ASP.NET Core
- Blazor
- Xamarin

## Supported Platforms

- Windows
- Linux
- macOS

## Common Use Cases

- **Invoice Generation** - Create professional invoices automatically
- **Report Generation** - Generate formatted business reports dynamically
- **Contract Automation** - Auto-generate contracts with variable data
- **Certificate Creation** - Produce certificates and diplomas
- **Resume Generation** - Build resume documents programmatically
- **Document Workflow Automation** - Automate document creation pipelines
- **HTML Content Import** - Convert HTML content into formatted Word documents
- **Document Assembly** - Merge sources and generate a table of contents automatically
- **PDF Export** - Convert business documents to PDF
- **Email Attachments** - Generate documents for email distribution
- **ASP.NET Services** - Provide document generation in web applications
- **Document APIs** - Build REST APIs that generate Word and PDF files

## Free 45-Day Trial

Experience the full capabilities of Xceed Words for .NET with a free 45-day trial. No credit card required.

[Start Your Free Trial](https://xceed.com/trial/)

## Resources

- **[Product Page](https://xceed.com/products/net/words-for-net/)** - Features and pricing
- **[Documentation](https://xceed.com/products/net/words-for-net/documentation/)** - Complete API documentation
- **[API Reference](https://xceed.com/products/net/words-for-net/documentation/)** - Detailed method and class reference
- **[GitHub Examples](https://github.com/xceedsoftware/DocX/tree/master/Examples)** - Code examples and sample projects
- **[Blog](https://xceed.com/blog/)** - Updates and articles
- **[Support](https://xceed.com/support/)** - Technical support and forums
- **[Licensing](https://xceed.com/licensing/)** - License information and pricing

## Related Xceed Products

### Document & Office Automation
- **[Xceed PDF Library for .NET](https://www.nuget.org/packages/Xceed.PdfLibrary.NET)** - Advanced PDF creation and manipulation
- **[Xceed Workbooks for .NET](https://www.nuget.org/packages/Xceed.Workbooks.NET)** - Excel workbook creation and editing
- **[DocX](https://www.nuget.org/packages/DocX)** - Lightweight free Word document library

### File Transfer & Compression
- **[Xceed Zip for .NET](https://www.nuget.org/packages/Xceed.Products.Zip.Full)** - ZIP, TAR, GZip compression
- **[Xceed Real-Time Zip for .NET](https://www.nuget.org/packages/Xceed.Products.RealTimeZip.Full)** - Streaming ZIP compression
- **[Xceed SFTP for .NET](https://www.nuget.org/packages/Xceed.Products.SFtp.Full)** - Secure SFTP file transfer
- **[Xceed FTP for .NET](https://www.nuget.org/packages/Xceed.Products.Ftp.Full)** - FTP and FTPS file transfer

### UI Components
- **[Xceed Toolkit Plus for WPF](https://www.nuget.org/packages/Xceed.Products.Wpf.Toolkit.Full)** - Advanced WPF controls
- **[Xceed DataGrid for WPF](https://www.nuget.org/packages/Xceed.Products.Wpf.DataGrid.Full)** - High-performance WPF DataGrid
- **[Xceed Toolkit for .NET MAUI](https://www.nuget.org/packages/Xceed.Product.Maui.Toolkit.Full)** - Cross-platform MAUI controls
- **[Extended WPF Toolkit](https://www.nuget.org/packages/Extended.Wpf.Toolkit)** - Open-source WPF controls

## Getting Help

- **Documentation** - [Complete API documentation and guides](https://xceed.com/products/net/words-for-net/documentation/)
- **Examples** - [GitHub code samples and projects](https://github.com/xceedsoftware/DocX/tree/master/Examples)
- **Support Portal** - [Get technical support](https://xceed.com/support/)
- **NuGet Package** - [View on NuGet.org](https://www.nuget.org/packages/Xceed.Words.NET/)

## License

Xceed Words for .NET is a commercial product with a free 45-day trial. Various license options are available for development and production use. See [licensing information](https://xceed.com/licensing/) for details.

---

[![Product](https://img.shields.io/badge/Product-FF6F00?style=for-the-badge&logo=googlechrome&logoColor=white)](https://xceed.com/products/net/words-for-net/)
[![Docs](https://img.shields.io/badge/Docs-FF6F00?style=for-the-badge&logo=readthedocs&logoColor=white)](https://xceed.com/products/net/words-for-net/documentation/)
[![API Reference](https://img.shields.io/badge/API_Reference-FF6F00?style=for-the-badge&logo=html5&logoColor=white)](https://xceed.com/products/net/words-for-net/documentation/)
[![Examples](https://img.shields.io/badge/Examples-FF6F00?style=for-the-badge&logo=github&logoColor=white)](https://github.com/xceedsoftware/DocX/tree/master/Examples)
[![Blog](https://img.shields.io/badge/Blog-FF6F00?style=for-the-badge&logo=wordpress&logoColor=white)](https://xceed.com/blog/)
[![Releases](https://img.shields.io/badge/Releases-FF6F00?style=for-the-badge&logo=rocket&logoColor=white)](https://xceed.com/blog/)
[![Support](https://img.shields.io/badge/Support-FF6F00?style=for-the-badge&logo=googlemessages&logoColor=white)](https://xceed.com/support/)
[![License](https://img.shields.io/badge/License-FF6F00?style=for-the-badge&logo=rocket&logoColor=white)](https://xceed.com/licensing/)

---

**Xceed Words for .NET** – Enterprise Word processing and PDF manipulation library for .NET applications.