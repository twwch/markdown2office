# Markdown2Office

A Java SDK for converting Markdown documents to various office formats including Word, Excel, PDF, and more.

## Features

- Convert Markdown to multiple formats:
  - Word (DOCX)
  - Excel (XLSX)
  - PDF
  - Plain Text (TXT)
  - Markdown (MD)
- Preserve Markdown structure and formatting
- Support for tables, lists, code blocks, and more
- Easy-to-use API
- Command-line interface

## Installation

### Maven

```xml
<dependency>
    <groupId>io.github.twwch</groupId>
    <artifactId>markdown2office</artifactId>
    <version>1.0.16</version>
</dependency>
```

### Gradle

```gradle
implementation 'io.github.twwch:markdown2office:1.0.16'
```

## Usage

### Basic Usage

```java
import io.github.twwch.markdown2office.Markdown2Office;
import io.github.twwch.markdown2office.model.FileType;

public class Example {
    public static void main(String[] args) throws IOException {
        Markdown2Office converter = new Markdown2Office();
        
        String markdown = "# Hello World\n\nThis is **bold** text.";
        
        // Convert to Word
        converter.convert(markdown, FileType.WORD, "output.docx");
        
        // Convert to PDF
        converter.convert(markdown, FileType.PDF, "output.pdf");
        
        // Convert to Excel
        converter.convert(markdown, FileType.EXCEL, "output.xlsx");
    }
}
```

### Convert Files

```java
// Convert markdown file to Word
converter.convertFile("input.md", FileType.WORD, "output.docx");

// Auto-detect output format from file extension
converter.convertFile("input.md", "output.pdf");
```

### Stream Output

```java
// Write to OutputStream
try (FileOutputStream fos = new FileOutputStream("output.docx")) {
    converter.convert(markdown, FileType.WORD, fos);
}

// Get as byte array
byte[] pdfBytes = converter.convertToBytes(markdown, FileType.PDF);
```

### Command Line

```bash
java -jar markdown2office.jar input.md output.docx
```

## Supported Markdown Features

- **Headings** (H1-H6)
- **Text formatting** (bold, italic, code)
- **Lists** (ordered, unordered, nested)
- **Blockquotes**
- **Code blocks** (with syntax highlighting)
- **Tables** (GitHub Flavored Markdown)
- **Links**
- **Images** (as references)
- **Horizontal rules**
- **Task lists**

## Building from Source

```bash
# Clone the repository
git clone https://github.com/twwch/markdown2office.git
cd markdown2office

# Build with Maven
mvn clean install

# Run tests
mvn test
```

## Configuration for Maven Central Release

To release to Maven Central, you need to configure the following GitHub Secrets:

1. **GPG_PRIVATE_KEY**: Your GPG private key for signing artifacts
2. **GPG_PASSPHRASE**: Passphrase for your GPG key
3. **MAVEN_USERNAME**: Your Sonatype JIRA username
4. **MAVEN_PASSWORD**: Your Sonatype JIRA password

### Creating a Release

1. Create and push a tag:
```bash
git tag v1.0.16
git push origin v1.0.16
```

2. The GitHub Action will automatically:
   - Build and test the project
   - Sign the artifacts with GPG
   - Deploy to Maven Central
   - Create a GitHub Release

## Requirements

- Java 8 or higher
- Maven 3.6 or higher (for building)

## Dependencies

- Apache POI (Word and Excel support)
- iText (PDF generation)
- CommonMark (Markdown parsing)
- SLF4J (Logging)

## License

Apache License 2.0

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## File Parsing (Enhanced in v1.0.16)

### Overview

The library includes a powerful and enhanced file parsing system that can extract content, metadata, and structure from various document formats:

- **PDF** (.pdf) - With hidden layer filtering
- **Word** (.doc, .docx) - Enhanced structure extraction
- **Excel** (.xls, .xlsx) - Improved multi-sheet handling
- **PowerPoint** (.ppt, .pptx)
- **Text** (.txt) - Smart encoding detection
- **Markdown** (.md) - Full CommonMark support
- **CSV** (.csv) - Advanced encoding detection (UTF-8, GBK, GB2312, etc.)
- And 20+ other formats via Apache Tika

### Basic Usage

```java
import io.github.twwch.markdown2office.parser.UniversalFileParser;
import io.github.twwch.markdown2office.parser.ParsedDocument;

UniversalFileParser parser = new UniversalFileParser();

// Parse any supported file
ParsedDocument document = parser.parse(new File("document.pdf"));

// Get content as markdown
String markdown = document.toMarkdown();

// Access document metadata
DocumentMetadata metadata = document.getDocumentMetadata();
System.out.println("Total pages: " + metadata.getTotalPages());
System.out.println("Word count: " + metadata.getTotalWords());

// Access page-by-page content
for (PageContent page : document.getPages()) {
    System.out.println("Page " + page.getPageNumber() + ":");
    System.out.println("  Words: " + page.getWordCount());
    System.out.println("  Headings: " + page.getHeadings());
    System.out.println("  Tables: " + page.getTables().size());
}
```

### ParsedDocument Fields

| Field | Type | Description |
|-------|------|-------------|
| `content` | String | Raw text content of the entire document |
| `markdownContent` | String | Content converted to Markdown format |
| `fileType` | FileType | Type of the parsed file (PDF, WORD, EXCEL, etc.) |
| `pages` | List<PageContent> | Page-by-page content and structure |
| `tables` | List<ParsedTable> | All tables found in the document |
| `documentMetadata` | DocumentMetadata | Comprehensive metadata about the document |
| `metadata` | Map<String,String> | Legacy metadata for backward compatibility |

### DocumentMetadata Fields

| Field | Type | Description |
|-------|------|-------------|
| `fileName` | String | Name of the source file |
| `fileSize` | Long | Size of the file in bytes |
| `fileType` | FileType | Document type |
| `title` | String | Document title (if available) |
| `author` | String | Document author |
| `subject` | String | Document subject |
| `keywords` | String | Document keywords |
| `creationDate` | Date | When the document was created |
| `modificationDate` | Date | Last modification date |
| `totalPages` | Integer | Total number of pages |
| `totalWords` | Integer | Total word count |
| `totalCharacters` | Integer | Total character count |
| `totalParagraphs` | Integer | Total paragraph count |
| `totalTables` | Integer | Total table count |
| `totalSheets` | Integer | For Excel: number of sheets |
| `totalSlides` | Integer | For PowerPoint: number of slides |

### PageContent Fields

| Field | Type | Description |
|-------|------|-------------|
| `pageNumber` | int | Page number (starting from 1) |
| `rawText` | String | Raw text content of the page |
| `markdownContent` | String | Page content in Markdown format |
| `headings` | List<String> | All headings found on the page |
| `paragraphs` | List<String> | All paragraphs on the page |
| `lists` | List<String> | All list items on the page |
| `tables` | List<ParsedTable> | Tables found on this page |
| `wordCount` | Integer | Word count for this page |
| `characterCount` | Integer | Character count for this page |

### ParsedTable Fields

| Field | Type | Description |
|-------|------|-------------|
| `title` | String | Table title or caption |
| `headers` | List<String> | Column headers |
| `data` | List<List<String>> | Table data rows |
| `rowCount` | int | Number of data rows |
| `columnCount` | int | Number of columns |

### Advanced Examples

#### Parsing with Specific Parser

```java
import io.github.twwch.markdown2office.parser.impl.*;

// Use specific parser for better control
PdfFileParser pdfParser = new PdfFileParser();
ParsedDocument pdfDoc = pdfParser.parse("document.pdf");

WordFileParser wordParser = new WordFileParser();
ParsedDocument wordDoc = wordParser.parse("document.docx");

ExcelFileParser excelParser = new ExcelFileParser();
ParsedDocument excelDoc = excelParser.parse("spreadsheet.xlsx");
```

#### Enhanced Features in v1.0.16

##### PDF Hidden Layer Filtering

The PDF parser now supports filtering out hidden layers, watermarks, and invisible text that may appear in some PDFs (e.g., from recruitment platforms like BOSS Zhipin).

```java
import io.github.twwch.markdown2office.parser.impl.PdfFileParser;
import io.github.twwch.markdown2office.parser.ParsedDocument;

// Default behavior: hidden layers are excluded
PdfFileParser parser = new PdfFileParser();
ParsedDocument doc = parser.parse("resume.pdf");
// Hidden watermarks and invisible text are automatically filtered out

// If you need to include hidden layers (not recommended for most cases)
PdfFileParser parserWithHidden = new PdfFileParser(true);
ParsedDocument docWithHidden = parserWithHidden.parse("document.pdf");

// Or configure dynamically
PdfFileParser configurable = new PdfFileParser();
configurable.setIncludeHiddenLayers(false); // Exclude hidden content (default)
ParsedDocument cleanDoc = configurable.parse("document.pdf");

configurable.setIncludeHiddenLayers(true); // Include everything
ParsedDocument fullDoc = configurable.parse("document.pdf");
```

**What gets filtered when `includeHiddenLayers` is `false` (default):**
- Invisible text layers (rendering mode NEITHER)
- Text with transparency below 30%
- White or nearly white text on white backgrounds
- Hidden annotations and watermarks
- XObjects with opacity below 50%
- Resources marked as watermarks or backgrounds

**Common use cases:**
- **Resume parsing**: Remove recruitment platform watermarks
- **Document cleaning**: Extract only visible content
- **Content migration**: Get clean text without metadata artifacts
- **Text analysis**: Focus on actual document content

#### Extract Tables from Documents

```java
ParsedDocument document = parser.parse(new File("report.pdf"));

// Get all tables
for (ParsedTable table : document.getTables()) {
    System.out.println("Table: " + table.getTitle());
    System.out.println("Headers: " + table.getHeaders());
    
    // Convert table to markdown
    String tableMarkdown = table.toMarkdown();
    System.out.println(tableMarkdown);
    
    // Access table data
    for (List<String> row : table.getData()) {
        System.out.println(String.join(" | ", row));
    }
}
```

##### CSV Encoding Detection (New)

The CSV parser now automatically detects and handles various character encodings, including Chinese encodings:

```java
import io.github.twwch.markdown2office.parser.impl.CsvFileParser;

// Automatically detects encoding (UTF-8, GBK, GB2312, etc.)
CsvFileParser csvParser = new CsvFileParser();
ParsedDocument doc = csvParser.parse("chinese_data.csv");

// The parser will:
// 1. Detect BOM markers
// 2. Try common encodings in order
// 3. Validate content to ensure correct encoding
// 4. Fall back gracefully if encoding cannot be determined

String content = doc.getContent(); // Correctly decoded content
```

#### Working with Excel Files

```java
ParsedDocument excel = parser.parse(new File("data.xlsx"));
DocumentMetadata metadata = excel.getDocumentMetadata();

System.out.println("Total sheets: " + metadata.getTotalSheets());

// Each sheet is treated as a page
for (PageContent sheet : excel.getPages()) {
    System.out.println("Sheet " + sheet.getPageNumber());
    
    // Excel sheets typically contain one table per sheet
    for (ParsedTable table : sheet.getTables()) {
        System.out.println("  Rows: " + table.getRowCount());
        System.out.println("  Columns: " + table.getColumnCount());
    }
}
```

#### Search and Analysis

```java
ParsedDocument document = parser.parse(new File("manual.pdf"));

// Search for specific content
for (PageContent page : document.getPages()) {
    if (page.getRawText().contains("installation")) {
        System.out.println("Found 'installation' on page " + page.getPageNumber());
    }
    
    // Find pages with tables
    if (!page.getTables().isEmpty()) {
        System.out.println("Page " + page.getPageNumber() + " has " + 
                         page.getTables().size() + " table(s)");
    }
    
    // Find pages with specific headings
    for (String heading : page.getHeadings()) {
        if (heading.toLowerCase().contains("introduction")) {
            System.out.println("Introduction section on page " + page.getPageNumber());
        }
    }
}
```

#### Convert Parsed Document Back to Office Format

```java
// Parse a PDF and convert to Word
ParsedDocument pdfDoc = parser.parse(new File("report.pdf"));
String markdown = pdfDoc.toMarkdown();

Markdown2Office converter = new Markdown2Office();
converter.convert(markdown, FileType.WORD, "report.docx");

// Parse Word and convert to PDF with formatting preserved
ParsedDocument wordDoc = parser.parse(new File("document.docx"));
converter.convert(wordDoc.toMarkdown(), FileType.PDF, "document.pdf");
```

### Supported Features by Format

| Format | Text Extraction | Table Parsing | Metadata | Structure Detection | Special Features (v1.0.16) |
|--------|----------------|---------------|----------|--------------------|--------------------------|
| **PDF** | ✅ Full text with formatting | ✅ Complex tables | ✅ Complete | ✅ Pages, headings, paragraphs | • Hidden layer filtering<br>• Watermark removal<br>• Invisible text detection |
| **Word (.docx/.doc)** | ✅ Rich text preservation | ✅ Nested tables | ✅ Complete | ✅ Sections, headings, lists | • Style preservation<br>• Comment extraction<br>• Track changes support |
| **Excel (.xlsx/.xls)** | ✅ Cell values & formulas | ✅ Native | ✅ Complete | ✅ Sheets as pages | • Formula evaluation<br>• Merged cell handling<br>• Multiple sheet support |
| **PowerPoint (.pptx/.ppt)** | ✅ Slide content | ✅ Slide tables | ✅ Complete | ✅ Slides as pages | • Speaker notes<br>• Slide layout detection<br>• Shape text extraction |
| **CSV** | ✅ Full content | ✅ Native | ⚠️ Limited | ❌ | • **Auto encoding detection (UTF-8, GBK, GB2312)**<br>• BOM handling<br>• Delimiter detection |
| **Text (.txt)** | ✅ Plain text | ❌ | ⚠️ File info only | ❌ | • Encoding auto-detection<br>• Line break preservation |
| **Markdown (.md)** | ✅ Formatted text | ✅ GFM tables | ❌ | ✅ Heading hierarchy | • CommonMark + GFM<br>• Code block preservation<br>• Task list support |
| **RTF** | ✅ Rich text | ✅ | ⚠️ Basic | ⚠️ Partial | • Via Apache Tika |
| **HTML** | ✅ Text content | ✅ | ⚠️ Basic | ✅ DOM structure | • Tag removal<br>• Link preservation |
| **XML** | ✅ Text nodes | ⚠️ Structured data | ⚠️ Basic | ✅ Element hierarchy | • Via Apache Tika |
| **Other formats** | ✅ Via Tika | ⚠️ Format dependent | ⚠️ Basic | ⚠️ Limited | • 20+ formats via Apache Tika |

**Legend:**
- ✅ Full support with comprehensive features
- ⚠️ Partial support or limited functionality
- ❌ Not supported
- **Bold** items indicate recent enhancements in v1.0.16

### Performance Considerations

- Large files are processed efficiently with streaming where possible
- Page-based extraction allows processing documents without loading entire content into memory
- Metadata is extracted without parsing full document content when possible
- Smart encoding detection minimizes re-reading of files
- Optimized table extraction for large Excel and CSV files

### Error Handling

```java
try {
    ParsedDocument document = parser.parse(new File("document.pdf"));
    // Process document
} catch (IOException e) {
    System.err.println("Failed to parse document: " + e.getMessage());
} catch (UnsupportedFileException e) {
    System.err.println("File format not supported: " + e.getMessage());
}
```

## Support

If you encounter any issues or have questions, please file an issue on the [GitHub repository](https://github.com/twwch/markdown2office/issues).
