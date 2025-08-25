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
    <version>1.0.0</version>
</dependency>
```

### Gradle

```gradle
implementation 'io.github.twwch:markdown2office:1.0.0'
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
git tag v1.0.0
git push origin v1.0.0
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

## Support

If you encounter any issues or have questions, please file an issue on the [GitHub repository](https://github.com/twwch/markdown2office/issues).# markdown2office
