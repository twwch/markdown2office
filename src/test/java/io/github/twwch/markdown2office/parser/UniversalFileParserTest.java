package io.github.twwch.markdown2office.parser;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Test class for UniversalFileParser
 */
public class UniversalFileParserTest {
    
    private UniversalFileParser parser;
    
    @BeforeEach
    void setUp() {
        parser = new UniversalFileParser();
    }
    
    @Test
    void testSupportedExtensions() {
        String[] extensions = parser.getSupportedExtensions();
        assertNotNull(extensions);
        assertTrue(extensions.length > 0);
        
        // Check for expected extensions
        boolean hasPdf = false, hasDocx = false, hasMarkdown = false;
        for (String ext : extensions) {
            if (ext.equals(".pdf")) hasPdf = true;
            if (ext.equals(".docx")) hasDocx = true;
            if (ext.equals(".md")) hasMarkdown = true;
        }
        assertTrue(hasPdf, "Should support PDF files");
        assertTrue(hasDocx, "Should support DOCX files");
        assertTrue(hasMarkdown, "Should support Markdown files");
    }
    
    @Test
    void testSupportsMethod() {
        // Test supported formats
        assertTrue(parser.supports("test.pdf"));
        assertTrue(parser.supports("test.docx"));
        assertTrue(parser.supports("test.doc"));
        assertTrue(parser.supports("test.xlsx"));
        assertTrue(parser.supports("test.xls"));
        assertTrue(parser.supports("test.pptx"));
        assertTrue(parser.supports("test.ppt"));
        assertTrue(parser.supports("test.csv"));
        assertTrue(parser.supports("test.txt"));
        assertTrue(parser.supports("test.md"));
        assertTrue(parser.supports("test.markdown"));
        assertTrue(parser.supports("test.html"));
        assertTrue(parser.supports("test.xml"));
        assertTrue(parser.supports("test.rtf"));
        
        // Test unsupported formats
        assertFalse(parser.supports("test.xyz"));
        assertFalse(parser.supports("test.unknown"));
    }
    
    @Test
    void testParseMarkdownFromStream() throws IOException {
        String markdownContent = "# Test Document\n\nThis is a **bold** text with *italic* formatting.\n\n## Section 2\n\n- Item 1\n- Item 2\n- Item 3";
        ByteArrayInputStream stream = new ByteArrayInputStream(markdownContent.getBytes(StandardCharsets.UTF_8));
        
        ParsedDocument result = parser.parse(stream, "test.md");
        
        assertNotNull(result);
        assertEquals(ParsedDocument.FileType.MARKDOWN, result.getFileType());
        assertNotNull(result.getContent());
        assertNotNull(result.getMarkdownContent());
        assertEquals("Test Document", result.getTitle());
        assertTrue(result.getMetadata().containsKey("Parser Used"));
        assertEquals("MarkdownFileParser", result.getMetadata().get("Parser Used"));
    }
    
    @Test
    void testParseTextFromStream() throws IOException {
        String textContent = "Test Document\n\nThis is plain text content.\nLine 2\nLine 3";
        ByteArrayInputStream stream = new ByteArrayInputStream(textContent.getBytes(StandardCharsets.UTF_8));
        
        ParsedDocument result = parser.parse(stream, "test.txt");
        
        assertNotNull(result);
        assertEquals(ParsedDocument.FileType.TEXT, result.getFileType());
        assertNotNull(result.getContent());
        assertNotNull(result.getMarkdownContent());
        assertTrue(result.getMetadata().containsKey("Parser Used"));
        assertEquals("TextFileParser", result.getMetadata().get("Parser Used"));
    }
    
    @Test
    void testParseCsvFromStream() throws IOException {
        String csvContent = "Name,Age,City\nJohn,30,New York\nJane,25,Los Angeles\nBob,35,Chicago";
        ByteArrayInputStream stream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));
        
        ParsedDocument result = parser.parse(stream, "test.csv");
        
        assertNotNull(result);
        assertEquals(ParsedDocument.FileType.CSV, result.getFileType());
        assertNotNull(result.getContent());
        assertNotNull(result.getMarkdownContent());
        assertEquals(1, result.getTables().size());
        
        ParsedDocument.ParsedTable table = result.getTables().get(0);
        assertEquals(3, table.getHeaders().size());
        assertEquals("Name", table.getHeaders().get(0));
        assertEquals("Age", table.getHeaders().get(1));
        assertEquals("City", table.getHeaders().get(2));
        assertEquals(3, table.getData().size());
        
        assertTrue(result.getMetadata().containsKey("Parser Used"));
        assertEquals("CsvFileParser", result.getMetadata().get("Parser Used"));
    }
    
    @Test
    void testParseWithOptionsFailSilently() throws IOException {
        // Create a temporary file for testing
        java.io.File tempFile = java.io.File.createTempFile("test", ".unsupported");
        tempFile.deleteOnExit();
        java.nio.file.Files.write(tempFile.toPath(), "test content".getBytes());
        
        // Test with unsupported file type - should return null when failSilently is true
        ParsedDocument result = parser.parseWithOptions(tempFile, true);
        assertNull(result);
        
        // Test with unsupported file type - should throw exception when failSilently is false
        assertThrows(UnsupportedOperationException.class, () -> {
            parser.parseWithOptions(tempFile, false);
        });
    }
    
    @Test
    void testGetParserInfo() {
        String info = parser.getParserInfo();
        assertNotNull(info);
        assertTrue(info.contains("PdfFileParser"));
        assertTrue(info.contains("WordFileParser"));
        assertTrue(info.contains("ExcelFileParser"));
        assertTrue(info.contains("PowerPointFileParser"));
        assertTrue(info.contains("CsvFileParser"));
        assertTrue(info.contains("MarkdownFileParser"));
        assertTrue(info.contains("TextFileParser"));
        assertTrue(info.contains("TikaFileParser"));
    }
    
    @Test
    void testInvalidInputs() {
        // Test null inputs
        assertThrows(IllegalArgumentException.class, () -> parser.parse((String) null));
        assertThrows(IllegalArgumentException.class, () -> parser.parse(""));
        assertThrows(IllegalArgumentException.class, () -> parser.parse("   "));
        
        assertThrows(IllegalArgumentException.class, () -> parser.parse(null, "test.txt"));
        assertThrows(IllegalArgumentException.class, () -> {
            ByteArrayInputStream stream = new ByteArrayInputStream("test".getBytes());
            parser.parse(stream, null);
        });
        assertThrows(IllegalArgumentException.class, () -> {
            ByteArrayInputStream stream = new ByteArrayInputStream("test".getBytes());
            parser.parse(stream, "");
        });
    }
    
    @Test
    void testNonExistentFile() {
        assertThrows(IOException.class, () -> parser.parse("/non/existent/file.txt"));
    }
}