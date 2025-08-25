package io.github.twwch.markdown2office;

import io.github.twwch.markdown2office.model.FileType;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;

public class Markdown2OfficeTest {
    
    private Markdown2Office converter;
    private String sampleMarkdown;
    
    @BeforeEach
    public void setUp() {
        converter = new Markdown2Office();
        sampleMarkdown = "# Heading 1\n\n" +
                "This is a **bold** text and this is *italic* text.\n\n" +
                "## Heading 2\n\n" +
                "- Item 1\n" +
                "- Item 2\n" +
                "  - Subitem 2.1\n" +
                "  - Subitem 2.2\n\n" +
                "### Heading 3\n\n" +
                "1. First\n" +
                "2. Second\n" +
                "3. Third\n\n" +
                "> This is a blockquote\n\n" +
                "```java\n" +
                "public class Test {\n" +
                "    public static void main(String[] args) {\n" +
                "        System.out.println(\"Hello, World!\");\n" +
                "    }\n" +
                "}\n" +
                "```\n\n" +
                "| Header 1 | Header 2 | Header 3 |\n" +
                "|----------|----------|----------|\n" +
                "| Cell 1   | Cell 2   | Cell 3   |\n" +
                "| Cell 4   | Cell 5   | Cell 6   |\n\n" +
                "[Link text](https://example.com)";
    }
    
    @Test
    public void testConvertToWord() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.WORD, outputStream);
        
        byte[] result = outputStream.toByteArray();
        assertNotNull(result);
        // 写入文件
        Files.write(new File("output.docx").toPath(), result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testConvertToExcel() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.EXCEL, outputStream);
        
        byte[] result = outputStream.toByteArray();
        assertNotNull(result);

        Files.write(new File("output.xlsx").toPath(), result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testConvertToPdf() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.PDF, outputStream);
        
        byte[] result = outputStream.toByteArray();

        Files.write(new File("output.pdf").toPath(), result);
        assertNotNull(result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testConvertToText() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.TEXT, outputStream);
        
        String result = outputStream.toString("UTF-8");
        assertNotNull(result);
        Files.write(new File("output.txt").toPath(), result.getBytes());
        assertTrue(result.contains("HEADING"));
        assertTrue(result.contains("**bold**"));
        assertTrue(result.contains("*italic*"));
    }
    
    @Test
    public void testConvertToMarkdown() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.MARKDOWN, outputStream);
        
        String result = outputStream.toString("UTF-8");
        Files.write(new File("output.md").toPath(), result.getBytes());
        assertEquals(sampleMarkdown, result);
    }
    
    @Test
    public void testConvertToFile(@TempDir Path tempDir) throws IOException {
        File outputFile = tempDir.resolve("test.docx").toFile();
        converter.convert(sampleMarkdown, FileType.WORD, outputFile);
        
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertFileToFile(@TempDir Path tempDir) throws IOException {
        Path inputFile = tempDir.resolve("input.md");
        Files.write(inputFile, sampleMarkdown.getBytes());
        
        File outputFile = tempDir.resolve("output.pdf").toFile();
        converter.convertFile(inputFile.toString(), FileType.PDF, outputFile.toString());
        
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertWithAutoDetection(@TempDir Path tempDir) throws IOException {
        Path inputFile = tempDir.resolve("input.md");
        Files.write(inputFile, sampleMarkdown.getBytes());
        
        String outputPath = tempDir.resolve("output.xlsx").toString();
        converter.convertFile(inputFile.toString(), outputPath);
        
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertToBytes() throws IOException {
        byte[] result = converter.convertToBytes(sampleMarkdown, FileType.WORD);
        
        assertNotNull(result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testNullMarkdownThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert(null, FileType.WORD, new ByteArrayOutputStream());
        });
    }
    
    @Test
    public void testEmptyMarkdownThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert("  ", FileType.WORD, new ByteArrayOutputStream());
        });
    }
    
    @Test
    public void testNullFileTypeThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert(sampleMarkdown, null, new ByteArrayOutputStream());
        });
    }
    
    @Test
    public void testNullOutputStreamThrowsException() {
        ByteArrayOutputStream nullStream = null;
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert(sampleMarkdown, FileType.WORD, nullStream);
        });
    }
    
    @Test
    public void testInvalidFileExtensionThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            FileType.fromExtension("invalid");
        });
    }
}