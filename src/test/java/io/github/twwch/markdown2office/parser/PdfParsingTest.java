package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.Markdown2Office;
import io.github.twwch.markdown2office.model.FileType;
import io.github.twwch.markdown2office.parser.impl.PdfFileParser;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Test class for PDF parsing functionality
 */
public class PdfParsingTest {
    
    private UniversalFileParser universalParser;
    private PdfFileParser pdfParser;
    private Markdown2Office converter;
    
    @BeforeEach
    public void setUp() {
        universalParser = new UniversalFileParser();
        pdfParser = new PdfFileParser();
        converter = new Markdown2Office();
    }
    
    @Test
    public void testParsePdfFile(@TempDir Path tempDir) throws IOException {
        // Step 1: Create a PDF file with known content
        String originalMarkdown = "# Test Document 测试文档\n\n" +
                "## Introduction 介绍\n\n" +
                "This is a test document for PDF parsing.\n" +
                "这是一个用于PDF解析的测试文档。\n\n" +
                "### Features 特性\n\n" +
                "- Feature 1 特性1\n" +
                "- Feature 2 特性2\n" +
                "- Feature 3 特性3\n\n" +
                "**Bold text 粗体文本** and *italic text 斜体文本*.\n\n" +
                "| Column 1 | Column 2 | Column 3 |\n" +
                "|----------|----------|----------|\n" +
                "| Data 1   | Data 2   | Data 3   |\n" +
                "| 数据 A   | 数据 B   | 数据 C   |\n\n" +
                "```java\n" +
                "public class Test {\n" +
                "    public static void main(String[] args) {\n" +
                "        System.out.println(\"Hello World!\");\n" +
                "    }\n" +
                "}\n" +
                "```\n";
        
        // Create PDF file
        File pdfFile = tempDir.resolve("output.pdf").toFile();
        converter.convert(originalMarkdown, FileType.PDF, pdfFile);
        
        assertTrue(pdfFile.exists(), "PDF file should be created");
        assertTrue(pdfFile.length() > 0, "PDF file should not be empty");
        
        // Step 2: Parse the PDF file using UniversalFileParser
        ParsedDocument parsedDoc = universalParser.parse(pdfFile);
        
        assertNotNull(parsedDoc, "Parsed document should not be null");
        assertEquals(ParsedDocument.FileType.PDF, parsedDoc.getFileType(), "File type should be PDF");
        
        // Step 3: Verify content extraction
        String content = parsedDoc.getContent();
        assertNotNull(content, "Content should not be null");
        assertFalse(content.isEmpty(), "Content should not be empty");
        
        // Check for key content (PDF text extraction may not preserve exact formatting)
        assertTrue(content.contains("Test Document") || content.contains("测试文档"), 
                "Content should contain title");
        assertTrue(content.contains("Introduction") || content.contains("介绍"), 
                "Content should contain introduction");
        assertTrue(content.contains("Feature") || content.contains("特性"), 
                "Content should contain features");
        assertTrue(content.contains("Hello World"), 
                "Content should contain code content");
        
        // Step 4: Check markdown conversion
        String markdownContent = parsedDoc.getMarkdownContent();
        assertNotNull(markdownContent, "Markdown content should not be null");
        assertFalse(markdownContent.isEmpty(), "Markdown content should not be empty");
        
        // Step 5: Check metadata
        assertNotNull(parsedDoc.getMetadata(), "Metadata should not be null");
        assertTrue(parsedDoc.getMetadata().containsKey("Page Count"), "Should have page count metadata");
        
        System.out.println("PDF Parsing Test Results:");
        System.out.println("-------------------------");
        System.out.println("Original markdown length: " + originalMarkdown.length());
        System.out.println("PDF file size: " + pdfFile.length() + " bytes");
        System.out.println("Extracted content length: " + content.length());
        System.out.println("Page count: " + parsedDoc.getMetadata().get("Page Count"));
        System.out.println("\nExtracted content preview (first 500 chars):");
        System.out.println(content.substring(0, Math.min(content.length(), 500)));
    }
    
    @Test
    public void testParsePdfInputStream(@TempDir Path tempDir) throws IOException {
        // Create a PDF file
        String markdown = "# Stream Test\n\nThis is a test for stream parsing.";
        File pdfFile = tempDir.resolve("stream_test.pdf").toFile();
        converter.convert(markdown, FileType.PDF, pdfFile);
        
        // Parse using input stream
        byte[] pdfBytes = Files.readAllBytes(pdfFile.toPath());
        ParsedDocument parsedDoc = universalParser.parse(
            new java.io.ByteArrayInputStream(pdfBytes), 
            "stream_test.pdf"
        );
        
        assertNotNull(parsedDoc, "Parsed document from stream should not be null");
        assertEquals(ParsedDocument.FileType.PDF, parsedDoc.getFileType());
        assertTrue(parsedDoc.getContent().contains("Stream Test"));
    }
    
    @Test
    public void testParsePdfWithMetadata(@TempDir Path tempDir) throws IOException {
        // Create a more complex PDF with metadata
        String markdown = "# Document with Metadata\n\n" +
                "**Author:** Test Author\n\n" +
                "**Date:** 2024-01-01\n\n" +
                "## Content\n\n" +
                "This document contains various metadata fields.";
        
        File pdfFile = tempDir.resolve("metadata_test.pdf").toFile();
        converter.convert(markdown, FileType.PDF, pdfFile);
        
        // Parse and check metadata
        ParsedDocument parsedDoc = pdfParser.parse(pdfFile);
        
        assertNotNull(parsedDoc);
        assertNotNull(parsedDoc.getMetadata());
        
        // PDF metadata might include creation info
        if (parsedDoc.getMetadata().containsKey("Producer")) {
            System.out.println("PDF Producer: " + parsedDoc.getMetadata().get("Producer"));
        }
        if (parsedDoc.getMetadata().containsKey("Creator")) {
            System.out.println("PDF Creator: " + parsedDoc.getMetadata().get("Creator"));
        }
    }
    
    @Test
    public void testParsePdfWithChineseContent(@TempDir Path tempDir) throws IOException {
        // Create PDF with Chinese content
        String chineseMarkdown = "# 中文标题\n\n" +
                "## 第一章\n\n" +
                "这是中文内容的测试。\n\n" +
                "### 1.1 子章节\n\n" +
                "包含中英文混合的content。\n\n" +
                "- 列表项1\n" +
                "- 列表项2\n" +
                "- 列表项3\n";
        
        File pdfFile = tempDir.resolve("chinese_test.pdf").toFile();
        converter.convert(chineseMarkdown, FileType.PDF, pdfFile);
        
        // Parse and verify Chinese content
        ParsedDocument parsedDoc = universalParser.parse(pdfFile);
        
        assertNotNull(parsedDoc);
        String content = parsedDoc.getContent();
        assertNotNull(content);
        
        // Check for Chinese characters
        assertTrue(content.contains("中文") || content.contains("第一章") || 
                  content.contains("列表项"), "Should contain Chinese content");
        
        System.out.println("\nChinese PDF content extracted successfully");
        System.out.println("Content length: " + content.length());
    }
    
    @Test
    public void testParseNonExistentPdf() {
        // Test error handling for non-existent file
        File nonExistentFile = new File("non_existent.pdf");
        
        // Should throw exception for non-existent file
        assertThrows(IOException.class, () -> {
            universalParser.parse(nonExistentFile);
        });
    }
    
    @Test
    public void testRoundTripConversion(@TempDir Path tempDir) throws IOException {
        // Test: Markdown -> PDF -> Parse -> Markdown
        String originalMarkdown = "# Round Trip Test\n\n" +
                "## Section 1\n\n" +
                "This is a test of round-trip conversion.\n\n" +
                "### Subsection\n\n" +
                "- Item 1\n" +
                "- Item 2\n\n" +
                "**Bold** and *italic* text.\n";
        
        // Step 1: Convert markdown to PDF
        File pdfFile = tempDir.resolve("roundtrip.pdf").toFile();
        converter.convert(originalMarkdown, FileType.PDF, pdfFile);
        
        // Step 2: Parse PDF back
        ParsedDocument parsedDoc = universalParser.parse(pdfFile);
        
        // Step 3: Get markdown representation
        String extractedMarkdown = parsedDoc.toMarkdown();
        
        assertNotNull(extractedMarkdown);
        assertFalse(extractedMarkdown.isEmpty());
        
        // The extracted markdown won't be identical due to PDF text extraction limitations,
        // but should contain the main content
        assertTrue(extractedMarkdown.contains("Round Trip Test") || 
                  parsedDoc.getContent().contains("Round Trip Test"));
        
        System.out.println("\nRound trip test:");
        System.out.println("Original markdown length: " + originalMarkdown.length());
        System.out.println("Extracted markdown length: " + extractedMarkdown.length());
    }
}