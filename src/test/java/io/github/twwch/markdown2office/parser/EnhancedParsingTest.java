package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.ExcelFileParser;
import io.github.twwch.markdown2office.parser.impl.PdfFileParser;
import io.github.twwch.markdown2office.parser.impl.PowerPointFileParser;
import io.github.twwch.markdown2office.parser.impl.WordFileParser;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Test class for enhanced parsing features with page-based content and metadata
 * Tests parsing of existing demo.pdf file
 */
public class EnhancedParsingTest {
    
    private UniversalFileParser universalParser;
    private static final String DEMO_PDF_PATH = "demo.pdf";
    
    @BeforeEach
    public void setUp() {
        universalParser = new UniversalFileParser();
    }
    
    @Test
    public void testPdfPageBasedExtraction() throws IOException {
        // Read existing demo.pdf file
        File pdfFile = new File(DEMO_PDF_PATH);
        
        // Skip test if demo.pdf doesn't exist
        if (!pdfFile.exists()) {
            System.out.println("Warning: demo.pdf not found at " + pdfFile.getAbsolutePath());
            System.out.println("Skipping PDF parsing test. Please create demo.pdf to test.");
            return;
        }
        
        System.out.println("=== Reading demo.pdf from: " + pdfFile.getAbsolutePath() + " ===");
//        PdfFileParser parser= new PdfFileParser(false);
        // Parse the PDF
        ParsedDocument document = universalParser.parse(pdfFile);
        
        // Check metadata
        assertNotNull(document.getDocumentMetadata());
        DocumentMetadata metadata = document.getDocumentMetadata();
        assertNotNull(metadata.getTotalPages());
        assertTrue(metadata.getTotalPages() >= 1);
        assertNotNull(metadata.getFileSize());
        assertTrue(metadata.getFileSize() > 0);
        assertEquals(ParsedDocument.FileType.PDF, metadata.getFileType());
        
        // Check page-based content
        assertNotNull(document.getPages());
        List<PageContent> pages = document.getPages();
        assertFalse(pages.isEmpty());
        
        // Check first page
        PageContent firstPage = pages.get(0);
        assertEquals(1, firstPage.getPageNumber());
        assertNotNull(firstPage.getRawText());
        assertTrue(firstPage.hasContent());
        
        // Print detailed information
        System.out.println("=== PDF Parsing Results ===");
        System.out.println(metadata.toString());
        System.out.println("Total pages extracted: " + pages.size());
        
        for (PageContent page : pages) {
            System.out.println("\n--- Page " + page.getPageNumber() + " ---");
            System.out.println("Word count: " + page.getWordCount());
            System.out.println("Character count: " + page.getCharacterCount());
            System.out.println("Content: " + page.getMarkdownContent());
        }
    }
    
    @Test
    public void testWordDocumentMetadata() throws IOException {
        // Test Word parser with an existing file if available
        File wordFile = new File("demo.docx");
        
        if (!wordFile.exists()) {
            System.out.println("Warning: demo.docx not found, skipping Word test");
            // Test that parser supports Word files
            WordFileParser parser = new WordFileParser();
            assertTrue(parser.supports("test.docx"));
            assertTrue(parser.supports("test.doc"));
            return;
        }
        
        // Parse the Word document
        ParsedDocument document = new WordFileParser().parse(wordFile);
        
        // Check metadata
        assertNotNull(document.getDocumentMetadata());
        DocumentMetadata metadata = document.getDocumentMetadata();
        assertEquals(ParsedDocument.FileType.WORD, metadata.getFileType());
        assertNotNull(metadata.getFileSize());
        
        // Check content structure
        assertNotNull(document.getPages());
        assertFalse(document.getPages().isEmpty());
        
        // Check markdown conversion preserves formatting
        String markdownOutput = document.toMarkdown();
        assertNotNull(markdownOutput);
        assertTrue(markdownOutput.contains("# ") || markdownOutput.contains("## "));
        
        System.out.println("\n=== Word Document Metadata ===");
        System.out.println(metadata.toString());
        System.out.println("Total word count: " + document.getTotalWordCount());
        System.out.println("Total pages: " + document.getPages().size());
        
        // Show markdown output sample
        String markdown = document.toMarkdown();
        System.out.println("\n=== Markdown Output (first 500 chars) ===");
        System.out.println(markdown.substring(0, Math.min(500, markdown.length())));
        
        List<PageContent> pages = document.getPages();
        for (PageContent page : pages) {
            System.out.println("\n--- Page " + page.getPageNumber() + " ---");
            System.out.println("Word count: " + page.getWordCount());
            System.out.println("Character count: " + page.getCharacterCount());
            if (!page.getHeadings().isEmpty()) {
                System.out.println("Headings: " + page.getHeadings());
            }
            if (!page.getTables().isEmpty()) {
                System.out.println("Tables: " + page.getTables().size());
            }
            // Show content preview
            String pageText = page.getRawText();
            if (pageText != null && !pageText.isEmpty()) {
                System.out.println("Content preview: " + pageText.substring(0, Math.min(200, pageText.length())).replace("\n", " "));
            }
        }
    }
    
    @Test
    public void testExcelSheetBySheetExtraction() throws IOException {
        // Test Excel parser with an existing file if available
        File excelFile = new File("demo.xlsx");
        
        if (!excelFile.exists()) {
            System.out.println("Warning: demo.xlsx not found, skipping Excel test");
            // Test that parser supports Excel files
            ExcelFileParser parser = new ExcelFileParser();
            assertTrue(parser.supports("test.xlsx"));
            assertTrue(parser.supports("test.xls"));
            return;
        }
        
        // Parse Excel
        ParsedDocument document = new ExcelFileParser().parse(excelFile);
        
        // Check sheet-based extraction
        assertNotNull(document.getPages());
        List<PageContent> sheets = document.getPages();
        
        // Each sheet should be a page
        assertFalse(sheets.isEmpty());
        
        DocumentMetadata metadata = document.getDocumentMetadata();
        assertNotNull(metadata.getTotalSheets());
        
        System.out.println("\n=== Excel Parsing Results ===");
        System.out.println("Total sheets: " + metadata.getTotalSheets());

        System.out.println(document.toMarkdown());
    }
    
    @Test
    public void testPowerPointSlideExtraction() throws IOException {
        // Test parsing an existing PowerPoint file
        // Since we can't create PowerPoint with converter, we'll test the parser directly
        
        // Create a simple test to verify PowerPointFileParser works
        PowerPointFileParser pptParser = new PowerPointFileParser();

        ParsedDocument document = pptParser.parse(new File("demo.pptx"));

        List<PageContent> pages = document.getPages();
        for (PageContent page : pages) {
            System.out.println("\n--- Page " + page.getPageNumber() + " ---");
            System.out.println("Word count: " + page.getWordCount());
            System.out.println("Character count: " + page.getCharacterCount());
            System.out.println("Content: " + page.getMarkdownContent());
        }
    }
    
    @Test
    public void testFormatPreservation() throws IOException {
        // Test format preservation with demo.pdf if it exists
        File pdfFile = new File(DEMO_PDF_PATH);
        
        if (!pdfFile.exists()) {
            System.out.println("Warning: demo.pdf not found, skipping format preservation test");
            return;
        }
        
        ParsedDocument pdfDoc = universalParser.parse(pdfFile);
        
        String pdfMarkdown = pdfDoc.toMarkdown();
        System.out.println(pdfMarkdown);
    }
    
    @Test
    public void testBackwardCompatibility() throws IOException {
        // Test that old API still works with demo.pdf
        File pdfFile = new File(DEMO_PDF_PATH);
        
        if (!pdfFile.exists()) {
            System.out.println("Warning: demo.pdf not found, skipping backward compatibility test");
            return;
        }
        
        ParsedDocument document = universalParser.parse(pdfFile);
        
        // Old API should still work
        assertNotNull(document.getContent());
        assertNotNull(document.getMarkdownContent());
        assertEquals(ParsedDocument.FileType.PDF, document.getFileType());
        
        // Legacy metadata should be accessible
        assertNotNull(document.getMetadata());
        assertTrue(document.getMetadata().containsKey("Page Count"));
        
        // New API should also work
        assertNotNull(document.getDocumentMetadata());
        assertNotNull(document.getPages());
        
        System.out.println("\n=== Backward Compatibility ===");
        System.out.println("✓ Legacy getContent() works");
        System.out.println("✓ Legacy getMetadata() works");
        System.out.println("✓ New getDocumentMetadata() works");
        System.out.println("✓ New getPages() works");
    }
}