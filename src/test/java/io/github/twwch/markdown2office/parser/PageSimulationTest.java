package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.MarkdownFileParser;
import io.github.twwch.markdown2office.parser.impl.TextFileParser;
import io.github.twwch.markdown2office.parser.impl.TikaFileParser;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class PageSimulationTest {

    @Test
    public void testMarkdownFileParserPageSimulation() throws IOException {
        MarkdownFileParser parser = new MarkdownFileParser();
        
        String markdownContent = "# Title\n\n" +
                "This is the first paragraph.\n\n" +
                "## Section 1\n\n" +
                "Content for section 1.\n\n" +
                "## Section 2\n\n" +
                "Content for section 2.\n\n" +
                "```java\n" +
                "public class Test {\n" +
                "    public static void main(String[] args) {\n" +
                "        System.out.println(\"Hello\");\n" +
                "    }\n" +
                "}\n" +
                "```\n\n" +
                "Final paragraph.";
        
        ByteArrayInputStream inputStream = new ByteArrayInputStream(
                markdownContent.getBytes(StandardCharsets.UTF_8));
        
        ParsedDocument doc = parser.parse(inputStream, "test.md");
        
        assertNotNull(doc);
        assertNotNull(doc.getPages());
        assertFalse(doc.getPages().isEmpty(), "Pages should not be empty");
        
        // Verify DocumentMetadata
        assertNotNull(doc.getDocumentMetadata());
        DocumentMetadata metadata = doc.getDocumentMetadata();
        assertEquals("test.md", metadata.getFileName());
        assertEquals(ParsedDocument.FileType.MARKDOWN, metadata.getFileType());
        assertNotNull(metadata.getTotalPages());
        assertNotNull(metadata.getTotalWords());
        assertNotNull(metadata.getTotalCharacters());
        
        // Verify pages have content
        List<PageContent> pages = doc.getPages();
        for (PageContent page : pages) {
            assertNotNull(page.getRawText());
            assertNotNull(page.getMarkdownContent());
            assertTrue(page.getPageNumber() > 0);
        }
        
        System.out.println("Markdown Parser - Total pages: " + pages.size());
        System.out.println("Markdown Parser - Total words: " + metadata.getTotalWords());
        System.out.println("Markdown Parser - Total characters: " + metadata.getTotalCharacters());
    }
    
    @Test
    public void testTextFileParserPageSimulation() throws IOException {
        TextFileParser parser = new TextFileParser();
        
        StringBuilder content = new StringBuilder();
        for (int i = 0; i < 100; i++) {
            content.append("Line ").append(i + 1).append(": This is a test line with some content.\n");
        }
        
        ByteArrayInputStream inputStream = new ByteArrayInputStream(
                content.toString().getBytes(StandardCharsets.UTF_8));
        
        ParsedDocument doc = parser.parse(inputStream, "test.txt");
        
        assertNotNull(doc);
        assertNotNull(doc.getPages());
        assertFalse(doc.getPages().isEmpty(), "Pages should not be empty");
        
        // Verify DocumentMetadata
        assertNotNull(doc.getDocumentMetadata());
        DocumentMetadata metadata = doc.getDocumentMetadata();
        assertEquals("test.txt", metadata.getFileName());
        assertEquals(ParsedDocument.FileType.TEXT, metadata.getFileType());
        assertNotNull(metadata.getTotalPages());
        assertEquals(1, metadata.getTotalPages(), "Should have exactly 1 page (simulated)");
        
        // Verify pages have content
        List<PageContent> pages = doc.getPages();
        for (PageContent page : pages) {
            assertNotNull(page.getRawText());
            assertNotNull(page.getMarkdownContent());
            assertTrue(page.getPageNumber() > 0);
            assertNotNull(page.getWordCount());
            assertNotNull(page.getCharacterCount());
        }
        
        System.out.println("Text Parser - Total pages: " + pages.size());
        System.out.println("Text Parser - Total words: " + metadata.getTotalWords());
        System.out.println("Text Parser - Total characters: " + metadata.getTotalCharacters());
    }
    
    @Test
    public void testTikaFileParserPageSimulation() throws IOException {
        TikaFileParser parser = new TikaFileParser();
        
        String htmlContent = "<html><head><title>Test Document</title></head><body>" +
                "<h1>Main Title</h1>" +
                "<p>This is the first paragraph of the document.</p>" +
                "<h2>Section 1</h2>" +
                "<p>Content for section 1.</p>" +
                "<h2>Section 2</h2>" +
                "<p>Content for section 2.</p>" +
                "<p>Another paragraph with more content.</p>" +
                "</body></html>";
        
        ByteArrayInputStream inputStream = new ByteArrayInputStream(
                htmlContent.getBytes(StandardCharsets.UTF_8));
        
        ParsedDocument doc = parser.parse(inputStream, "test.html");
        
        assertNotNull(doc);
        assertNotNull(doc.getPages());
        assertFalse(doc.getPages().isEmpty(), "Pages should not be empty");
        
        // Verify DocumentMetadata
        assertNotNull(doc.getDocumentMetadata());
        DocumentMetadata metadata = doc.getDocumentMetadata();
        assertEquals("test.html", metadata.getFileName());
        assertEquals(ParsedDocument.FileType.HTML, metadata.getFileType());
        assertNotNull(metadata.getTotalPages());
        assertNotNull(metadata.getTotalWords());
        assertNotNull(metadata.getTotalCharacters());
        
        // Verify pages have content
        List<PageContent> pages = doc.getPages();
        for (PageContent page : pages) {
            assertNotNull(page.getRawText());
            assertNotNull(page.getMarkdownContent());
            assertTrue(page.getPageNumber() > 0);
        }
        
        System.out.println("Tika Parser - Total pages: " + pages.size());
        System.out.println("Tika Parser - Total words: " + metadata.getTotalWords());
        System.out.println("Tika Parser - Total characters: " + metadata.getTotalCharacters());
    }
    
    @Test
    public void testConsistentReturnFormat() throws IOException {
        // Test that all three parsers return consistent format
        MarkdownFileParser mdParser = new MarkdownFileParser();
        TextFileParser txtParser = new TextFileParser();
        TikaFileParser tikaParser = new TikaFileParser();
        
        String testContent = "Test content for all parsers";
        
        // Test Markdown parser
        ParsedDocument mdDoc = mdParser.parse(
                new ByteArrayInputStream(testContent.getBytes(StandardCharsets.UTF_8)), 
                "test.md");
        
        // Test Text parser
        ParsedDocument txtDoc = txtParser.parse(
                new ByteArrayInputStream(testContent.getBytes(StandardCharsets.UTF_8)), 
                "test.txt");
        
        // Test Tika parser with HTML
        String htmlContent = "<html><body>" + testContent + "</body></html>";
        ParsedDocument tikaDoc = tikaParser.parse(
                new ByteArrayInputStream(htmlContent.getBytes(StandardCharsets.UTF_8)), 
                "test.html");
        
        // All should have pages
        assertNotNull(mdDoc.getPages(), "Markdown parser should have pages");
        assertNotNull(txtDoc.getPages(), "Text parser should have pages");
        assertNotNull(tikaDoc.getPages(), "Tika parser should have pages");
        
        // All should have DocumentMetadata
        assertNotNull(mdDoc.getDocumentMetadata(), "Markdown parser should have metadata");
        assertNotNull(txtDoc.getDocumentMetadata(), "Text parser should have metadata");
        assertNotNull(tikaDoc.getDocumentMetadata(), "Tika parser should have metadata");
        
        // All metadata should have basic fields set
        assertNotNull(mdDoc.getDocumentMetadata().getTotalPages());
        assertNotNull(txtDoc.getDocumentMetadata().getTotalPages());
        assertNotNull(tikaDoc.getDocumentMetadata().getTotalPages());
        
        assertNotNull(mdDoc.getDocumentMetadata().getTotalWords());
        assertNotNull(txtDoc.getDocumentMetadata().getTotalWords());
        assertNotNull(tikaDoc.getDocumentMetadata().getTotalWords());
        
        assertNotNull(mdDoc.getDocumentMetadata().getTotalCharacters());
        assertNotNull(txtDoc.getDocumentMetadata().getTotalCharacters());
        assertNotNull(tikaDoc.getDocumentMetadata().getTotalCharacters());
        
        System.out.println("Format consistency test passed!");
    }
}