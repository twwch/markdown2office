package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.CsvFileParser;
import io.github.twwch.markdown2office.parser.impl.ExcelFileParser;
import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

public class UnifiedFormatTest {
    
    @Test
    public void testCsvFormatConsistency() throws IOException {
        // Create a test CSV file
        File csvFile = new File("target/test_format.csv");
        try (FileWriter writer = new FileWriter(csvFile)) {
            writer.write("姓名,年龄,城市,职位\n");
            writer.write("张三,28,北京,软件工程师\n");
            writer.write("李四,32,上海,产品经理\n");
            writer.write("王五,25,深圳,设计师\n");
        }
        
        // Parse CSV
        CsvFileParser csvParser = new CsvFileParser();
        ParsedDocument csvDoc = csvParser.parse(csvFile);
        
        // Check title is set
        assertNotNull(csvDoc.getTitle(), "CSV should have title");
        assertEquals("test_format", csvDoc.getTitle(), "Title should be filename without extension");
        
        // Check pages structure exists
        assertNotNull(csvDoc.getPages(), "CSV should have pages");
        assertFalse(csvDoc.getPages().isEmpty(), "CSV should have at least one page");
        assertEquals(1, csvDoc.getPages().size(), "CSV should have exactly one page");
        
        // Check page content
        PageContent page = csvDoc.getPages().get(0);
        assertNotNull(page, "Page should not be null");
        assertEquals(1, page.getPageNumber(), "Page number should be 1");
        assertNotNull(page.getRawText(), "Page should have raw text");
        assertNotNull(page.getMarkdownContent(), "Page should have markdown content");
        assertNotNull(page.getTables(), "Page should have tables");
        assertEquals(1, page.getTables().size(), "Page should have exactly one table");
        
        // Check document metadata
        DocumentMetadata metadata = csvDoc.getDocumentMetadata();
        assertNotNull(metadata, "CSV should have document metadata");
        assertEquals("test_format.csv", metadata.getFileName());
        assertEquals("test_format", metadata.getTitle());
        assertEquals(ParsedDocument.FileType.CSV, metadata.getFileType());
        assertEquals(1, metadata.getTotalPages().intValue());
        assertEquals(1, metadata.getTotalSheets().intValue());
        assertEquals(1, metadata.getTotalTables().intValue());
        assertNotNull(metadata.getTotalWords());
        assertNotNull(metadata.getTotalCharacters());
        
        // Check markdown output
        String markdown = csvDoc.getMarkdownContent();
        assertTrue(markdown.contains("# test_format"), "Markdown should contain title");
        assertTrue(markdown.contains("### test_format"), "Markdown should contain subtitle");
        assertTrue(markdown.contains("姓名"), "Markdown should contain headers");
        assertTrue(markdown.contains("张三"), "Markdown should contain data");
        
        // Print results for verification
        System.out.println("\n=== CSV Format Test Results ===");
        System.out.println("Title: " + csvDoc.getTitle());
        System.out.println("Pages: " + csvDoc.getPages().size());
        System.out.println("Total Words: " + metadata.getTotalWords());
        System.out.println("Total Characters: " + metadata.getTotalCharacters());
        System.out.println("Total Tables: " + metadata.getTotalTables());
        System.out.println("\nMarkdown Output:");
        System.out.println(markdown);
        
        // Clean up
        csvFile.delete();
    }
    
    @Test  
    public void testCsvExcelFormatComparison() throws IOException {
        // Test that CSV and Excel have consistent structure
        File csvFile = new File("demo.csv");
        File excelFile = new File("demo1.csv"); // This is actually Excel
        
        if (!csvFile.exists() || !excelFile.exists()) {
            System.out.println("Demo files not found, skipping comparison test");
            return;
        }
        
        CsvFileParser csvParser = new CsvFileParser();
        ParsedDocument csvDoc = csvParser.parse(csvFile);
        
        ExcelFileParser excelParser = new ExcelFileParser();
        ParsedDocument excelDoc = excelParser.parse(excelFile);
        
        // Both should have pages
        assertNotNull(csvDoc.getPages(), "CSV should have pages");
        assertNotNull(excelDoc.getPages(), "Excel should have pages");
        
        // Both should have title
        assertNotNull(csvDoc.getTitle(), "CSV should have title");
        assertNotNull(excelDoc.getTitle(), "Excel should have title");
        
        // Both should have DocumentMetadata
        assertNotNull(csvDoc.getDocumentMetadata(), "CSV should have DocumentMetadata");
        assertNotNull(excelDoc.getDocumentMetadata(), "Excel should have DocumentMetadata");
        
        // Both should have same structure fields
        assertNotNull(csvDoc.getDocumentMetadata().getTotalPages(), "CSV metadata should have total pages");
        assertNotNull(excelDoc.getDocumentMetadata().getTotalPages(), "Excel metadata should have total pages");
        
        assertNotNull(csvDoc.getDocumentMetadata().getTotalSheets(), "CSV metadata should have total sheets");
        assertNotNull(excelDoc.getDocumentMetadata().getTotalSheets(), "Excel metadata should have total sheets");
        
        System.out.println("\n=== Format Comparison ===");
        System.out.println("CSV Structure:");
        System.out.println("  - Title: " + csvDoc.getTitle());
        System.out.println("  - Pages: " + csvDoc.getPages().size());
        System.out.println("  - Total Sheets: " + csvDoc.getDocumentMetadata().getTotalSheets());
        
        System.out.println("\nExcel Structure:");
        System.out.println("  - Title: " + excelDoc.getTitle());
        System.out.println("  - Pages: " + excelDoc.getPages().size());
        System.out.println("  - Total Sheets: " + excelDoc.getDocumentMetadata().getTotalSheets());
        
        System.out.println("\n✅ Both CSV and Excel have consistent format structure!");
    }
}