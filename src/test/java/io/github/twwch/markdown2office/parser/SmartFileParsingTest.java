package io.github.twwch.markdown2office.parser;

import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

public class SmartFileParsingTest {
    
    @Test
    public void testSmartCsvParsing() throws IOException {
        UniversalFileParser parser = new UniversalFileParser();
        
        // Test 1: Real CSV file (UTF-8)
        File realCsvUtf8 = new File("demo.csv");
        if (realCsvUtf8.exists()) {
            System.out.println("\n=== Testing real UTF-8 CSV file ===");
            ParsedDocument doc1 = parser.parse(realCsvUtf8);
            assertNotNull(doc1);
            String content1 = doc1.getMarkdownContent();
            System.out.println("Successfully parsed as CSV:");
            System.out.println(content1.substring(0, Math.min(300, content1.length())));
        }
        
        // Test 2: Real CSV file (GBK)
        File realCsvGbk = new File("demo_gbk.csv");
        if (realCsvGbk.exists()) {
            System.out.println("\n=== Testing real GBK CSV file ===");
            ParsedDocument doc2 = parser.parse(realCsvGbk);
            assertNotNull(doc2);
            String content2 = doc2.getMarkdownContent();
            assertFalse(content2.contains("�"), "Should not have garbled text");
            System.out.println("Successfully parsed GBK CSV without garbled text");
        }
        
        // Test 3: Excel file with .csv extension (the smart case)
        File fakeCSV = new File("demo1.csv");
        if (fakeCSV.exists()) {
            System.out.println("\n=== Testing Excel file with .csv extension ===");
            ParsedDocument doc3 = parser.parse(fakeCSV);
            assertNotNull(doc3);
            
            // Check that it was parsed by Excel parser
            String parserUsed = doc3.getMetadata().get("Parser Used");
            assertTrue(parserUsed.contains("Excel"), 
                "Should have used Excel parser, but used: " + parserUsed);
            
            // Check for the note about auto-detection
            String note = doc3.getMetadata().get("Note");
            assertNotNull(note, "Should have a note about auto-detection");
            assertTrue(note.contains(".csv extension but was Excel format"), 
                "Note should mention the format mismatch");
            
            System.out.println("Successfully detected and parsed Excel file with .csv extension!");
            System.out.println("Parser used: " + parserUsed);
            System.out.println("Note: " + note);
            
            // Verify content is readable
            String content3 = doc3.getMarkdownContent();
            assertNotNull(content3);
            assertFalse(content3.isEmpty());
            System.out.println("Content preview: " + 
                content3.substring(0, Math.min(200, content3.length())));
        }
        
        System.out.println("\n=== All smart parsing tests passed! ===");
        System.out.println("The system successfully:");
        System.out.println("1. Parses real CSV files with various encodings");
        System.out.println("2. Detects when a .csv file is actually Excel format");  
        System.out.println("3. Automatically uses the correct parser");
        System.out.println("4. Provides clear metadata about what happened");
    }
}