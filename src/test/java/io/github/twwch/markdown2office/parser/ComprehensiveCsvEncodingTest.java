package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.CsvFileParser;
import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

public class ComprehensiveCsvEncodingTest {
    
    @Test
    public void testAllCsvFiles() throws IOException {
        CsvFileParser parser = new CsvFileParser();
        
        // Test UTF-8 CSV
        File utf8File = new File("demo.csv");
        if (utf8File.exists()) {
            System.out.println("\n=== Testing UTF-8 CSV ===");
            ParsedDocument doc1 = parser.parse(utf8File);
            String markdown1 = doc1.getMarkdownContent();
            System.out.println(markdown1);
            assertFalse(markdown1.contains("�"), "UTF-8 CSV should not have garbled text");
        }
        
        // Test GBK CSV
        File gbkFile = new File("demo_gbk.csv");
        if (gbkFile.exists()) {
            System.out.println("\n=== Testing GBK CSV ===");
            ParsedDocument doc2 = parser.parse(gbkFile);
            String markdown2 = doc2.getMarkdownContent();
            System.out.println(markdown2);
            assertFalse(markdown2.contains("�"), "GBK CSV should not have garbled text");
        }
        
        // Test GB18030 CSV
        File gb18030File = new File("demo_gb18030.csv");
        if (gb18030File.exists()) {
            System.out.println("\n=== Testing GB18030 CSV ===");
            ParsedDocument doc3 = parser.parse(gb18030File);
            String markdown3 = doc3.getMarkdownContent();
            System.out.println(markdown3);
            assertFalse(markdown3.contains("�"), "GB18030 CSV should not have garbled text");
        }
        
        System.out.println("\n=== All CSV encoding tests passed successfully! ===");
    }
}