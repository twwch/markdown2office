package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.CsvFileParser;
import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

public class GbkCsvTest {
    
    @Test
    public void testRealGbkCsvFile() throws IOException {
        // Test with demo_gbk.csv file
        File gbkFile = new File("demo_gbk.csv");
        if (!gbkFile.exists()) {
            System.out.println("demo_gbk.csv not found, skipping test");
            return;
        }
        
        CsvFileParser parser = new CsvFileParser();
        ParsedDocument document = parser.parse(gbkFile);
        
        assertNotNull(document);
        String markdown = document.getMarkdownContent();
        System.out.println("GBK CSV file parsed:");
        System.out.println(markdown);
        
        // Verify it contains expected Chinese content
        assertTrue(markdown.contains("姓名") || markdown.contains("张三") || 
                  markdown.contains("北京"), "Should contain Chinese characters");
        
        // Check that there are no garbled characters
        assertFalse(markdown.contains("�"), "Should not contain replacement characters");
        assertFalse(markdown.contains("???"), "Should not contain question marks");
    }
}