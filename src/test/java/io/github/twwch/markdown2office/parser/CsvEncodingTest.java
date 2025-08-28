package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.CsvFileParser;
import org.junit.jupiter.api.Test;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

public class CsvEncodingTest {
    
    @Test
    public void testUtf8ChineseCsv() throws IOException {
        // Create UTF-8 CSV file with Chinese characters
        String csvContent = "姓名,年龄,城市,职位\n" +
                           "张三,28,北京,软件工程师\n" +
                           "李四,32,上海,产品经理\n" +
                           "王五,25,深圳,设计师";
        
        File utf8File = new File("target/test_utf8.csv");
        try (FileOutputStream fos = new FileOutputStream(utf8File)) {
            fos.write(csvContent.getBytes(StandardCharsets.UTF_8));
        }
        
        // Parse the file
        CsvFileParser parser = new CsvFileParser();
        ParsedDocument document = parser.parse(utf8File);
        
        // Verify content is correctly parsed
        assertNotNull(document);
        String markdown = document.getMarkdownContent();
        assertTrue(markdown.contains("张三"), "Should contain Chinese name 张三");
        assertTrue(markdown.contains("北京"), "Should contain Chinese city 北京");
        assertTrue(markdown.contains("软件工程师"), "Should contain Chinese job title");
        
        System.out.println("UTF-8 CSV parsed successfully:");
        System.out.println(markdown);
        
        // Clean up
        utf8File.delete();
    }
    
    @Test
    public void testGbkChineseCsv() throws IOException {
        // Create GBK CSV file with Chinese characters
        String csvContent = "姓名,年龄,城市,职位\n" +
                           "张三,28,北京,软件工程师\n" +
                           "李四,32,上海,产品经理\n" +
                           "王五,25,深圳,设计师";
        
        File gbkFile = new File("target/test_gbk.csv");
        try (FileOutputStream fos = new FileOutputStream(gbkFile)) {
            fos.write(csvContent.getBytes(Charset.forName("GBK")));
        }
        
        // Parse the file
        CsvFileParser parser = new CsvFileParser();
        ParsedDocument document = parser.parse(gbkFile);
        
        // Verify content is correctly parsed
        assertNotNull(document);
        String markdown = document.getMarkdownContent();
        assertTrue(markdown.contains("张三"), "Should contain Chinese name 张三");
        assertTrue(markdown.contains("北京"), "Should contain Chinese city 北京");
        assertTrue(markdown.contains("软件工程师"), "Should contain Chinese job title");
        
        System.out.println("GBK CSV parsed successfully:");
        System.out.println(markdown);
        
        // Clean up
        gbkFile.delete();
    }
    
    @Test
    public void testExistingChineseCsvFile() throws IOException {
        // Test with the existing test_chinese.csv file
        File csvFile = new File("test_chinese.csv");
        if (!csvFile.exists()) {
            System.out.println("test_chinese.csv not found, skipping test");
            return;
        }
        
        CsvFileParser parser = new CsvFileParser();
        ParsedDocument document = parser.parse(csvFile);
        
        assertNotNull(document);
        String markdown = document.getMarkdownContent();
        
        System.out.println("Existing Chinese CSV file parsed:");
        System.out.println(markdown);
        
        // Verify it contains expected Chinese content
        assertTrue(markdown.contains("姓名") || markdown.contains("张三") || 
                  markdown.contains("北京"), "Should contain Chinese characters");
    }
}