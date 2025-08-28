package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.ExcelFileParser;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

public class ExcelFormatDetectionTest {
    
    @Test
    public void testExcelFormatAutoDetection() throws IOException {
        // Create test .xls file (old format)
        File xlsFile = new File("target/test_old.xls");
        try (Workbook wb = new HSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Test XLS Format");
            
            try (FileOutputStream fos = new FileOutputStream(xlsFile)) {
                wb.write(fos);
            }
        }
        
        // Create test .xlsx file (new format)
        File xlsxFile = new File("target/test_new.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Test XLSX Format");
            
            try (FileOutputStream fos = new FileOutputStream(xlsxFile)) {
                wb.write(fos);
            }
        }
        
        // Test parsing both formats
        ExcelFileParser parser = new ExcelFileParser();
        
        // Test .xls format
        System.out.println("\n=== Testing .xls (old Excel) format ===");
        ParsedDocument xlsDoc = parser.parse(xlsFile);
        assertNotNull(xlsDoc);
        assertTrue(xlsDoc.getMarkdownContent().contains("Test XLS Format"));
        System.out.println("Successfully parsed .xls file");
        
        // Test .xlsx format
        System.out.println("\n=== Testing .xlsx (new Excel) format ===");
        ParsedDocument xlsxDoc = parser.parse(xlsxFile);
        assertNotNull(xlsxDoc);
        assertTrue(xlsxDoc.getMarkdownContent().contains("Test XLSX Format"));
        System.out.println("Successfully parsed .xlsx file");
        
        // Test with wrongly named files
        // Rename .xls as .xlsx
        File wrongXls = new File("target/wrong.xlsx");
        xlsFile.renameTo(wrongXls);
        
        System.out.println("\n=== Testing .xls file with .xlsx extension ===");
        ParsedDocument wrongXlsDoc = parser.parse(wrongXls);
        assertNotNull(wrongXlsDoc);
        assertTrue(wrongXlsDoc.getMarkdownContent().contains("Test XLS Format"));
        System.out.println("Successfully detected and parsed as .xls despite wrong extension");
        
        // Rename .xlsx as .xls
        File wrongXlsx = new File("target/wrong.xls");
        xlsxFile.renameTo(wrongXlsx);
        
        System.out.println("\n=== Testing .xlsx file with .xls extension ===");
        ParsedDocument wrongXlsxDoc = parser.parse(wrongXlsx);
        assertNotNull(wrongXlsxDoc);
        assertTrue(wrongXlsxDoc.getMarkdownContent().contains("Test XLSX Format"));
        System.out.println("Successfully detected and parsed as .xlsx despite wrong extension");
        
        // Clean up
        wrongXls.delete();
        wrongXlsx.delete();
        
        System.out.println("\n=== All Excel format detection tests passed! ===");
    }
    
    @Test
    public void testRealWorldExcelWithCsvExtension() throws IOException {
        // Test with demo1.csv which is actually Excel
        File demo1 = new File("demo1.csv");
        
        if (!demo1.exists()) {
            System.out.println("demo1.csv not found, skipping test");
            return;
        }
        
        UniversalFileParser parser = new UniversalFileParser();
        
        System.out.println("\n=== Testing real-world Excel file with .csv extension ===");
        ParsedDocument doc = parser.parse(demo1);
        
        assertNotNull(doc);
        assertNotNull(doc.getMarkdownContent());
        assertFalse(doc.getMarkdownContent().isEmpty());
        
        // Should have been auto-detected as Excel
        String parserUsed = doc.getMetadata().get("Parser Used");
        assertTrue(parserUsed.contains("Excel"), "Should use Excel parser");
        
        System.out.println("Successfully handled real-world case!");
        System.out.println("File size: " + demo1.length() + " bytes");
        System.out.println("Content length: " + doc.getContent().length() + " characters");
    }
}