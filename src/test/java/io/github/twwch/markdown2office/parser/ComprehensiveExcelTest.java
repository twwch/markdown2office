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

public class ComprehensiveExcelTest {
    
    @Test
    public void testAllExcelScenarios() throws IOException {
        ExcelFileParser parser = new ExcelFileParser();
        
        // Scenario 1: Create proper .xls file with .xls extension
        System.out.println("\n=== Scenario 1: .xls file with .xls extension ===");
        File properXls = createXlsFile("target/proper.xls", "Proper XLS");
        ParsedDocument doc1 = parser.parse(properXls);
        assertNotNull(doc1);
        assertTrue(doc1.getMarkdownContent().contains("Proper XLS"));
        System.out.println("✅ Successfully parsed");
        
        // Scenario 2: Create proper .xlsx file with .xlsx extension
        System.out.println("\n=== Scenario 2: .xlsx file with .xlsx extension ===");
        File properXlsx = createXlsxFile("target/proper.xlsx", "Proper XLSX");
        ParsedDocument doc2 = parser.parse(properXlsx);
        assertNotNull(doc2);
        assertTrue(doc2.getMarkdownContent().contains("Proper XLSX"));
        System.out.println("✅ Successfully parsed");
        
        // Scenario 3: .xls file with wrong extension (.csv)
        System.out.println("\n=== Scenario 3: .xls file with .csv extension ===");
        File xlsAsCsv = createXlsFile("target/fake.csv", "XLS as CSV");
        ParsedDocument doc3 = parser.parse(xlsAsCsv);
        assertNotNull(doc3);
        assertTrue(doc3.getMarkdownContent().contains("XLS as CSV"));
        System.out.println("✅ Successfully detected as .xls and parsed");
        
        // Scenario 4: .xlsx file with wrong extension (.csv)
        System.out.println("\n=== Scenario 4: .xlsx file with .csv extension ===");
        File xlsxAsCsv = createXlsxFile("target/fake2.csv", "XLSX as CSV");
        ParsedDocument doc4 = parser.parse(xlsxAsCsv);
        assertNotNull(doc4);
        assertTrue(doc4.getMarkdownContent().contains("XLSX as CSV"));
        System.out.println("✅ Successfully detected as .xlsx and parsed");
        
        // Scenario 5: .xls file with .xlsx extension (wrong!)
        System.out.println("\n=== Scenario 5: .xls file with .xlsx extension ===");
        File xlsAsXlsx = createXlsFile("target/wrong.xlsx", "XLS as XLSX");
        ParsedDocument doc5 = parser.parse(xlsAsXlsx);
        assertNotNull(doc5);
        assertTrue(doc5.getMarkdownContent().contains("XLS as XLSX"));
        System.out.println("✅ Successfully detected as .xls despite .xlsx extension");
        
        // Scenario 6: .xlsx file with .xls extension (wrong!)
        System.out.println("\n=== Scenario 6: .xlsx file with .xls extension ===");
        File xlsxAsXls = createXlsxFile("target/wrong.xls", "XLSX as XLS");
        ParsedDocument doc6 = parser.parse(xlsxAsXls);
        assertNotNull(doc6);
        assertTrue(doc6.getMarkdownContent().contains("XLSX as XLS"));
        System.out.println("✅ Successfully detected as .xlsx despite .xls extension");
        
        // Scenario 7: .xls file with no extension
        System.out.println("\n=== Scenario 7: .xls file with no extension ===");
        File xlsNoExt = createXlsFile("target/noext", "XLS no extension");
        ParsedDocument doc7 = parser.parse(xlsNoExt);
        assertNotNull(doc7);
        assertTrue(doc7.getMarkdownContent().contains("XLS no extension"));
        System.out.println("✅ Successfully detected as .xls without extension");
        
        // Scenario 8: .xlsx file with no extension
        System.out.println("\n=== Scenario 8: .xlsx file with no extension ===");
        File xlsxNoExt = createXlsxFile("target/noext2", "XLSX no extension");
        ParsedDocument doc8 = parser.parse(xlsxNoExt);
        assertNotNull(doc8);
        assertTrue(doc8.getMarkdownContent().contains("XLSX no extension"));
        System.out.println("✅ Successfully detected as .xlsx without extension");
        
        // Clean up
        properXls.delete();
        properXlsx.delete();
        xlsAsCsv.delete();
        xlsxAsCsv.delete();
        xlsAsXlsx.delete();
        xlsxAsXls.delete();
        xlsNoExt.delete();
        xlsxNoExt.delete();
        
        System.out.println("\n=== All Excel scenarios passed! ===");
        System.out.println("The parser correctly handles:");
        System.out.println("- Proper file extensions");
        System.out.println("- Wrong file extensions");
        System.out.println("- No file extensions");
        System.out.println("- Both .xls and .xlsx formats");
    }
    
    private File createXlsFile(String path, String content) throws IOException {
        File file = new File(path);
        try (Workbook wb = new HSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Test");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(content);
            
            Row row2 = sheet.createRow(1);
            Cell cell2 = row2.createCell(0);
            cell2.setCellValue("This is XLS format (old Excel)");
            
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
        return file;
    }
    
    private File createXlsxFile(String path, String content) throws IOException {
        File file = new File(path);
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Test");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(content);
            
            Row row2 = sheet.createRow(1);
            Cell cell2 = row2.createCell(0);
            cell2.setCellValue("This is XLSX format (new Excel)");
            
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
        return file;
    }
}