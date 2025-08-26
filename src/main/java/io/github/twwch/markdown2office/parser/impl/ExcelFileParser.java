package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.DocumentMetadata;
import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Parser for Excel files (XLS, XLSX)
 */
public class ExcelFileParser implements FileParser {
    
    @Override
    public ParsedDocument parse(String filePath) throws IOException {
        return parse(new File(filePath));
    }
    
    @Override
    public ParsedDocument parse(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            ParsedDocument parsedDoc = parse(fis, file.getName());
            // Set file size if available
            if (parsedDoc.getDocumentMetadata() != null) {
                parsedDoc.getDocumentMetadata().setFileSize(file.length());
            }
            return parsedDoc;
        }
    }
    
    @Override
    public ParsedDocument parse(InputStream inputStream, String fileName) throws IOException {
        Workbook workbook;
        if (fileName.toLowerCase().endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else {
            workbook = new HSSFWorkbook(inputStream);
        }
        
        try {
            return extractContent(workbook, fileName);
        } finally {
            workbook.close();
        }
    }
    
    @Override
    public boolean supports(String fileName) {
        if (fileName == null) return false;
        String lower = fileName.toLowerCase();
        return lower.endsWith(".xlsx") || lower.endsWith(".xls");
    }
    
    private ParsedDocument extractContent(Workbook workbook, String fileName) {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.EXCEL);
        
        // Create and populate enhanced metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.EXCEL);
        metadata.setTotalSheets(workbook.getNumberOfSheets());
        
        // Extract properties if available (for XLSX files)
        if (workbook instanceof XSSFWorkbook) {
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
            try {
                POIXMLProperties properties = xssfWorkbook.getProperties();
                if (properties != null && properties.getCoreProperties() != null) {
                    POIXMLProperties.CoreProperties coreProps = properties.getCoreProperties();
                    if (coreProps.getTitle() != null) {
                        metadata.setTitle(coreProps.getTitle());
                        parsedDoc.setTitle(coreProps.getTitle());
                    }
                    if (coreProps.getCreator() != null) {
                        metadata.setAuthor(coreProps.getCreator());
                        parsedDoc.setAuthor(coreProps.getCreator());
                    }
                    if (coreProps.getSubject() != null) {
                        metadata.setSubject(coreProps.getSubject());
                    }
                    if (coreProps.getDescription() != null) {
                        metadata.setDescription(coreProps.getDescription());
                    }
                    if (coreProps.getKeywords() != null) {
                        metadata.setKeywords(coreProps.getKeywords());
                    }
                    if (coreProps.getCreated() != null) {
                        metadata.setCreationDate(coreProps.getCreated());
                    }
                    if (coreProps.getModified() != null) {
                        metadata.setModificationDate(coreProps.getModified());
                    }
                }
            } catch (Exception e) {
                // Properties might not be accessible, continue without them
            }
        }
        
        StringBuilder allContent = new StringBuilder();
        StringBuilder allMarkdown = new StringBuilder();
        int totalWords = 0;
        int totalChars = 0;
        int totalTables = 0;
        
        // Process each sheet as a separate page
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            
            // Create page content for this sheet
            PageContent pageContent = new PageContent(i + 1); // Page numbers start from 1
            StringBuilder sheetContent = new StringBuilder();
            StringBuilder sheetMarkdown = new StringBuilder();
            
            sheetContent.append("Sheet: ").append(sheetName).append("\n");
            sheetMarkdown.append("# ").append(sheetName).append("\n\n");
            pageContent.addHeading("# " + sheetName);
            
            // Process sheet data
            ParsedDocument.ParsedTable parsedTable = processSheet(sheet, sheetName, sheetContent, sheetMarkdown);
            if (parsedTable != null) {
                parsedDoc.addTable(parsedTable);
                pageContent.addTable(parsedTable);
                totalTables++;
            }
            
            String pageText = sheetContent.toString();
            String pageMarkdownText = sheetMarkdown.toString();
            
            pageContent.setRawText(pageText);
            pageContent.setMarkdownContent(pageMarkdownText);
            
            // Calculate statistics for this sheet
            int sheetWords = countWords(pageText);
            int sheetChars = pageText.length();
            
            totalWords += sheetWords;
            totalChars += sheetChars;
            
            parsedDoc.addPage(pageContent);
            
            allContent.append(pageText);
            allMarkdown.append(pageMarkdownText).append("\n\n");
        }
        
        // Update metadata with statistics
        metadata.setTotalWords(totalWords);
        metadata.setTotalCharacters(totalChars);
        metadata.setTotalCharactersWithSpaces(totalChars); // Same as totalChars for Excel
        metadata.setTotalTables(totalTables);
        metadata.setTotalPages(workbook.getNumberOfSheets()); // Each sheet is a page
        
        parsedDoc.setDocumentMetadata(metadata);
        
        // Set legacy content for backward compatibility
        parsedDoc.setContent(allContent.toString());
        parsedDoc.setMarkdownContent(allMarkdown.toString());
        
        // Add legacy metadata
        parsedDoc.addMetadata("Sheet Count", String.valueOf(workbook.getNumberOfSheets()));
        parsedDoc.addMetadata("Word Count", String.valueOf(totalWords));
        parsedDoc.addMetadata("Character Count", String.valueOf(totalChars));
        parsedDoc.addMetadata("Table Count", String.valueOf(totalTables));
        
        return parsedDoc;
    }
    
    private ParsedDocument.ParsedTable processSheet(Sheet sheet, String sheetName, StringBuilder content, StringBuilder markdown) {
        ParsedDocument.ParsedTable parsedTable = new ParsedDocument.ParsedTable();
        parsedTable.setTitle(sheetName);
        
        List<List<String>> tableData = new ArrayList<>();
        boolean firstRow = true;
        boolean hasData = false;
        
        for (Row row : sheet) {
            List<String> rowData = new ArrayList<>();
            int lastCellNum = row.getLastCellNum();
            
            if (lastCellNum <= 0) continue; // Skip empty rows
            
            for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                String cellValue = getCellValueAsString(cell);
                rowData.add(cellValue);
                content.append(cellValue).append("\t");
            }
            
            if (!rowData.stream().allMatch(String::isEmpty)) {
                hasData = true;
                if (firstRow) {
                    parsedTable.setHeaders(rowData);
                    firstRow = false;
                } else {
                    tableData.add(rowData);
                }
            }
            
            content.append("\n");
        }
        
        if (hasData) {
            parsedTable.setData(tableData);
            markdown.append(parsedTable.toMarkdown()).append("\n");
            return parsedTable;
        }
        
        return null;
    }
    
    private int countWords(String text) {
        if (text == null || text.trim().isEmpty()) {
            return 0;
        }
        // Remove tabs and extra whitespace, then count words
        String cleanText = text.replaceAll("\\t+", " ").replaceAll("\\s+", " ").trim();
        return cleanText.isEmpty() ? 0 : cleanText.split("\\s+").length;
    }
    
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == (long) value) {
                        return String.valueOf((long) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return cell.getCellFormula();
                    }
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}