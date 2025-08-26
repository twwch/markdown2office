package io.github.twwch.markdown2office.parser.impl;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;
import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Parser for CSV files
 */
public class CsvFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(CsvFileParser.class);
    
    @Override
    public ParsedDocument parse(String filePath) throws IOException {
        return parse(new File(filePath));
    }
    
    @Override
    public ParsedDocument parse(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            return parse(fis, file.getName());
        }
    }
    
    @Override
    public ParsedDocument parse(InputStream inputStream, String fileName) throws IOException {
        try (InputStreamReader reader = new InputStreamReader(inputStream, "UTF-8")) {
            return extractContent(reader, fileName);
        }
    }
    
    @Override
    public boolean supports(String fileName) {
        if (fileName == null) return false;
        String lower = fileName.toLowerCase();
        return lower.endsWith(".csv");
    }
    
    private ParsedDocument extractContent(Reader reader, String fileName) throws IOException {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.CSV);
        
        try (CSVReader csvReader = new CSVReaderBuilder(reader).build()) {
            List<String[]> allRows = csvReader.readAll();
            
            if (allRows.isEmpty()) {
                parsedDoc.setContent("");
                parsedDoc.setMarkdownContent("");
                return parsedDoc;
            }
            
            // Create table structure
            ParsedDocument.ParsedTable parsedTable = new ParsedDocument.ParsedTable();
            
            // Set file name as table title if available
            if (fileName != null && !fileName.isEmpty()) {
                String title = fileName.replaceAll("\\.[^.]+$", ""); // Remove extension
                parsedTable.setTitle(title);
            }
            
            // First row as headers
            if (!allRows.isEmpty()) {
                List<String> headers = Arrays.asList(allRows.get(0));
                parsedTable.setHeaders(headers);
            }
            
            // Remaining rows as data
            List<List<String>> tableData = new ArrayList<>();
            for (int i = 1; i < allRows.size(); i++) {
                List<String> rowData = Arrays.asList(allRows.get(i));
                tableData.add(rowData);
            }
            parsedTable.setData(tableData);
            
            // Add table to parsed document
            parsedDoc.addTable(parsedTable);
            
            // Generate content and markdown
            StringBuilder content = new StringBuilder();
            StringBuilder markdown = new StringBuilder();
            
            // Add title if available
            if (parsedTable.getTitle() != null && !parsedTable.getTitle().isEmpty()) {
                content.append("Table: ").append(parsedTable.getTitle()).append("\n");
                markdown.append("# ").append(parsedTable.getTitle()).append("\n\n");
            }
            
            // Add all rows to content (tab-separated)
            for (String[] row : allRows) {
                content.append(String.join("\t", row)).append("\n");
            }
            
            // Add table to markdown
            markdown.append(parsedTable.toMarkdown());
            
            parsedDoc.setContent(content.toString());
            parsedDoc.setMarkdownContent(markdown.toString());
            
            // Add metadata
            parsedDoc.addMetadata("Total Rows", String.valueOf(allRows.size()));
            parsedDoc.addMetadata("Total Columns", String.valueOf(allRows.isEmpty() ? 0 : allRows.get(0).length));
            
            return parsedDoc;
            
        } catch (CsvException e) {
            logger.error("Error parsing CSV file: " + fileName, e);
            throw new IOException("Failed to parse CSV file: " + fileName, e);
        }
    }
}