package io.github.twwch.markdown2office.parser.impl;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;
import io.github.twwch.markdown2office.parser.DocumentMetadata;
import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Parser for CSV files with automatic encoding detection
 */
public class CsvFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(CsvFileParser.class);
    
    @Override
    public ParsedDocument parse(String filePath) throws IOException {
        return parse(new File(filePath));
    }
    
    @Override
    public ParsedDocument parse(File file) throws IOException {
        // First check if this is really a CSV file
        if (!isRealCsvFile(file)) {
            throw new IOException("File '" + file.getName() + "' is not a valid CSV file. It appears to be an Excel or other binary file.");
        }
        
        // Detect encoding first
        Charset charset = detectEncoding(file);
        logger.debug("Detected charset for CSV file {}: {}", file.getName(), charset.displayName());
        
        try (InputStreamReader reader = new InputStreamReader(new FileInputStream(file), charset)) {
            return extractContent(reader, file.getName());
        }
    }
    
    @Override
    public ParsedDocument parse(InputStream inputStream, String fileName) throws IOException {
        // For InputStream, we need to buffer it to detect encoding
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        while ((len = inputStream.read(buffer)) > -1) {
            baos.write(buffer, 0, len);
        }
        baos.flush();
        
        byte[] bytes = baos.toByteArray();
        
        // Check if this is really a CSV file
        if (!isRealCsvContent(bytes)) {
            throw new IOException("Stream '" + fileName + "' is not a valid CSV file. It appears to be an Excel or other binary file.");
        }
        
        Charset charset = detectEncoding(bytes);
        logger.debug("Detected charset for CSV stream {}: {}", fileName, charset.displayName());
        
        try (InputStreamReader reader = new InputStreamReader(new ByteArrayInputStream(bytes), charset)) {
            return extractContent(reader, fileName);
        }
    }
    
    @Override
    public boolean supports(String fileName) {
        if (fileName == null) return false;
        String lower = fileName.toLowerCase();
        return lower.endsWith(".csv");
    }
    
    /**
     * Check if file is really a CSV file (not Excel or other binary format)
     */
    private boolean isRealCsvFile(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            byte[] header = new byte[4];
            int bytesRead = fis.read(header);
            if (bytesRead < 4) {
                return true; // Too small to be Excel
            }
            return isRealCsvContent(header);
        }
    }
    
    /**
     * Check if content is really CSV (not Excel or other binary format)
     */
    private boolean isRealCsvContent(byte[] bytes) {
        if (bytes == null || bytes.length < 4) {
            return true; // Too small to determine
        }
        
        // Check for Excel/ZIP signature (PK)
        if (bytes[0] == 0x50 && bytes[1] == 0x4B) {
            logger.warn("File appears to be Excel/ZIP format, not CSV");
            return false;
        }
        
        // Check for old Excel format (D0CF11E0)
        if (bytes[0] == (byte)0xD0 && bytes[1] == (byte)0xCF && 
            bytes[2] == (byte)0x11 && bytes[3] == (byte)0xE0) {
            logger.warn("File appears to be old Excel format, not CSV");
            return false;
        }
        
        // Check for PDF signature (%PDF)
        if (bytes[0] == 0x25 && bytes[1] == 0x50 && 
            bytes[2] == 0x44 && bytes[3] == 0x46) {
            logger.warn("File appears to be PDF format, not CSV");
            return false;
        }
        
        return true;
    }
    
    /**
     * Detect encoding of a file
     */
    private Charset detectEncoding(File file) throws IOException {
        byte[] bytes = Files.readAllBytes(file.toPath());
        return detectEncoding(bytes);
    }
    
    /**
     * Detect encoding from byte array
     * Supports BOM detection and content-based detection for common Chinese encodings
     */
    private Charset detectEncoding(byte[] bytes) {
        if (bytes == null || bytes.length == 0) {
            return StandardCharsets.UTF_8;
        }
        
        // Check for BOM (Byte Order Mark)
        if (bytes.length >= 3) {
            // UTF-8 BOM
            if (bytes[0] == (byte) 0xEF && bytes[1] == (byte) 0xBB && bytes[2] == (byte) 0xBF) {
                logger.debug("Detected UTF-8 BOM");
                return StandardCharsets.UTF_8;
            }
            // UTF-16 BE BOM
            if (bytes[0] == (byte) 0xFE && bytes[1] == (byte) 0xFF) {
                logger.debug("Detected UTF-16 BE BOM");
                return StandardCharsets.UTF_16BE;
            }
            // UTF-16 LE BOM
            if (bytes[0] == (byte) 0xFF && bytes[1] == (byte) 0xFE) {
                logger.debug("Detected UTF-16 LE BOM");
                return StandardCharsets.UTF_16LE;
            }
        }
        
        // Try different encodings and score them
        Charset bestCharset = StandardCharsets.UTF_8;
        int bestScore = 0;
        
        // Test UTF-8
        if (isValidUtf8(bytes)) {
            try {
                String utf8String = new String(bytes, StandardCharsets.UTF_8);
                int score = scoreEncoding(utf8String);
                if (score > bestScore) {
                    bestScore = score;
                    bestCharset = StandardCharsets.UTF_8;
                }
                logger.debug("UTF-8 score: {}", score);
            } catch (Exception e) {
                // Ignore
            }
        }
        
        // Test GBK
        try {
            String gbkString = new String(bytes, Charset.forName("GBK"));
            int score = scoreEncoding(gbkString);
            if (score > bestScore) {
                bestScore = score;
                bestCharset = Charset.forName("GBK");
            }
            logger.debug("GBK score: {}", score);
        } catch (Exception e) {
            // Ignore
        }
        
        // Test GB18030
        try {
            String gb18030String = new String(bytes, Charset.forName("GB18030"));
            int score = scoreEncoding(gb18030String);
            if (score > bestScore) {
                bestScore = score;
                bestCharset = Charset.forName("GB18030");
            }
            logger.debug("GB18030 score: {}", score);
        } catch (Exception e) {
            // Ignore
        }
        
        // Test GB2312
        try {
            String gb2312String = new String(bytes, Charset.forName("GB2312"));
            int score = scoreEncoding(gb2312String);
            if (score > bestScore) {
                bestScore = score;
                bestCharset = Charset.forName("GB2312");
            }
            logger.debug("GB2312 score: {}", score);
        } catch (Exception e) {
            // Ignore
        }
        
        // Test Big5 for Traditional Chinese
        try {
            String big5String = new String(bytes, Charset.forName("Big5"));
            int score = scoreEncoding(big5String);
            if (score > bestScore) {
                bestScore = score;
                bestCharset = Charset.forName("Big5");
            }
            logger.debug("Big5 score: {}", score);
        } catch (Exception e) {
            // Ignore
        }
        
        // Test Windows-1252 for Western European
        try {
            String win1252String = new String(bytes, Charset.forName("Windows-1252"));
            int score = scoreEncoding(win1252String);
            if (score > bestScore) {
                bestScore = score;
                bestCharset = Charset.forName("Windows-1252");
            }
            logger.debug("Windows-1252 score: {}", score);
        } catch (Exception e) {
            // Ignore
        }
        
        logger.info("Best encoding detected: {} with score: {}", bestCharset.displayName(), bestScore);
        return bestCharset;
    }
    
    /**
     * Score an encoding based on various heuristics
     * Higher score means more likely to be correct
     */
    private int scoreEncoding(String str) {
        if (str == null || str.isEmpty()) {
            return 0;
        }
        
        int score = 100; // Start with base score
        
        // Check for replacement characters (indicates wrong encoding)
        int replacementCount = 0;
        for (char c : str.toCharArray()) {
            if (c == '\uFFFD') {
                replacementCount++;
            }
        }
        if (replacementCount > 0) {
            score -= replacementCount * 50;
        }
        
        // Check for control characters (except common ones like tab, newline)
        int controlCharCount = 0;
        for (char c : str.toCharArray()) {
            if (Character.isISOControl(c) && c != '\t' && c != '\n' && c != '\r') {
                controlCharCount++;
            }
        }
        if (controlCharCount > str.length() / 100) {
            score -= controlCharCount * 10;
        }
        
        // Bonus for Chinese characters
        if (containsChineseCharacters(str)) {
            score += 50;
        }
        
        // Bonus for common CSV patterns
        if (str.contains(",") || str.contains("\t") || str.contains("|")) {
            score += 20;
        }
        
        // Check for readable content (alphabetic, numeric, common punctuation)
        int readableCount = 0;
        for (char c : str.toCharArray()) {
            if (Character.isLetterOrDigit(c) || " ,.;:!?()[]{}\"'-_/\\@#$%^&*+=<>".indexOf(c) >= 0 ||
                c == '\n' || c == '\r' || c == '\t' || (c >= 0x4E00 && c <= 0x9FA5)) {
                readableCount++;
            }
        }
        
        // Calculate readability ratio
        double readabilityRatio = (double) readableCount / str.length();
        if (readabilityRatio > 0.95) {
            score += 30;
        } else if (readabilityRatio > 0.9) {
            score += 20;
        } else if (readabilityRatio < 0.7) {
            score -= 30;
        }
        
        return Math.max(0, score);
    }
    
    /**
     * Check if byte array is valid UTF-8
     */
    private boolean isValidUtf8(byte[] bytes) {
        int i = 0;
        while (i < bytes.length) {
            int b = bytes[i] & 0xFF;
            int charBytes;
            
            if (b < 0x80) {
                charBytes = 1;
            } else if ((b & 0xE0) == 0xC0) {
                charBytes = 2;
            } else if ((b & 0xF0) == 0xE0) {
                charBytes = 3;
            } else if ((b & 0xF8) == 0xF0) {
                charBytes = 4;
            } else {
                return false;
            }
            
            if (i + charBytes > bytes.length) {
                return false;
            }
            
            for (int j = 1; j < charBytes; j++) {
                if ((bytes[i + j] & 0xC0) != 0x80) {
                    return false;
                }
            }
            
            i += charBytes;
        }
        return true;
    }
    
    /**
     * Check if string contains Chinese characters
     */
    private boolean containsChineseCharacters(String str) {
        for (char c : str.toCharArray()) {
            if (c >= 0x4E00 && c <= 0x9FA5) { // Common Chinese Unicode range
                return true;
            }
        }
        return false;
    }
    
    /**
     * Check if string contains garbage characters (indication of wrong encoding)
     */
    private boolean containsGarbage(String str) {
        int garbageCount = 0;
        for (char c : str.toCharArray()) {
            // Common garbage characters from wrong encoding
            if (c == '\uFFFD' || // Replacement character
                (c >= 0xFFF0 && c <= 0xFFFF) || // Specials
                (c >= 0xD800 && c <= 0xDFFF)) { // Surrogates
                garbageCount++;
            }
        }
        // If more than 1% of characters are garbage, encoding is likely wrong
        return garbageCount > str.length() / 100;
    }
    
    private ParsedDocument extractContent(Reader reader, String fileName) throws IOException {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.CSV);
        
        // Create and populate enhanced metadata (consistent with Excel)
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileType(ParsedDocument.FileType.CSV);
        metadata.setFileName(fileName);
        
        try (CSVReader csvReader = new CSVReaderBuilder(reader).build()) {
            List<String[]> allRows = csvReader.readAll();
            
            if (allRows.isEmpty()) {
                parsedDoc.setContent("");
                parsedDoc.setMarkdownContent("");
                metadata.setTotalPages(0);
                metadata.setTotalSheets(0);
                parsedDoc.setDocumentMetadata(metadata);
                return parsedDoc;
            }
            
            // Set title from filename (remove extension)
            String title = fileName != null && !fileName.isEmpty() ? 
                fileName.replaceAll("\\.[^.]+$", "") : "CSV Data";
            parsedDoc.setTitle(title);
            metadata.setTitle(title);
            
            // Create table structure
            ParsedDocument.ParsedTable parsedTable = new ParsedDocument.ParsedTable();
            parsedTable.setTitle(title);
            
            // First row as headers
            List<String> headers = null;
            if (!allRows.isEmpty()) {
                headers = Arrays.asList(allRows.get(0));
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
            
            // Create a PageContent object to be consistent with Excel
            PageContent pageContent = new PageContent(1);
            
            // Generate content and markdown
            StringBuilder content = new StringBuilder();
            StringBuilder markdown = new StringBuilder();
            
            // Add title
            markdown.append("# ").append(title).append("\n\n");
            markdown.append("### ").append(title).append("\n\n");
            
            // Add table to markdown
            String tableMarkdown = parsedTable.toMarkdown();
            markdown.append(tableMarkdown);
            
            // Set raw text for the page (tab-separated for consistency)
            StringBuilder rawText = new StringBuilder();
            for (String[] row : allRows) {
                rawText.append(String.join("\t", row)).append("\n");
                content.append(String.join("\t", row)).append("\n");
            }
            
            // Set page content
            pageContent.setRawText(rawText.toString());
            pageContent.setMarkdownContent(markdown.toString());
            pageContent.addTable(parsedTable);
            
            // Add headers as headings for the page
            if (headers != null) {
                pageContent.addHeading(title);
            }
            
            // Calculate word and character counts
            int totalWords = 0;
            int totalChars = 0;
            for (String[] row : allRows) {
                for (String cell : row) {
                    if (cell != null) {
                        totalWords += cell.split("\\s+").length;
                        totalChars += cell.length();
                    }
                }
            }
            
            pageContent.setWordCount(totalWords);
            pageContent.setCharacterCount(totalChars);
            
            // Add the page to the document
            parsedDoc.addPage(pageContent);
            
            // Set document-level content
            parsedDoc.setContent(content.toString());
            parsedDoc.setMarkdownContent(markdown.toString());
            
            // Update metadata
            metadata.setTotalPages(1);
            metadata.setTotalSheets(1); // CSV is like one sheet
            metadata.setTotalWords(totalWords);
            metadata.setTotalCharacters(totalChars);
            metadata.setTotalCharactersWithSpaces(content.toString().length());
            metadata.setTotalTables(1);
            
            parsedDoc.setDocumentMetadata(metadata);
            
            // Add legacy metadata for backward compatibility
            parsedDoc.addMetadata("Total Rows", String.valueOf(allRows.size()));
            parsedDoc.addMetadata("Total Columns", String.valueOf(allRows.isEmpty() ? 0 : allRows.get(0).length));
            parsedDoc.addMetadata("Page Count", "1");
            parsedDoc.addMetadata("Word Count", String.valueOf(totalWords));
            parsedDoc.addMetadata("Character Count", String.valueOf(totalChars));
            
            return parsedDoc;
            
        } catch (CsvException e) {
            logger.error("Error parsing CSV file: " + fileName, e);
            throw new IOException("Failed to parse CSV file: " + fileName, e);
        }
    }
}