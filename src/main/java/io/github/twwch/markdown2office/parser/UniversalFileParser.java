package io.github.twwch.markdown2office.parser;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * Universal file parser that automatically detects file type and uses appropriate parser
 * This is the main entry point for parsing any supported file format
 */
public class UniversalFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(UniversalFileParser.class);
    
    /**
     * Parse file from file path using automatic type detection
     * @param filePath the path to the file
     * @return ParsedDocument containing the extracted content
     * @throws IOException if file cannot be read or parsed
     * @throws UnsupportedOperationException if no parser supports the file type
     */
    @Override
    public ParsedDocument parse(String filePath) throws IOException {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IllegalArgumentException("File path cannot be null or empty");
        }
        
        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File does not exist: " + filePath);
        }
        
        if (!file.isFile()) {
            throw new IOException("Path is not a file: " + filePath);
        }
        
        return parse(file);
    }
    
    /**
     * Parse file from File object using automatic type detection
     * @param file the file to parse
     * @return ParsedDocument containing the extracted content
     * @throws IOException if file cannot be read or parsed
     * @throws UnsupportedOperationException if no parser supports the file type
     */
    @Override
    public ParsedDocument parse(File file) throws IOException {
        if (file == null) {
            throw new IllegalArgumentException("File cannot be null");
        }
        
        if (!file.exists()) {
            throw new IOException("File does not exist: " + file.getAbsolutePath());
        }
        
        if (!file.isFile()) {
            throw new IOException("Path is not a file: " + file.getAbsolutePath());
        }
        
        if (!file.canRead()) {
            throw new IOException("Cannot read file: " + file.getAbsolutePath());
        }
        
        String fileName = file.getName();
        FileParser parser = FileParserFactory.getParser(fileName);
        
        if (parser == null) {
            throw new UnsupportedOperationException(
                "No parser available for file type: " + fileName + 
                ". Supported extensions: " + String.join(", ", FileParserFactory.getSupportedExtensions())
            );
        }
        
        try {
            logger.info("Parsing file '{}' using {}", fileName, parser.getClass().getSimpleName());
            ParsedDocument result = parser.parse(file);
            
            if (result != null) {
                // Add parsing metadata
                result.addMetadata("Parser Used", parser.getClass().getSimpleName());
                result.addMetadata("File Size", String.valueOf(file.length()));
                result.addMetadata("File Path", file.getAbsolutePath());
                
                logger.info("Successfully parsed file '{}' - extracted {} characters", 
                    fileName, result.getContent() != null ? result.getContent().length() : 0);
            }
            
            return result;
            
        } catch (Exception e) {
            logger.error("Error parsing file '{}' with {}: {}", 
                fileName, parser.getClass().getSimpleName(), e.getMessage(), e);
            throw new IOException("Failed to parse file: " + fileName, e);
        }
    }
    
    /**
     * Parse file from InputStream using automatic type detection
     * @param inputStream the input stream to parse
     * @param fileName the original file name (for type detection)
     * @return ParsedDocument containing the extracted content
     * @throws IOException if stream cannot be read or parsed
     * @throws UnsupportedOperationException if no parser supports the file type
     */
    @Override
    public ParsedDocument parse(InputStream inputStream, String fileName) throws IOException {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        
        if (fileName == null || fileName.trim().isEmpty()) {
            throw new IllegalArgumentException("File name cannot be null or empty for type detection");
        }
        
        FileParser parser = FileParserFactory.getParser(fileName);
        
        if (parser == null) {
            throw new UnsupportedOperationException(
                "No parser available for file type: " + fileName + 
                ". Supported extensions: " + String.join(", ", FileParserFactory.getSupportedExtensions())
            );
        }
        
        try {
            logger.info("Parsing stream for file '{}' using {}", fileName, parser.getClass().getSimpleName());
            ParsedDocument result = parser.parse(inputStream, fileName);
            
            if (result != null) {
                // Add parsing metadata
                result.addMetadata("Parser Used", parser.getClass().getSimpleName());
                result.addMetadata("Source", "InputStream");
                result.addMetadata("File Name", fileName);
                
                logger.info("Successfully parsed stream for file '{}' - extracted {} characters", 
                    fileName, result.getContent() != null ? result.getContent().length() : 0);
            }
            
            return result;
            
        } catch (Exception e) {
            logger.error("Error parsing stream for file '{}' with {}: {}", 
                fileName, parser.getClass().getSimpleName(), e.getMessage(), e);
            throw new IOException("Failed to parse file from stream: " + fileName, e);
        }
    }
    
    /**
     * Check if this parser supports the given file type
     * This implementation always returns true since it delegates to specialized parsers
     * @param fileName the file name to check
     * @return true if any registered parser can handle the file type
     */
    @Override
    public boolean supports(String fileName) {
        return FileParserFactory.isSupported(fileName);
    }
    
    /**
     * Parse file with custom error handling options
     * @param filePath the path to the file
     * @param failSilently if true, returns null instead of throwing exception for unsupported files
     * @return ParsedDocument containing the extracted content, or null if failSilently is true and parsing fails
     * @throws IOException if file cannot be read or parsed (and failSilently is false)
     */
    public ParsedDocument parseWithOptions(String filePath, boolean failSilently) throws IOException {
        try {
            return parse(filePath);
        } catch (UnsupportedOperationException e) {
            if (failSilently) {
                logger.warn("File type not supported, returning null: {}", filePath);
                return null;
            } else {
                throw e;
            }
        }
    }
    
    /**
     * Parse file with custom error handling options
     * @param file the file to parse
     * @param failSilently if true, returns null instead of throwing exception for unsupported files
     * @return ParsedDocument containing the extracted content, or null if failSilently is true and parsing fails
     * @throws IOException if file cannot be read or parsed (and failSilently is false)
     */
    public ParsedDocument parseWithOptions(File file, boolean failSilently) throws IOException {
        try {
            return parse(file);
        } catch (UnsupportedOperationException e) {
            if (failSilently) {
                logger.warn("File type not supported, returning null: {}", file != null ? file.getName() : "null");
                return null;
            } else {
                throw e;
            }
        }
    }
    
    /**
     * Get information about available parsers and supported formats
     * @return string containing parser information
     */
    public String getParserInfo() {
        return FileParserFactory.getParserInfo();
    }
    
    /**
     * Get list of all supported file extensions
     * @return array of supported file extensions
     */
    public String[] getSupportedExtensions() {
        return FileParserFactory.getSupportedExtensions();
    }
}