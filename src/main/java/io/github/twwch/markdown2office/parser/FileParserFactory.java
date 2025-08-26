package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Factory class for creating appropriate file parsers based on file type
 */
public class FileParserFactory {
    
    private static final Logger logger = LoggerFactory.getLogger(FileParserFactory.class);
    
    private static final List<FileParser> PARSERS = new ArrayList<>();
    
    static {
        // Initialize parsers in order of preference
        // More specific parsers first, fallback parsers last
        PARSERS.add(new PdfFileParser()); // PDF parser with hidden layer filtering disabled by default
        PARSERS.add(new WordFileParser());
        PARSERS.add(new ExcelFileParser());
        PARSERS.add(new PowerPointFileParser());
        PARSERS.add(new CsvFileParser());
        PARSERS.add(new MarkdownFileParser());
        PARSERS.add(new TextFileParser());
        PARSERS.add(new TikaFileParser()); // Fallback parser - should be last
    }
    
    /**
     * Get the appropriate parser for a given file name
     * @param fileName the name of the file to parse
     * @return FileParser that can handle the file type, or null if no parser supports it
     */
    public static FileParser getParser(String fileName) {
        if (fileName == null || fileName.trim().isEmpty()) {
            logger.warn("Cannot determine parser for null or empty filename");
            return null;
        }
        
        // Find the first parser that supports this file type
        for (FileParser parser : PARSERS) {
            try {
                if (parser.supports(fileName)) {
                    logger.debug("Selected {} for file: {}", parser.getClass().getSimpleName(), fileName);
                    return parser;
                }
            } catch (Exception e) {
                logger.warn("Error checking parser support for {}: {}", parser.getClass().getSimpleName(), e.getMessage());
            }
        }
        
        logger.warn("No parser found for file: {}", fileName);
        return null;
    }
    
    /**
     * Get all available parsers
     * @return list of all registered parsers
     */
    public static List<FileParser> getAllParsers() {
        return new ArrayList<>(PARSERS);
    }
    
    /**
     * Register a custom parser
     * @param parser the parser to register
     */
    public static void registerParser(FileParser parser) {
        if (parser != null) {
            PARSERS.add(0, parser); // Add at beginning to give priority
            logger.info("Registered custom parser: {}", parser.getClass().getSimpleName());
        }
    }
    
    /**
     * Remove a parser from the factory
     * @param parserClass the class of the parser to remove
     * @return true if a parser was removed, false otherwise
     */
    public static boolean removeParser(Class<? extends FileParser> parserClass) {
        boolean removed = PARSERS.removeIf(parser -> parser.getClass().equals(parserClass));
        if (removed) {
            logger.info("Removed parser: {}", parserClass.getSimpleName());
        }
        return removed;
    }
    
    /**
     * Get supported file extensions from all parsers
     * @return array of supported file extensions (with dots, e.g., ".pdf", ".docx")
     */
    public static String[] getSupportedExtensions() {
        // This is a best-effort method - we test common extensions
        List<String> extensions = new ArrayList<>();
        String[] testExtensions = {
            ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".csv", ".txt", ".text", ".md", ".markdown", ".html", ".htm",
            ".xml", ".rtf", ".log", ".odt", ".ods", ".odp", ".pages",
            ".numbers", ".key", ".epub", ".mobi", ".azw", ".azw3"
        };
        
        for (String ext : testExtensions) {
            String testFile = "test" + ext;
            if (getParser(testFile) != null) {
                extensions.add(ext);
            }
        }
        
        return extensions.toArray(new String[0]);
    }
    
    /**
     * Check if a file type is supported
     * @param fileName the file name to check
     * @return true if the file type is supported, false otherwise
     */
    public static boolean isSupported(String fileName) {
        return getParser(fileName) != null;
    }
    
    /**
     * Get parser information for debugging
     * @return string containing information about all registered parsers
     */
    public static String getParserInfo() {
        StringBuilder info = new StringBuilder();
        info.append("Registered Parsers (in order):\n");
        
        for (int i = 0; i < PARSERS.size(); i++) {
            FileParser parser = PARSERS.get(i);
            info.append(String.format("%d. %s\n", i + 1, parser.getClass().getSimpleName()));
        }
        
        return info.toString();
    }
}