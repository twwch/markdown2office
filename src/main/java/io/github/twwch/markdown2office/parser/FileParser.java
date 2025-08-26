package io.github.twwch.markdown2office.parser;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * Interface for parsing various file formats
 */
public interface FileParser {
    
    /**
     * Parse file from file path
     * @param filePath the path to the file
     * @return ParsedDocument containing the extracted content
     * @throws IOException if file cannot be read or parsed
     */
    ParsedDocument parse(String filePath) throws IOException;
    
    /**
     * Parse file from File object
     * @param file the file to parse
     * @return ParsedDocument containing the extracted content
     * @throws IOException if file cannot be read or parsed
     */
    ParsedDocument parse(File file) throws IOException;
    
    /**
     * Parse file from InputStream
     * @param inputStream the input stream to parse
     * @param fileName the original file name (for type detection)
     * @return ParsedDocument containing the extracted content
     * @throws IOException if stream cannot be read or parsed
     */
    ParsedDocument parse(InputStream inputStream, String fileName) throws IOException;
    
    /**
     * Check if this parser supports the given file type
     * @param fileName the file name to check
     * @return true if this parser can handle the file type
     */
    boolean supports(String fileName);
}