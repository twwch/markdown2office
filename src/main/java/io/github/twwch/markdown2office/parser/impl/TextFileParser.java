package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.DocumentMetadata;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * Parser for plain text files (TXT)
 */
public class TextFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(TextFileParser.class);
    
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
        return extractContent(inputStream, fileName);
    }
    
    @Override
    public boolean supports(String fileName) {
        if (fileName == null) return false;
        String lower = fileName.toLowerCase();
        return lower.endsWith(".txt") || 
               lower.endsWith(".text") ||
               lower.endsWith(".log") ||
               // Check if it's a file without extension (might be plain text)
               (!lower.contains(".") && fileName.length() > 0);
    }
    
    private ParsedDocument extractContent(InputStream inputStream, String fileName) throws IOException {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.TEXT);
        
        // Create and populate metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.TEXT);
        
        List<String> lines = new ArrayList<>();
        StringBuilder content = new StringBuilder();
        String firstLine = null;
        
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(inputStream, StandardCharsets.UTF_8))) {
            
            String line;
            while ((line = reader.readLine()) != null) {
                if (firstLine == null && !line.trim().isEmpty()) {
                    firstLine = line.trim();
                }
                content.append(line).append("\n");
                lines.add(line);
            }
            
        } catch (IOException e) {
            logger.error("Error reading text file: " + fileName, e);
            throw new IOException("Failed to read text file: " + fileName, e);
        }
        
        String textContent = content.toString();
        
        // Set title from filename or first line
        String title = null;
        if (fileName != null && !fileName.isEmpty()) {
            title = fileName.replaceAll("\\.[^.]+$", ""); // Remove extension
        } else if (firstLine != null && firstLine.length() <= 100) {
            // Use first line as title if it's reasonably short
            title = firstLine;
        }
        
        if (title != null) {
            parsedDoc.setTitle(title);
            metadata.setTitle(title);
        }
        
        // For plain text, content and markdown are essentially the same
        // Just add some basic formatting for markdown
        parsedDoc.setContent(textContent);
        
        // Convert to markdown with minimal processing
        String markdownContent = convertToMarkdown(textContent, parsedDoc.getTitle());
        parsedDoc.setMarkdownContent(markdownContent);
        
        // Create pages
        List<PageContent> pages = createPages(lines, textContent);
        parsedDoc.setPages(pages);
        
        // Add metadata
        metadata.setTotalWords(textContent.split("\\s+").length);
        metadata.setTotalCharacters(textContent.length());
        metadata.setTotalPages(pages.size());
        parsedDoc.setDocumentMetadata(metadata);
        
        parsedDoc.addMetadata("File Type", "Plain Text");
        parsedDoc.addMetadata("Line Count", String.valueOf(lines.size()));
        parsedDoc.addMetadata("Character Count", String.valueOf(textContent.length()));
        parsedDoc.addMetadata("Page Count", String.valueOf(pages.size()));
        
        return parsedDoc;
    }
    
    private String convertToMarkdown(String content, String title) {
        StringBuilder markdown = new StringBuilder();
        
        // Add title if available
        if (title != null && !title.isEmpty()) {
            markdown.append("# ").append(title).append("\n\n");
        }
        
        // Process content line by line for basic markdown formatting
        String[] lines = content.split("\n");
        boolean inCodeBlock = false;
        
        for (String line : lines) {
            String trimmedLine = line.trim();
            
            // Detect potential code blocks (lines that start with spaces/tabs)
            if (line.startsWith("    ") || line.startsWith("\t")) {
                if (!inCodeBlock) {
                    markdown.append("```\n");
                    inCodeBlock = true;
                }
                markdown.append(line).append("\n");
            } else {
                if (inCodeBlock) {
                    markdown.append("```\n\n");
                    inCodeBlock = false;
                }
                
                // Detect potential headers (lines that are all caps or followed by ====/----)
                if (isLikelyHeader(trimmedLine)) {
                    markdown.append("## ").append(trimmedLine).append("\n\n");
                } else if (trimmedLine.isEmpty()) {
                    markdown.append("\n");
                } else {
                    // Regular paragraph
                    markdown.append(trimmedLine).append("\n\n");
                }
            }
        }
        
        // Close code block if still open
        if (inCodeBlock) {
            markdown.append("```\n");
        }
        
        return markdown.toString();
    }
    
    private boolean isLikelyHeader(String line) {
        if (line.isEmpty() || line.length() > 80) {
            return false;
        }
        
        // Check if line is all uppercase (likely a header)
        if (line.equals(line.toUpperCase()) && line.matches(".*[A-Z].*")) {
            return true;
        }
        
        // Check if line ends with colon (likely a section header)
        if (line.endsWith(":") && !line.contains(" ") && line.length() > 3) {
            return true;
        }
        
        return false;
    }
    
    private List<PageContent> createPages(List<String> lines, String fullContent) {
        List<PageContent> pages = new ArrayList<>();
        
        // For text files, put all content in a single page
        PageContent page = new PageContent(1);
        
        if (fullContent == null || fullContent.isEmpty()) {
            page.setRawText("");
            page.setMarkdownContent("");
            pages.add(page);
            return pages;
        }
        
        // Convert to markdown with minimal processing
        String markdownContent = convertToMarkdown(fullContent, null);
        
        page.setRawText(fullContent);
        page.setMarkdownContent(markdownContent);
        page.setWordCount(fullContent.split("\\s+").length);
        page.setCharacterCount(fullContent.length());
        
        // Extract structured data
        for (String line : lines) {
            String trimmedLine = line.trim();
            if (isLikelyHeader(trimmedLine)) {
                page.addHeading("## " + trimmedLine);
            } else if (trimmedLine.length() > 10) {
                page.addParagraph(trimmedLine);
            }
        }
        
        pages.add(page);
        return pages;
    }
}