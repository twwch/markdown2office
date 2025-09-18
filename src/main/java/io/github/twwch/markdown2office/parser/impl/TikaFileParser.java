package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.DocumentMetadata;
import org.apache.tika.Tika;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.sax.BodyContentHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;
import java.util.List;
import java.util.ArrayList;

/**
 * Fallback parser using Apache Tika for various document formats
 * Supports HTML, XML, RTF, and other formats not covered by specific parsers
 */
public class TikaFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(TikaFileParser.class);
    
    // File extensions supported by this parser
    private static final Set<String> SUPPORTED_EXTENSIONS = new HashSet<>(Arrays.asList(
        ".html", ".htm", ".xml", ".rtf", ".odt", ".ods", ".odp", 
        ".pages", ".numbers", ".key", ".epub", ".mobi", ".azw", ".azw3"
    ));
    
    private final Tika tika;
    
    public TikaFileParser() {
        this.tika = new Tika();
    }
    
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
        
        // Check if any supported extension matches
        for (String ext : SUPPORTED_EXTENSIONS) {
            if (lower.endsWith(ext)) {
                return true;
            }
        }
        
        return false;
    }
    
    private ParsedDocument extractContent(InputStream inputStream, String fileName) throws IOException {
        ParsedDocument parsedDoc = new ParsedDocument();

        // Create and populate document metadata
        DocumentMetadata docMetadata = new DocumentMetadata();
        docMetadata.setFileName(fileName);

        // Wrap the input stream in a BufferedInputStream to ensure proper mark/reset support
        BufferedInputStream bufferedInputStream = inputStream instanceof BufferedInputStream ?
            (BufferedInputStream) inputStream : new BufferedInputStream(inputStream);

        try {
            // Use Tika to parse the document
            BodyContentHandler handler = new BodyContentHandler(-1); // No limit
            Metadata metadata = new Metadata();
            ParseContext context = new ParseContext();

            // Set filename in metadata for better type detection
            if (fileName != null) {
                metadata.set("resourceName", fileName);
                metadata.set("Content-Disposition", "inline; filename=\"" + fileName + "\"");

                // For HTML files, bypass auto-detection and parse directly
                if (fileName.toLowerCase().endsWith(".html") || fileName.toLowerCase().endsWith(".htm")) {
                    // Mark the stream for reset if needed
                    bufferedInputStream.mark(Integer.MAX_VALUE);

                    // Read the HTML content directly - keep original content
                    byte[] bytes = bufferedInputStream.readAllBytes();
                    String content = new String(bytes, StandardCharsets.UTF_8);

                    // Keep the original HTML content as-is
                    parsedDoc.setContent(content);

                    // Convert content to markdown with basic formatting
                    String markdownContent = convertToMarkdown(content, fileName);
                    parsedDoc.setMarkdownContent(markdownContent);

                    // Extract metadata
                    extractMetadata(metadata, parsedDoc, fileName, docMetadata);

                    // Set file type
                    setFileType(parsedDoc, metadata, fileName, docMetadata);

                    // Create pages from content
                    List<PageContent> pages = createPages(content, markdownContent);
                    parsedDoc.setPages(pages);

                    // Set document metadata
                    docMetadata.setTotalWords(content.split("\\s+").length);
                    docMetadata.setTotalCharacters(content.length());
                    docMetadata.setTotalPages(pages.size());
                    parsedDoc.setDocumentMetadata(docMetadata);

                    // Add page count to metadata
                    parsedDoc.addMetadata("Page Count", String.valueOf(pages.size()));

                    return parsedDoc;
                }
            }

            // Use AutoDetectParser for other formats
            AutoDetectParser parser = new AutoDetectParser();
            context.set(Parser.class, parser);
            parser.parse(bufferedInputStream, handler, metadata, context);
            
            String content = handler.toString();
            parsedDoc.setContent(content);
            
            // Convert content to markdown with basic formatting
            String markdownContent = convertToMarkdown(content, fileName);
            parsedDoc.setMarkdownContent(markdownContent);
            
            // Extract metadata
            extractMetadata(metadata, parsedDoc, fileName, docMetadata);
            
            // Set file type based on detected MIME type or extension
            setFileType(parsedDoc, metadata, fileName, docMetadata);
            
            // Create pages from content
            List<PageContent> pages = createPages(content, markdownContent);
            parsedDoc.setPages(pages);
            
            // Set document metadata
            docMetadata.setTotalWords(content.split("\\s+").length);
            docMetadata.setTotalCharacters(content.length());
            docMetadata.setTotalPages(pages.size());
            parsedDoc.setDocumentMetadata(docMetadata);
            
            // Add page count to metadata
            parsedDoc.addMetadata("Page Count", String.valueOf(pages.size()));
            
            return parsedDoc;
            
        } catch (Exception e) {
            logger.error("Error parsing file with Tika: " + fileName, e);
            throw new IOException("Failed to parse file with Tika: " + fileName, e);
        }
    }
    
    private void extractMetadata(Metadata metadata, ParsedDocument parsedDoc, String fileName, DocumentMetadata docMetadata) {
        // Extract common metadata fields
        String title = metadata.get("dc:title");
        if (title != null && !title.trim().isEmpty()) {
            parsedDoc.setTitle(title.trim());
            docMetadata.setTitle(title.trim());
        } else if (fileName != null) {
            // Use filename as title if no title found
            String fileTitle = fileName.replaceAll("\\.[^.]+$", "");
            parsedDoc.setTitle(fileTitle);
            docMetadata.setTitle(fileTitle);
        }
        
        String author = metadata.get("meta:author");
        if (author != null && !author.trim().isEmpty()) {
            parsedDoc.setAuthor(author.trim());
            docMetadata.setAuthor(author.trim());
        }
        
        String creator = metadata.get("meta:creator");
        if (creator != null && !creator.trim().isEmpty() && parsedDoc.getAuthor() == null) {
            parsedDoc.setAuthor(creator.trim());
            docMetadata.setAuthor(creator.trim());
        }
        
        // Add other metadata
        String subject = metadata.get("dc:subject");
        if (subject != null && !subject.trim().isEmpty()) {
            parsedDoc.addMetadata("Subject", subject.trim());
        }
        
        String description = metadata.get("dc:description");
        if (description != null && !description.trim().isEmpty()) {
            parsedDoc.addMetadata("Description", description.trim());
        }
        
        String contentType = metadata.get("Content-Type");
        if (contentType != null && !contentType.trim().isEmpty()) {
            parsedDoc.addMetadata("Content Type", contentType.trim());
        }
        
        String language = metadata.get("dc:language");
        if (language != null && !language.trim().isEmpty()) {
            parsedDoc.addMetadata("Language", language.trim());
        }
        
        String creationDate = metadata.get("meta:creation-date");
        if (creationDate != null && !creationDate.trim().isEmpty()) {
            parsedDoc.addMetadata("Creation Date", creationDate.trim());
            // Could parse and set to docMetadata.setCreationDate() if needed
        }
        
        String lastModified = metadata.get("dcterms:modified");
        if (lastModified != null && !lastModified.trim().isEmpty()) {
            parsedDoc.addMetadata("Last Modified", lastModified.trim());
            // Could parse and set to docMetadata.setModificationDate() if needed
        }
        
        // Set subject and keywords if available  
        String subjectStr = metadata.get("dc:subject");
        if (subjectStr != null && !subjectStr.trim().isEmpty()) {
            docMetadata.setSubject(subjectStr.trim());
        }
        
        String keywords = metadata.get("meta:keyword");
        if (keywords != null && !keywords.trim().isEmpty()) {
            docMetadata.setKeywords(keywords.trim());
        }
    }
    
    private void setFileType(ParsedDocument parsedDoc, Metadata metadata, String fileName, DocumentMetadata docMetadata) {
        String contentType = metadata.get("Content-Type");
        
        ParsedDocument.FileType fileType;
        
        if (contentType != null) {
            if (contentType.contains("html")) {
                fileType = ParsedDocument.FileType.HTML;
            } else if (contentType.contains("xml")) {
                fileType = ParsedDocument.FileType.XML;
            } else if (contentType.contains("rtf")) {
                fileType = ParsedDocument.FileType.RTF;
            } else {
                fileType = ParsedDocument.FileType.UNKNOWN;
            }
        } else if (fileName != null) {
            String lower = fileName.toLowerCase();
            if (lower.endsWith(".html") || lower.endsWith(".htm")) {
                fileType = ParsedDocument.FileType.HTML;
            } else if (lower.endsWith(".xml")) {
                fileType = ParsedDocument.FileType.XML;
            } else if (lower.endsWith(".rtf")) {
                fileType = ParsedDocument.FileType.RTF;
            } else {
                fileType = ParsedDocument.FileType.UNKNOWN;
            }
        } else {
            fileType = ParsedDocument.FileType.UNKNOWN;
        }
        
        parsedDoc.setFileType(fileType);
        docMetadata.setFileType(fileType);
    }
    
    private String convertToMarkdown(String content, String fileName) {
        if (content == null || content.trim().isEmpty()) {
            return "";
        }

        // For HTML files, return content as-is without conversion
        if (fileName != null && (fileName.toLowerCase().endsWith(".html") || fileName.toLowerCase().endsWith(".htm"))) {
            return content;
        }

        // For other formats, do basic markdown conversion
        StringBuilder markdown = new StringBuilder();
        String[] paragraphs = content.split("\n\n+");

        for (String paragraph : paragraphs) {
            paragraph = paragraph.trim();
            if (paragraph.isEmpty()) {
                continue;
            }

            // Basic formatting - this is very simple and could be enhanced
            // Remove multiple spaces and clean up
            paragraph = paragraph.replaceAll("\\s+", " ").trim();

            // Check if it looks like a heading (short line, potentially all caps or title case)
            if (isLikelyHeading(paragraph)) {
                markdown.append("## ").append(paragraph).append("\n\n");
            } else {
                markdown.append(paragraph).append("\n\n");
            }
        }

        return markdown.toString();
    }
    
    private boolean isLikelyHeading(String text) {
        if (text == null || text.isEmpty() || text.length() > 80) {
            return false;
        }
        
        // Check if text is shorter and might be a heading
        if (text.length() < 50 && (
            text.equals(text.toUpperCase()) ||  // All caps
            Character.isUpperCase(text.charAt(0)) && text.matches(".*[A-Z].*") // Title case with capitals
        )) {
            return true;
        }
        
        return false;
    }
    
    private List<PageContent> createPages(String content, String markdownContent) {
        List<PageContent> pages = new ArrayList<>();

        // Always put all content in a single page
        PageContent page = new PageContent(1);

        if (content == null || content.trim().isEmpty()) {
            page.setRawText("");
            page.setMarkdownContent("");
            pages.add(page);
            return pages;
        }

        // Set the raw content as-is
        page.setRawText(content);
        page.setMarkdownContent(markdownContent);

        // Calculate word count (handle HTML content appropriately)
        String textForCounting = content;
        if (content.contains("<") && content.contains(">")) {
            // For HTML content, strip tags for word counting only
            textForCounting = content.replaceAll("<[^>]+>", " ")
                                    .replaceAll("&[^;]+;", " ")
                                    .replaceAll("\\s+", " ")
                                    .trim();
        }
        page.setWordCount(textForCounting.split("\\s+").length);
        page.setCharacterCount(content.length());

        // Don't extract structured data, keep content as-is
        // Just add the whole content as a single paragraph
        if (!content.trim().isEmpty()) {
            page.addParagraph(content);
        }

        pages.add(page);
        return pages;
    }
}