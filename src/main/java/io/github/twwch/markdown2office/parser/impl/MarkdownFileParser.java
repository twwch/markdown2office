package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.DocumentMetadata;
import org.commonmark.node.FencedCodeBlock;
import org.commonmark.node.Heading;
import org.commonmark.node.Node;
import org.commonmark.parser.Parser;
import org.commonmark.renderer.text.TextContentRenderer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Parser for Markdown files (MD)
 */
public class MarkdownFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(MarkdownFileParser.class);
    private static final Pattern TABLE_PATTERN = Pattern.compile(
        "(?:^|\\n)\\s*\\|.*\\|\\s*\\n\\s*\\|[-:]+\\|.*\\n(?:\\s*\\|.*\\|\\s*\\n)*", 
        Pattern.MULTILINE
    );
    
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
        return lower.endsWith(".md") || 
               lower.endsWith(".markdown") ||
               lower.endsWith(".mdown") ||
               lower.endsWith(".mkd") ||
               lower.endsWith(".mkdn");
    }
    
    private ParsedDocument extractContent(InputStream inputStream, String fileName) throws IOException {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.MARKDOWN);
        
        // Create and populate metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.MARKDOWN);
        
        StringBuilder content = new StringBuilder();
        List<String> lines = new ArrayList<>();
        
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(inputStream, StandardCharsets.UTF_8))) {
            
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line).append("\n");
                lines.add(line);
            }
            
        } catch (IOException e) {
            logger.error("Error reading markdown file: " + fileName, e);
            throw new IOException("Failed to read markdown file: " + fileName, e);
        }
        
        String markdownContent = content.toString();
        
        // Parse markdown using CommonMark
        Parser parser = Parser.builder().build();
        Node document = parser.parse(markdownContent);
        
        // Extract plain text content
        TextContentRenderer textRenderer = TextContentRenderer.builder().build();
        String plainTextContent = textRenderer.render(document);
        
        parsedDoc.setContent(plainTextContent);
        parsedDoc.setMarkdownContent(markdownContent); // Keep original markdown
        
        // Extract structured information and create pages
        List<PageContent> pages = createPages(document, lines, markdownContent);
        parsedDoc.setPages(pages);
        
        // Extract title and set in metadata
        extractTitle(document, parsedDoc, fileName, metadata);
        
        // Add metadata
        metadata.setTotalWords(plainTextContent.split("\\s+").length);
        metadata.setTotalCharacters(markdownContent.length());
        metadata.setTotalPages(pages.size());
        parsedDoc.setDocumentMetadata(metadata);
        
        parsedDoc.addMetadata("File Type", "Markdown");
        parsedDoc.addMetadata("Line Count", String.valueOf(lines.size()));
        parsedDoc.addMetadata("Character Count", String.valueOf(markdownContent.length()));
        parsedDoc.addMetadata("Page Count", String.valueOf(pages.size()));
        
        // Extract tables using regex (CommonMark might not parse all table formats)
        extractTables(markdownContent, parsedDoc);
        
        return parsedDoc;
    }
    
    private void extractTitle(Node document, ParsedDocument parsedDoc, String fileName, DocumentMetadata metadata) {
        // Walk through the AST to find title
        Node node = document.getFirstChild();
        String title = null;
        
        while (node != null && title == null) {
            if (node instanceof Heading) {
                Heading heading = (Heading) node;
                if (heading.getLevel() == 1) {
                    // Use first H1 as title
                    TextContentRenderer textRenderer = TextContentRenderer.builder().build();
                    title = textRenderer.render(heading).trim();
                }
            }
            node = node.getNext();
        }
        
        // Set title from content or filename
        if (title == null && fileName != null && !fileName.isEmpty()) {
            title = fileName.replaceAll("\\.[^.]+$", ""); // Remove extension
        }
        
        if (title != null) {
            parsedDoc.setTitle(title);
            metadata.setTitle(title);
        }
    }
    
    private List<PageContent> createPages(Node document, List<String> lines, String markdownContent) {
        List<PageContent> pages = new ArrayList<>();
        
        // For markdown files, put all content in a single page
        PageContent page = new PageContent(1);
        
        // Extract plain text content
        TextContentRenderer textRenderer = TextContentRenderer.builder().build();
        String plainTextContent = textRenderer.render(document);
        
        page.setRawText(plainTextContent);
        page.setMarkdownContent(markdownContent);
        page.setWordCount(plainTextContent.split("\\s+").length);
        page.setCharacterCount(plainTextContent.length());
        
        // Walk through the document to extract structured data
        Node node = document.getFirstChild();
        while (node != null) {
            if (node instanceof Heading) {
                Heading heading = (Heading) node;
                String headingText = textRenderer.render(heading).trim();
                String headingMarkdown = "#".repeat(heading.getLevel()) + " " + headingText;
                page.addHeading(headingMarkdown);
            } else if (!(node instanceof FencedCodeBlock)) {
                // Extract paragraphs
                String nodeContent = textRenderer.render(node).trim();
                if (nodeContent != null && !nodeContent.isEmpty() && nodeContent.length() > 10) {
                    page.addParagraph(nodeContent);
                }
            }
            node = node.getNext();
        }
        
        pages.add(page);
        return pages;
    }
    
    private void extractTables(String markdownContent, ParsedDocument parsedDoc) {
        Matcher tableMatcher = TABLE_PATTERN.matcher(markdownContent);
        
        while (tableMatcher.find()) {
            String tableText = tableMatcher.group().trim();
            ParsedDocument.ParsedTable table = parseMarkdownTable(tableText);
            if (table != null) {
                parsedDoc.addTable(table);
            }
        }
    }
    
    private ParsedDocument.ParsedTable parseMarkdownTable(String tableText) {
        try {
            String[] lines = tableText.split("\n");
            if (lines.length < 2) {
                return null; // Need at least header and separator
            }
            
            ParsedDocument.ParsedTable table = new ParsedDocument.ParsedTable();
            
            // Parse header row
            String headerLine = lines[0].trim();
            if (headerLine.startsWith("|")) headerLine = headerLine.substring(1);
            if (headerLine.endsWith("|")) headerLine = headerLine.substring(0, headerLine.length() - 1);
            
            String[] headers = headerLine.split("\\|");
            List<String> headerList = new ArrayList<>();
            for (String header : headers) {
                headerList.add(header.trim());
            }
            table.setHeaders(headerList);
            
            // Parse data rows (skip separator row at index 1)
            List<List<String>> data = new ArrayList<>();
            for (int i = 2; i < lines.length; i++) {
                String dataLine = lines[i].trim();
                if (dataLine.isEmpty()) continue;
                
                if (dataLine.startsWith("|")) dataLine = dataLine.substring(1);
                if (dataLine.endsWith("|")) dataLine = dataLine.substring(0, dataLine.length() - 1);
                
                String[] cells = dataLine.split("\\|");
                List<String> rowData = new ArrayList<>();
                for (String cell : cells) {
                    rowData.add(cell.trim());
                }
                data.add(rowData);
            }
            table.setData(data);
            
            return table;
            
        } catch (Exception e) {
            logger.warn("Error parsing markdown table", e);
            return null;
        }
    }
}