package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.DocumentMetadata;
import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Parser for PowerPoint presentations (PPT/PPTX)
 */
public class PowerPointFileParser implements FileParser {
    
    private static final Logger logger = LoggerFactory.getLogger(PowerPointFileParser.class);
    
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
        try {
            SlideShow<?,?> slideShow;
            
            // Determine if it's PPTX or PPT based on file extension
            if (fileName != null && fileName.toLowerCase().endsWith(".pptx")) {
                slideShow = new XMLSlideShow(inputStream);
            } else {
                slideShow = new HSLFSlideShow(inputStream);
            }
            
            return extractContent(slideShow, fileName);
        } catch (Exception e) {
            logger.error("Error parsing PowerPoint file: " + fileName, e);
            throw new IOException("Failed to parse PowerPoint file: " + fileName, e);
        }
    }
    
    @Override
    public boolean supports(String fileName) {
        if (fileName == null) return false;
        String lower = fileName.toLowerCase();
        return lower.endsWith(".pptx") || lower.endsWith(".ppt");
    }
    
    private ParsedDocument extractContent(SlideShow<?,?> slideShow, String fileName) {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.POWERPOINT);
        
        // Create and populate enhanced metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.POWERPOINT);
        metadata.setTotalSlides(slideShow.getSlides().size());
        metadata.setTotalPages(slideShow.getSlides().size()); // Each slide is a page
        
        // Extract properties if available (for PPTX files)
        if (slideShow instanceof XMLSlideShow) {
            XMLSlideShow xmlSlideShow = (XMLSlideShow) slideShow;
            try {
                POIXMLProperties properties = xmlSlideShow.getProperties();
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
                logger.warn("Error extracting PowerPoint properties", e);
            }
        }
        
        // Try to get presentation title from first slide if not set from metadata
        if (metadata.getTitle() == null && !slideShow.getSlides().isEmpty()) {
            String title = extractSlideTitle(slideShow.getSlides().get(0));
            if (title != null && !title.trim().isEmpty()) {
                metadata.setTitle(title);
                parsedDoc.setTitle(title);
            }
        }
        
        StringBuilder allContent = new StringBuilder();
        StringBuilder allMarkdown = new StringBuilder();
        int totalWords = 0;
        int totalChars = 0;
        int totalTables = 0;
        
        // Process each slide as a separate page
        int slideNumber = 1;
        for (Slide<?,?> slide : slideShow.getSlides()) {
            PageContent pageContent = new PageContent(slideNumber);
            StringBuilder slideContent = new StringBuilder();
            StringBuilder slideMarkdown = new StringBuilder();
            
            processSlide(slide, slideNumber, parsedDoc, slideContent, slideMarkdown, pageContent);
            
            String pageText = slideContent.toString();
            String pageMarkdownText = slideMarkdown.toString();
            
            pageContent.setRawText(pageText);
            pageContent.setMarkdownContent(pageMarkdownText);
            
            // Calculate statistics for this slide
            int slideWords = countWords(pageText);
            int slideChars = pageText.length();
            
            totalWords += slideWords;
            totalChars += slideChars;
            
            if (!pageContent.getTables().isEmpty()) {
                totalTables += pageContent.getTables().size();
            }
            
            parsedDoc.addPage(pageContent);
            
            allContent.append(pageText);
            allMarkdown.append(pageMarkdownText).append("\n\n");
            
            slideNumber++;
        }
        
        // Update metadata with statistics
        metadata.setTotalWords(totalWords);
        metadata.setTotalCharacters(totalChars);
        metadata.setTotalCharactersWithSpaces(totalChars); // Same as totalChars for PowerPoint
        metadata.setTotalTables(totalTables);
        
        parsedDoc.setDocumentMetadata(metadata);
        
        // Set legacy content for backward compatibility
        parsedDoc.setContent(allContent.toString());
        parsedDoc.setMarkdownContent(allMarkdown.toString());
        
        // Add legacy metadata
        parsedDoc.addMetadata("Total Slides", String.valueOf(slideShow.getSlides().size()));
        parsedDoc.addMetadata("File Format", fileName != null && fileName.toLowerCase().endsWith(".pptx") ? "PPTX" : "PPT");
        parsedDoc.addMetadata("Word Count", String.valueOf(totalWords));
        parsedDoc.addMetadata("Character Count", String.valueOf(totalChars));
        parsedDoc.addMetadata("Table Count", String.valueOf(totalTables));
        
        try {
            slideShow.close();
        } catch (IOException e) {
            logger.warn("Error closing slideshow", e);
        }
        
        return parsedDoc;
    }
    
    private void processSlide(Slide<?,?> slide, int slideNumber, ParsedDocument parsedDoc, 
                             StringBuilder content, StringBuilder markdown, PageContent pageContent) {
        
        // Add slide header
        content.append("=== Slide ").append(slideNumber).append(" ===\n");
        markdown.append("# Slide ").append(slideNumber).append("\n\n");
        
        String slideTitle = extractSlideTitle(slide);
        if (slideTitle != null && !slideTitle.trim().isEmpty()) {
            content.append("Title: ").append(slideTitle).append("\n");
            markdown.append("## ").append(slideTitle).append("\n\n");
            pageContent.addHeading("## " + slideTitle);
        }
        
        // Process all shapes in the slide
        for (Shape<?,?> shape : slide.getShapes()) {
            if (shape instanceof TextShape) {
                processTextShape((TextShape<?,?>) shape, content, markdown, pageContent);
            } else if (shape instanceof TableShape) {
                processTableShape((TableShape<?,?>) shape, parsedDoc, content, markdown, pageContent);
            }
        }
        
        content.append("\n");
        markdown.append("\n");
    }
    
    private String extractSlideTitle(Slide<?,?> slide) {
        // Look for the title text in the slide
        for (Shape<?,?> shape : slide.getShapes()) {
            if (shape instanceof TextShape) {
                TextShape<?,?> textShape = (TextShape<?,?>) shape;
                String text = extractTextFromShape(textShape);
                if (text != null && !text.trim().isEmpty()) {
                    // Assume first non-empty text shape is the title
                    return text.trim().split("\n")[0]; // Take first line only
                }
            }
        }
        return null;
    }
    
    private void processTextShape(TextShape<?,?> textShape, StringBuilder content, StringBuilder markdown, PageContent pageContent) {
        String text = extractTextFromShape(textShape);
        if (text != null && !text.trim().isEmpty()) {
            content.append(text).append("\n");
            
            // Convert to markdown with basic formatting
            String[] lines = text.split("\n");
            StringBuilder listItems = new StringBuilder();
            boolean inList = false;
            
            for (String line : lines) {
                line = line.trim();
                if (!line.isEmpty()) {
                    // Simple bullet point detection
                    if (line.startsWith("•") || line.startsWith("-") || line.startsWith("*")) {
                        String bulletText = "- " + line.substring(1).trim();
                        markdown.append(bulletText).append("\n");
                        listItems.append(bulletText).append("\n");
                        inList = true;
                    } else {
                        if (inList && listItems.length() > 0) {
                            pageContent.addList(listItems.toString().trim());
                            listItems.setLength(0);
                            inList = false;
                        }
                        
                        markdown.append(line).append("\n\n");
                        if (line.length() > 10) { // Assume longer text is a paragraph
                            pageContent.addParagraph(line);
                        }
                    }
                }
            }
            
            // Add any remaining list items
            if (inList && listItems.length() > 0) {
                pageContent.addList(listItems.toString().trim());
            }
        }
    }
    
    private String extractTextFromShape(TextShape<?,?> textShape) {
        StringBuilder text = new StringBuilder();
        
        try {
            for (TextParagraph<?,?,?> paragraph : textShape.getTextParagraphs()) {
                for (TextRun textRun : paragraph.getTextRuns()) {
                    String runText = textRun.getRawText();
                    if (runText != null) {
                        text.append(runText);
                    }
                }
                text.append("\n");
            }
        } catch (Exception e) {
            logger.warn("Error extracting text from shape", e);
            // Fallback to simple text extraction
            String shapeText = textShape.getText();
            if (shapeText != null) {
                return shapeText;
            }
        }
        
        return text.toString().trim();
    }
    
    private void processTableShape(TableShape<?,?> tableShape, ParsedDocument parsedDoc, 
                                  StringBuilder content, StringBuilder markdown, PageContent pageContent) {
        try {
            ParsedDocument.ParsedTable parsedTable = new ParsedDocument.ParsedTable();
            List<List<String>> tableData = new ArrayList<>();
            
            int rowCount = tableShape.getNumberOfRows();
            int colCount = tableShape.getNumberOfColumns();
            
            boolean firstRow = true;
            boolean hasData = false;
            
            for (int row = 0; row < rowCount; row++) {
                List<String> rowData = new ArrayList<>();
                for (int col = 0; col < colCount; col++) {
                    try {
                        TextShape<?,?> cell = tableShape.getCell(row, col);
                        String cellText = cell != null ? extractTextFromShape(cell) : "";
                        rowData.add(cellText != null ? cellText.trim() : "");
                    } catch (Exception e) {
                        rowData.add(""); // Add empty cell on error
                    }
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
                
                // Add to content
                content.append(String.join("\t", rowData)).append("\n");
            }
            
            if (hasData) {
                parsedTable.setData(tableData);
                parsedDoc.addTable(parsedTable);
                pageContent.addTable(parsedTable);
                
                // Add table to markdown
                markdown.append("\n").append(parsedTable.toMarkdown()).append("\n");
            }
            
        } catch (Exception e) {
            logger.warn("Error processing table shape", e);
            content.append("[Table content could not be extracted]\n");
            markdown.append("*[Table content could not be extracted]*\n\n");
        }
    }
    
    private int countWords(String text) {
        if (text == null || text.trim().isEmpty()) {
            return 0;
        }
        // Remove extra whitespace and count words
        String cleanText = text.replaceAll("\\s+", " ").trim();
        return cleanText.isEmpty() ? 0 : cleanText.split("\\s+").length;
    }
}