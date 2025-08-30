package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.DocumentMetadata;
import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Date;

/**
 * Parser for Word documents (DOC and DOCX)
 */
public class WordFileParser implements FileParser {
    
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
        // Buffer the input stream to allow mark/reset
        BufferedInputStream bufferedStream = new BufferedInputStream(inputStream);
        bufferedStream.mark(8);
        
        // Detect file format
        FileMagic fileMagic = FileMagic.valueOf(bufferedStream);
        bufferedStream.reset();
        
        if (fileMagic == FileMagic.OLE2) {
            // This is a DOC file (old format)
            return parseDocFile(bufferedStream, fileName);
        } else if (fileMagic == FileMagic.OOXML) {
            // This is a DOCX file (new format)
            return parseDocxFile(bufferedStream, fileName);
        } else {
            throw new IOException("Unsupported Word file format: " + fileMagic);
        }
    }
    
    @Override
    public boolean supports(String fileName) {
        if (fileName == null) return false;
        String lower = fileName.toLowerCase();
        return lower.endsWith(".docx") || lower.endsWith(".doc");
    }
    
    /**
     * Parse DOC file (old format)
     */
    private ParsedDocument parseDocFile(InputStream inputStream, String fileName) throws IOException {
        try (HWPFDocument document = new HWPFDocument(inputStream)) {
            return extractDocContent(document, fileName);
        }
    }
    
    /**
     * Parse DOCX file (new format)
     */
    private ParsedDocument parseDocxFile(InputStream inputStream, String fileName) throws IOException {
        try (XWPFDocument document = new XWPFDocument(inputStream)) {
            return extractDocxContent(document, fileName);
        }
    }
    
    /**
     * Extract content from DOC file
     */
    private ParsedDocument extractDocContent(HWPFDocument document, String fileName) {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.WORD);
        
        // Create and populate metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.WORD);
        
        // Extract document properties
        if (document.getSummaryInformation() != null) {
            var sumInfo = document.getSummaryInformation();
            if (sumInfo.getTitle() != null) {
                metadata.setTitle(sumInfo.getTitle());
                parsedDoc.setTitle(sumInfo.getTitle());
            }
            if (sumInfo.getAuthor() != null) {
                metadata.setAuthor(sumInfo.getAuthor());
                parsedDoc.setAuthor(sumInfo.getAuthor());
            }
            if (sumInfo.getSubject() != null) {
                metadata.setSubject(sumInfo.getSubject());
            }
            if (sumInfo.getKeywords() != null) {
                metadata.setKeywords(sumInfo.getKeywords());
            }
            if (sumInfo.getCreateDateTime() != null) {
                metadata.setCreationDate(sumInfo.getCreateDateTime());
            }
            if (sumInfo.getLastSaveDateTime() != null) {
                metadata.setModificationDate(sumInfo.getLastSaveDateTime());
            }
        }
        
        StringBuilder allContent = new StringBuilder();
        StringBuilder allMarkdown = new StringBuilder();
        
        // Use WordExtractor for simpler text extraction
        WordExtractor extractor = new WordExtractor(document);
        
        // Get paragraphs
        String[] paragraphs = extractor.getParagraphText();
        
        // Create pages (simple division for DOC files)
        List<PageContent> pages = new ArrayList<>();
        PageContent currentPage = new PageContent(1);
        StringBuilder pageContent = new StringBuilder();
        StringBuilder pageMarkdown = new StringBuilder();
        
        int paragraphsPerPage = 50;
        int paragraphCount = 0;
        int totalWords = 0;
        int totalChars = 0;
        
        for (String paragraph : paragraphs) {
            if (paragraph != null && !paragraph.trim().isEmpty()) {
                String trimmed = paragraph.trim();
                
                // Add to content
                pageContent.append(trimmed).append("\n");
                allContent.append(trimmed).append("\n");
                
                // Check if it's a heading (simple heuristic)
                if (isLikelyHeading(trimmed)) {
                    pageMarkdown.append("## ").append(trimmed).append("\n\n");
                    allMarkdown.append("## ").append(trimmed).append("\n\n");
                    currentPage.addHeading("## " + trimmed);
                } else {
                    pageMarkdown.append(trimmed).append("\n\n");
                    allMarkdown.append(trimmed).append("\n\n");
                    if (trimmed.length() > 10) {
                        currentPage.addParagraph(trimmed);
                    }
                }
                
                // Count words and characters
                String[] words = trimmed.split("\\s+");
                totalWords += words.length;
                totalChars += trimmed.length();
                
                paragraphCount++;
                
                // Check if we should start a new page
                if (paragraphCount >= paragraphsPerPage) {
                    currentPage.setRawText(pageContent.toString());
                    currentPage.setMarkdownContent(pageMarkdown.toString());
                    currentPage.setWordCount(countWords(pageContent.toString()));
                    currentPage.setCharacterCount(pageContent.toString().length());
                    pages.add(currentPage);
                    
                    // Start new page
                    currentPage = new PageContent(pages.size() + 1);
                    pageContent = new StringBuilder();
                    pageMarkdown = new StringBuilder();
                    paragraphCount = 0;
                }
            }
        }
        
        // Add the last page if it has content
        if (pageContent.length() > 0) {
            currentPage.setRawText(pageContent.toString());
            currentPage.setMarkdownContent(pageMarkdown.toString());
            currentPage.setWordCount(countWords(pageContent.toString()));
            currentPage.setCharacterCount(pageContent.toString().length());
            pages.add(currentPage);
        }
        
        // Add pages to document
        for (PageContent page : pages) {
            parsedDoc.addPage(page);
        }
        
        // Update metadata
        metadata.setTotalWords(totalWords);
        metadata.setTotalCharacters(totalChars);
        metadata.setTotalPages(pages.size());
        
        parsedDoc.setDocumentMetadata(metadata);
        parsedDoc.setContent(allContent.toString());
        parsedDoc.setMarkdownContent(allMarkdown.toString());
        
        // Add legacy metadata
        parsedDoc.addMetadata("Pages", String.valueOf(pages.size()));
        parsedDoc.addMetadata("Word Count", String.valueOf(totalWords));
        parsedDoc.addMetadata("Character Count", String.valueOf(totalChars));
        
        // Clean up
        try {
            extractor.close();
        } catch (IOException e) {
            // Log and continue, already extracted content
        }
        
        return parsedDoc;
    }
    
    /**
     * Extract content from DOCX file (existing implementation)
     */
    private ParsedDocument extractDocxContent(XWPFDocument document, String fileName) {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.WORD);
        
        // Create and populate enhanced metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.WORD);
        
        // Extract document properties
        if (document.getProperties() != null && document.getProperties().getCoreProperties() != null) {
            POIXMLProperties.CoreProperties props = document.getProperties().getCoreProperties();
            if (props.getTitle() != null) {
                metadata.setTitle(props.getTitle());
                parsedDoc.setTitle(props.getTitle());
            }
            if (props.getCreator() != null) {
                metadata.setAuthor(props.getCreator());
                parsedDoc.setAuthor(props.getCreator());
            }
            if (props.getSubject() != null) {
                metadata.setSubject(props.getSubject());
            }
            if (props.getDescription() != null) {
                metadata.setDescription(props.getDescription());
            }
            if (props.getKeywords() != null) {
                metadata.setKeywords(props.getKeywords());
            }
            if (props.getCreated() != null) {
                metadata.setCreationDate(props.getCreated());
            }
            if (props.getModified() != null) {
                metadata.setModificationDate(props.getModified());
            }
        }
        
        // Extract extended properties for statistics
        if (document.getProperties() != null && document.getProperties().getExtendedProperties() != null) {
            POIXMLProperties.ExtendedProperties extProps = document.getProperties().getExtendedProperties();
            if (extProps.getUnderlyingProperties() != null) {
                try {
                    int pages = extProps.getUnderlyingProperties().getPages();
                    if (pages > 0) {
                        metadata.setTotalPages(pages);
                    }
                    int words = extProps.getUnderlyingProperties().getWords();
                    if (words > 0) {
                        metadata.setTotalWords(words);
                    }
                    int characters = extProps.getUnderlyingProperties().getCharacters();
                    if (characters > 0) {
                        metadata.setTotalCharacters(characters);
                    }
                    int charactersWithSpaces = extProps.getUnderlyingProperties().getCharactersWithSpaces();
                    if (charactersWithSpaces > 0) {
                        metadata.setTotalCharactersWithSpaces(charactersWithSpaces);
                    }
                    int paragraphs = extProps.getUnderlyingProperties().getParagraphs();
                    if (paragraphs > 0) {
                        metadata.setTotalParagraphs(paragraphs);
                    }
                } catch (Exception e) {
                    // Extended properties might not be available, continue without them
                }
            }
        }
        
        StringBuilder allContent = new StringBuilder();
        StringBuilder allMarkdown = new StringBuilder();
        
        // Process content page by page (detect page breaks)
        List<List<IBodyElement>> pages = extractPages(document);
        int pageNum = 1;
        
        int totalWords = 0;
        int totalChars = 0;
        int totalTables = 0;
        
        for (List<IBodyElement> pageElements : pages) {
            PageContent pageContent = new PageContent(pageNum);
            StringBuilder pageContentBuilder = new StringBuilder();
            StringBuilder pageMarkdownBuilder = new StringBuilder();
            
            // Process elements in this page
            for (IBodyElement element : pageElements) {
                if (element instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) element;
                    processParagraph(paragraph, pageContentBuilder, pageMarkdownBuilder, pageContent);
                } else if (element instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) element;
                    processTable(table, parsedDoc, pageContentBuilder, pageMarkdownBuilder, pageContent);
                    totalTables++;
                }
            }
            
            String pageText = pageContentBuilder.toString();
            String pageMarkdown = pageMarkdownBuilder.toString();
            
            if (!pageText.trim().isEmpty()) {
                pageContent.setRawText(pageText);
                pageContent.setMarkdownContent(pageMarkdown);
                
                parsedDoc.addPage(pageContent);
                
                allContent.append(pageText);
                allMarkdown.append(pageMarkdown).append("\n\n");
                
                if (pageContent.getWordCount() != null) {
                    totalWords += pageContent.getWordCount();
                }
                if (pageContent.getCharacterCount() != null) {
                    totalChars += pageContent.getCharacterCount();
                }
            }
            
            pageNum++;
        }
        
        // Update metadata with calculated statistics if not already set
        if (metadata.getTotalWords() == null) {
            metadata.setTotalWords(totalWords);
        }
        if (metadata.getTotalCharacters() == null) {
            metadata.setTotalCharacters(totalChars);
        }
        if (metadata.getTotalPages() == null) {
            metadata.setTotalPages(pages.size());
        }
        metadata.setTotalTables(totalTables);
        
        parsedDoc.setDocumentMetadata(metadata);
        
        // Set legacy content for backward compatibility
        parsedDoc.setContent(allContent.toString());
        parsedDoc.setMarkdownContent(allMarkdown.toString());
        
        // Add legacy metadata
        parsedDoc.addMetadata("Pages", String.valueOf(pages.size()));
        parsedDoc.addMetadata("Word Count", String.valueOf(totalWords));
        parsedDoc.addMetadata("Character Count", String.valueOf(totalChars));
        parsedDoc.addMetadata("Table Count", String.valueOf(totalTables));
        
        return parsedDoc;
    }
    
    private List<List<IBodyElement>> extractPages(XWPFDocument document) {
        List<List<IBodyElement>> pages = new ArrayList<>();
        List<IBodyElement> currentPage = new ArrayList<>();
        boolean pageBreakFound = false;
        
        for (IBodyElement element : document.getBodyElements()) {
            // Check for page break in paragraph
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                if (hasPageBreak(paragraph)) {
                    // Add current element to page before breaking
                    currentPage.add(element);
                    // Start new page
                    if (!currentPage.isEmpty()) {
                        pages.add(new ArrayList<>(currentPage));
                        currentPage.clear();
                        pageBreakFound = true;
                    }
                    continue; // Don't add the page break paragraph twice
                }
            }
            
            currentPage.add(element);
        }
        
        // Add the last page
        if (!currentPage.isEmpty()) {
            pages.add(currentPage);
        }
        
        // If no page breaks detected, split by estimated content size
        if (!pageBreakFound && pages.size() == 1) {
            // Try to split large documents into reasonable pages
            List<IBodyElement> allElements = pages.get(0);
            pages.clear();
            
            int elementsPerPage = 50; // Approximate elements per page
            for (int i = 0; i < allElements.size(); i += elementsPerPage) {
                int end = Math.min(i + elementsPerPage, allElements.size());
                pages.add(new ArrayList<>(allElements.subList(i, end)));
            }
        }
        
        return pages;
    }
    
    private boolean hasPageBreak(XWPFParagraph paragraph) {
        // Check for explicit page breaks in runs
        for (XWPFRun run : paragraph.getRuns()) {
            if (run.getCTR() != null) {
                // Check for page break in break list
                if (run.getCTR().getBrList() != null && run.getCTR().getBrList().size() > 0) {
                    for (int i = 0; i < run.getCTR().getBrList().size(); i++) {
                        if (run.getCTR().getBrList().get(i).getType() != null &&
                            "page".equals(run.getCTR().getBrList().get(i).getType().toString())) {
                            return true;
                        }
                    }
                }
                // Check for page break in lastRenderedPageBreak
                if (run.getCTR().getLastRenderedPageBreakList() != null && 
                    run.getCTR().getLastRenderedPageBreakList().size() > 0) {
                    return true;
                }
            }
        }
        
        // Check if paragraph has page break before property
        if (paragraph.getCTP() != null && paragraph.getCTP().getPPr() != null &&
            paragraph.getCTP().getPPr().getPageBreakBefore() != null) {
            return true;
        }
        
        return false;
    }
    
    private void processParagraph(XWPFParagraph paragraph, StringBuilder content, StringBuilder markdown, PageContent pageContent) {
        String text = paragraph.getText();
        if (text == null || text.isEmpty()) {
            content.append("\n");
            markdown.append("\n");
            return;
        }
        
        // Check paragraph style for headings
        String style = paragraph.getStyle();
        boolean isHeading = false;
        String markdownPrefix = "";
        int headingLevel = 0;
        
        if (style != null) {
            if (style.startsWith("Heading1") || style.equals("Title")) {
                markdownPrefix = "# ";
                isHeading = true;
                headingLevel = 1;
            } else if (style.startsWith("Heading2") || style.equals("Subtitle")) {
                markdownPrefix = "## ";
                isHeading = true;
                headingLevel = 2;
            } else if (style.startsWith("Heading3")) {
                markdownPrefix = "### ";
                isHeading = true;
                headingLevel = 3;
            } else if (style.startsWith("Heading4")) {
                markdownPrefix = "#### ";
                isHeading = true;
                headingLevel = 4;
            } else if (style.startsWith("Heading5")) {
                markdownPrefix = "##### ";
                isHeading = true;
                headingLevel = 5;
            } else if (style.startsWith("Heading6")) {
                markdownPrefix = "###### ";
                isHeading = true;
                headingLevel = 6;
            }
        }
        
        // If not detected by style, check by formatting and content patterns
        if (!isHeading) {
            // Check if all runs are bold and font size is larger
            boolean allBold = true;
            int fontSize = 0;
            
            if (!paragraph.getRuns().isEmpty()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    if (run.getText(0) != null && !run.getText(0).trim().isEmpty()) {
                        if (!run.isBold()) {
                            allBold = false;
                        }
                        if (run.getFontSize() > 0) {
                            fontSize = Math.max(fontSize, run.getFontSize());
                        }
                    }
                }
                
                // Consider it a heading if bold and larger font, or matches Chinese heading patterns
                if ((allBold && fontSize >= 14) || 
                    (allBold && text.length() < 50) ||
                    isChineseHeading(text)) {
                    
                    // Determine level based on font size or pattern
                    if (fontSize >= 20 || text.matches("^第[一二三四五六七八九十]+[章节部分篇]")) {
                        markdownPrefix = "# ";
                        headingLevel = 1;
                    } else if (fontSize >= 16 || text.matches("^[一二三四五六七八九十]+[、.]")) {
                        markdownPrefix = "## ";
                        headingLevel = 2;
                    } else if (fontSize >= 14 || text.matches("^\\d+[、.]")) {
                        markdownPrefix = "### ";
                        headingLevel = 3;
                    } else {
                        markdownPrefix = "### ";
                        headingLevel = 3;
                    }
                    isHeading = true;
                }
            }
        }
        
        content.append(text).append("\n");
        
        if (isHeading) {
            markdown.append(markdownPrefix).append(text).append("\n\n");
            pageContent.addHeading(markdownPrefix + text);
        } else {
            // Process runs for formatting
            StringBuilder formattedText = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                String runText = run.getText(0);
                if (runText != null) {
                    if (run.isBold() && run.isItalic()) {
                        formattedText.append("***").append(runText).append("***");
                    } else if (run.isBold()) {
                        formattedText.append("**").append(runText).append("**");
                    } else if (run.isItalic()) {
                        formattedText.append("*").append(runText).append("*");
                    } else if (run.isStrikeThrough()) {
                        formattedText.append("~~").append(runText).append("~~");
                    } else {
                        formattedText.append(runText);
                    }
                }
            }
            
            String formatted = formattedText.length() > 0 ? formattedText.toString() : text;
            
            // Check for list items
            if (paragraph.getNumID() != null) {
                markdown.append("- ").append(formatted).append("\n");
                pageContent.addList("- " + formatted);
            } else {
                markdown.append(formatted).append("\n\n");
                if (text.length() > 10) { // Assume longer text is a paragraph
                    pageContent.addParagraph(text);
                }
            }
        }
    }
    
    private void processTable(XWPFTable table, ParsedDocument parsedDoc, StringBuilder content, StringBuilder markdown, PageContent pageContent) {
        ParsedDocument.ParsedTable parsedTable = new ParsedDocument.ParsedTable();
        List<List<String>> tableData = new ArrayList<>();
        
        boolean firstRow = true;
        for (XWPFTableRow row : table.getRows()) {
            List<String> rowData = new ArrayList<>();
            for (XWPFTableCell cell : row.getTableCells()) {
                String cellText = cell.getText();
                rowData.add(cellText != null ? cellText : "");
            }
            
            if (firstRow && !rowData.isEmpty() && !rowData.stream().allMatch(String::isEmpty)) {
                parsedTable.setHeaders(rowData);
                firstRow = false;
            } else {
                tableData.add(rowData);
            }
            
            // Add to content
            content.append(String.join("\t", rowData)).append("\n");
        }
        
        parsedTable.setData(tableData);
        parsedDoc.addTable(parsedTable);
        pageContent.addTable(parsedTable);
        
        // Add table to markdown
        markdown.append("\n").append(parsedTable.toMarkdown()).append("\n");
    }
    
    private boolean isChineseHeading(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        
        // Chinese heading patterns
        return text.matches("^第[一二三四五六七八九十]+[章节部分篇].*") || // 第一章, 第二节, etc.
               text.matches("^[一二三四五六七八九十]+[、.。].*") || // 一、引言, 二、内容, etc.
               text.matches("^\\d+[、.。].*") || // 1、内容, 2. 标题, etc.
               text.matches("^[(（][一二三四五六七八九十0-9]+[)）].*") || // (一), (1), etc.
               text.matches("^[①②③④⑤⑥⑦⑧⑨⑩].*") || // ①标题, etc.
               (text.length() < 30 && text.matches(".*[概述|简介|介绍|总结|结论|背景|目的|方法|结果]$")); // Common heading endings
    }
    
    private boolean isLikelyHeading(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        
        // Simple heuristics for headings
        return text.length() < 100 && (
            text.matches("^[0-9]+\\..*") || // Numbered headings
            text.matches("^Chapter \\d+.*") || // Chapter headings
            text.matches("^Section \\d+.*") || // Section headings
            isChineseHeading(text) || // Chinese headings
            text.equals(text.toUpperCase()) && text.length() < 50 // All caps short text
        );
    }
    
    private int countWords(String text) {
        if (text == null || text.trim().isEmpty()) {
            return 0;
        }
        return text.trim().split("\\s+").length;
    }
}