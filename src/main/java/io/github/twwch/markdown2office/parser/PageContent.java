package io.github.twwch.markdown2office.parser;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents content from a specific page in the document
 */
public class PageContent {
    
    private int pageNumber;
    private String rawText;           // Original text from the page
    private String markdownContent;   // Converted to markdown format
    private List<ParsedDocument.ParsedTable> tables;
    private List<ParsedDocument.ParsedImage> images;
    private List<String> headings;
    private List<String> paragraphs;
    private List<String> lists;
    private Integer wordCount;
    private Integer characterCount;
    
    public PageContent(int pageNumber) {
        this.pageNumber = pageNumber;
        this.tables = new ArrayList<>();
        this.images = new ArrayList<>();
        this.headings = new ArrayList<>();
        this.paragraphs = new ArrayList<>();
        this.lists = new ArrayList<>();
    }
    
    // Getters and Setters
    public int getPageNumber() {
        return pageNumber;
    }
    
    public void setPageNumber(int pageNumber) {
        this.pageNumber = pageNumber;
    }
    
    public String getRawText() {
        return rawText;
    }
    
    public void setRawText(String rawText) {
        this.rawText = rawText;
        // Auto-calculate counts
        if (rawText != null) {
            this.characterCount = rawText.length();
            this.wordCount = rawText.split("\\s+").length;
        }
    }
    
    public String getMarkdownContent() {
        return markdownContent;
    }
    
    public void setMarkdownContent(String markdownContent) {
        this.markdownContent = markdownContent;
    }
    
    public List<ParsedDocument.ParsedTable> getTables() {
        return tables;
    }
    
    public void setTables(List<ParsedDocument.ParsedTable> tables) {
        this.tables = tables;
    }
    
    public void addTable(ParsedDocument.ParsedTable table) {
        this.tables.add(table);
    }
    
    public List<ParsedDocument.ParsedImage> getImages() {
        return images;
    }
    
    public void setImages(List<ParsedDocument.ParsedImage> images) {
        this.images = images;
    }
    
    public void addImage(ParsedDocument.ParsedImage image) {
        this.images.add(image);
    }
    
    public List<String> getHeadings() {
        return headings;
    }
    
    public void setHeadings(List<String> headings) {
        this.headings = headings;
    }
    
    public void addHeading(String heading) {
        this.headings.add(heading);
    }
    
    public List<String> getParagraphs() {
        return paragraphs;
    }
    
    public void setParagraphs(List<String> paragraphs) {
        this.paragraphs = paragraphs;
    }
    
    public void addParagraph(String paragraph) {
        this.paragraphs.add(paragraph);
    }
    
    public List<String> getLists() {
        return lists;
    }
    
    public void setLists(List<String> lists) {
        this.lists = lists;
    }
    
    public void addList(String list) {
        this.lists.add(list);
    }
    
    public Integer getWordCount() {
        return wordCount;
    }
    
    public void setWordCount(Integer wordCount) {
        this.wordCount = wordCount;
    }
    
    public Integer getCharacterCount() {
        return characterCount;
    }
    
    public void setCharacterCount(Integer characterCount) {
        this.characterCount = characterCount;
    }
    
    /**
     * Convert page content to markdown format
     */
    public String toMarkdown() {
        if (markdownContent != null && !markdownContent.isEmpty()) {
            return markdownContent;
        }
        
        StringBuilder sb = new StringBuilder();
        sb.append("---\n");
        sb.append("Page: ").append(pageNumber).append("\n");
        sb.append("---\n\n");
        
        // Add headings
        if (!headings.isEmpty()) {
            for (String heading : headings) {
                sb.append(heading).append("\n\n");
            }
        }
        
        // Add paragraphs
        if (!paragraphs.isEmpty()) {
            for (String paragraph : paragraphs) {
                sb.append(paragraph).append("\n\n");
            }
        }
        
        // Add lists
        if (!lists.isEmpty()) {
            for (String list : lists) {
                sb.append(list).append("\n");
            }
            sb.append("\n");
        }
        
        // Add tables
        if (!tables.isEmpty()) {
            for (ParsedDocument.ParsedTable table : tables) {
                sb.append(table.toMarkdown()).append("\n");
            }
        }
        
        // Add images
        if (!images.isEmpty()) {
            for (ParsedDocument.ParsedImage image : images) {
                sb.append(image.toMarkdown()).append("\n\n");
            }
        }
        
        // If no structured content, fall back to raw text
        if (sb.length() == 0 && rawText != null) {
            sb.append(rawText);
        }
        
        return sb.toString();
    }
    
    /**
     * Check if the page has any content
     */
    public boolean hasContent() {
        return (rawText != null && !rawText.trim().isEmpty()) ||
               !headings.isEmpty() ||
               !paragraphs.isEmpty() ||
               !lists.isEmpty() ||
               !tables.isEmpty() ||
               !images.isEmpty();
    }
}