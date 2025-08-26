package io.github.twwch.markdown2office.parser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Represents a parsed document with structured content
 */
public class ParsedDocument {
    
    // New enhanced metadata
    private DocumentMetadata documentMetadata;
    
    // Page-based content for better structure preservation
    private List<PageContent> pages;
    
    // Legacy fields for backward compatibility
    private String title;
    private String author;
    private String content;  // Main text content
    private String markdownContent;  // Converted to markdown format
    private List<ParsedTable> tables;
    private List<ParsedImage> images;
    private Map<String, String> metadata;
    private FileType fileType;
    
    public ParsedDocument() {
        this.documentMetadata = new DocumentMetadata();
        this.pages = new ArrayList<>();
        this.tables = new ArrayList<>();
        this.images = new ArrayList<>();
        this.metadata = new HashMap<>();
    }
    
    // New enhanced getters and setters
    public DocumentMetadata getDocumentMetadata() {
        return documentMetadata;
    }
    
    public void setDocumentMetadata(DocumentMetadata documentMetadata) {
        this.documentMetadata = documentMetadata;
        // Sync with legacy fields for backward compatibility
        if (documentMetadata != null) {
            this.title = documentMetadata.getTitle();
            this.author = documentMetadata.getAuthor();
            this.fileType = documentMetadata.getFileType();
        }
    }
    
    public List<PageContent> getPages() {
        return pages;
    }
    
    public void setPages(List<PageContent> pages) {
        this.pages = pages;
    }
    
    public void addPage(PageContent page) {
        this.pages.add(page);
    }
    
    // Legacy getters and setters for backward compatibility
    public String getTitle() {
        return title != null ? title : (documentMetadata != null ? documentMetadata.getTitle() : null);
    }
    
    public void setTitle(String title) {
        this.title = title;
        if (documentMetadata != null) {
            documentMetadata.setTitle(title);
        }
    }
    
    public String getAuthor() {
        return author != null ? author : (documentMetadata != null ? documentMetadata.getAuthor() : null);
    }
    
    public void setAuthor(String author) {
        this.author = author;
        if (documentMetadata != null) {
            documentMetadata.setAuthor(author);
        }
    }
    
    public String getContent() {
        return content;
    }
    
    public void setContent(String content) {
        this.content = content;
    }
    
    public String getMarkdownContent() {
        return markdownContent;
    }
    
    public void setMarkdownContent(String markdownContent) {
        this.markdownContent = markdownContent;
    }
    
    public List<ParsedTable> getTables() {
        return tables;
    }
    
    public void setTables(List<ParsedTable> tables) {
        this.tables = tables;
    }
    
    public void addTable(ParsedTable table) {
        this.tables.add(table);
    }
    
    public List<ParsedImage> getImages() {
        return images;
    }
    
    public void setImages(List<ParsedImage> images) {
        this.images = images;
    }
    
    public void addImage(ParsedImage image) {
        this.images.add(image);
    }
    
    public Map<String, String> getMetadata() {
        return metadata;
    }
    
    public void setMetadata(Map<String, String> metadata) {
        this.metadata = metadata;
    }
    
    public void addMetadata(String key, String value) {
        this.metadata.put(key, value);
    }
    
    public FileType getFileType() {
        return fileType != null ? fileType : (documentMetadata != null ? documentMetadata.getFileType() : null);
    }
    
    public void setFileType(FileType fileType) {
        this.fileType = fileType;
        if (documentMetadata != null) {
            documentMetadata.setFileType(fileType);
        }
    }
    
    /**
     * File types that can be parsed
     */
    public enum FileType {
        PDF,
        WORD,
        EXCEL,
        POWERPOINT,
        CSV,
        TEXT,
        MARKDOWN,
        RTF,
        HTML,
        XML,
        JSON,
        UNKNOWN
    }
    
    /**
     * Represents a parsed table
     */
    public static class ParsedTable {
        private String title;
        private List<List<String>> data;
        private List<String> headers;
        
        public ParsedTable() {
            this.data = new ArrayList<>();
            this.headers = new ArrayList<>();
        }
        
        // Getters and Setters
        public String getTitle() {
            return title;
        }
        
        public void setTitle(String title) {
            this.title = title;
        }
        
        public List<List<String>> getData() {
            return data;
        }
        
        public void setData(List<List<String>> data) {
            this.data = data;
        }
        
        public List<String> getHeaders() {
            return headers;
        }
        
        public void setHeaders(List<String> headers) {
            this.headers = headers;
        }
        
        /**
         * Convert table to markdown format
         */
        public String toMarkdown() {
            StringBuilder sb = new StringBuilder();
            
            if (title != null && !title.isEmpty()) {
                sb.append("### ").append(title).append("\n\n");
            }
            
            // Headers
            if (!headers.isEmpty()) {
                sb.append("| ");
                for (String header : headers) {
                    sb.append(header).append(" | ");
                }
                sb.append("\n");
                
                // Separator
                sb.append("|");
                for (int i = 0; i < headers.size(); i++) {
                    sb.append("---|");
                }
                sb.append("\n");
            }
            
            // Data rows
            for (List<String> row : data) {
                sb.append("| ");
                for (String cell : row) {
                    sb.append(cell != null ? cell : "").append(" | ");
                }
                sb.append("\n");
            }
            
            return sb.toString();
        }
    }
    
    /**
     * Represents a parsed image
     */
    public static class ParsedImage {
        private String name;
        private String altText;
        private byte[] data;
        private String base64Data;
        private String url;
        private String mimeType;
        
        // Getters and Setters
        public String getName() {
            return name;
        }
        
        public void setName(String name) {
            this.name = name;
        }
        
        public String getAltText() {
            return altText;
        }
        
        public void setAltText(String altText) {
            this.altText = altText;
        }
        
        public byte[] getData() {
            return data;
        }
        
        public void setData(byte[] data) {
            this.data = data;
        }
        
        public String getBase64Data() {
            return base64Data;
        }
        
        public void setBase64Data(String base64Data) {
            this.base64Data = base64Data;
        }
        
        public String getUrl() {
            return url;
        }
        
        public void setUrl(String url) {
            this.url = url;
        }
        
        public String getMimeType() {
            return mimeType;
        }
        
        public void setMimeType(String mimeType) {
            this.mimeType = mimeType;
        }
        
        /**
         * Convert image to markdown format
         */
        public String toMarkdown() {
            String alt = altText != null ? altText : name;
            String src = url != null ? url : name;
            return "![" + alt + "](" + src + ")";
        }
    }
    
    /**
     * Get total word count from all pages or legacy content
     */
    public int getTotalWordCount() {
        if (!pages.isEmpty()) {
            return pages.stream()
                    .filter(page -> page.getWordCount() != null)
                    .mapToInt(PageContent::getWordCount)
                    .sum();
        }
        
        // Fallback to legacy content
        String textContent = content != null ? content : markdownContent;
        if (textContent != null && !textContent.trim().isEmpty()) {
            return textContent.trim().split("\\s+").length;
        }
        
        return 0;
    }
    
    /**
     * Get total character count from all pages or legacy content
     */
    public int getTotalCharacterCount() {
        if (!pages.isEmpty()) {
            return pages.stream()
                    .filter(page -> page.getCharacterCount() != null)
                    .mapToInt(PageContent::getCharacterCount)
                    .sum();
        }
        
        // Fallback to legacy content
        String textContent = content != null ? content : markdownContent;
        if (textContent != null) {
            return textContent.length();
        }
        
        return 0;
    }
    
    /**
     * Convert the entire document to markdown format
     */
    public String toMarkdown() {
        StringBuilder sb = new StringBuilder();
        
        String docTitle = getTitle();
        String docAuthor = getAuthor();
        
        // Title
        if (docTitle != null && !docTitle.isEmpty()) {
            sb.append("# ").append(docTitle).append("\n\n");
        }
        
        // Author
        if (docAuthor != null && !docAuthor.isEmpty()) {
            sb.append("**Author:** ").append(docAuthor).append("\n\n");
        }
        
        // Enhanced metadata
        if (documentMetadata != null) {
            addMetadataToMarkdown(sb);
        } else if (!metadata.isEmpty()) {
            // Fallback to legacy metadata
            sb.append("## Metadata\n\n");
            for (Map.Entry<String, String> entry : metadata.entrySet()) {
                sb.append("- **").append(entry.getKey()).append(":** ")
                  .append(entry.getValue()).append("\n");
            }
            sb.append("\n");
        }
        
        // Page-based content (preferred)
        if (!pages.isEmpty()) {
            for (PageContent page : pages) {
                if (page.hasContent()) {
                    sb.append(page.toMarkdown()).append("\n\n");
                }
            }
        } else {
            // Fallback to legacy content
            if (markdownContent != null && !markdownContent.isEmpty()) {
                sb.append(markdownContent).append("\n\n");
            } else if (content != null && !content.isEmpty()) {
                sb.append(content).append("\n\n");
            }
            
            // Legacy tables
            if (!tables.isEmpty()) {
                for (ParsedTable table : tables) {
                    sb.append(table.toMarkdown()).append("\n");
                }
            }
            
            // Legacy images
            if (!images.isEmpty()) {
                sb.append("\n## Images\n\n");
                for (ParsedImage image : images) {
                    sb.append(image.toMarkdown()).append("\n\n");
                }
            }
        }
        
        return sb.toString();
    }
    
    private void addMetadataToMarkdown(StringBuilder sb) {
        sb.append("## Document Information\n\n");
        
        if (documentMetadata.getFileSize() != null) {
            sb.append("- **File Size:** ").append(formatFileSize(documentMetadata.getFileSize())).append("\n");
        }
        if (documentMetadata.getTotalPages() != null) {
            sb.append("- **Total Pages:** ").append(documentMetadata.getTotalPages()).append("\n");
        }
        if (documentMetadata.getTotalWords() != null) {
            sb.append("- **Total Words:** ").append(documentMetadata.getTotalWords()).append("\n");
        }
        if (documentMetadata.getTotalCharacters() != null) {
            sb.append("- **Total Characters:** ").append(documentMetadata.getTotalCharacters()).append("\n");
        }
        if (documentMetadata.getSubject() != null) {
            sb.append("- **Subject:** ").append(documentMetadata.getSubject()).append("\n");
        }
        if (documentMetadata.getKeywords() != null) {
            sb.append("- **Keywords:** ").append(documentMetadata.getKeywords()).append("\n");
        }
        if (documentMetadata.getCreationDate() != null) {
            sb.append("- **Created:** ").append(documentMetadata.getCreationDate()).append("\n");
        }
        if (documentMetadata.getModificationDate() != null) {
            sb.append("- **Modified:** ").append(documentMetadata.getModificationDate()).append("\n");
        }
        
        sb.append("\n");
    }
    
    private String formatFileSize(long size) {
        if (size < 1024) return size + " B";
        if (size < 1024 * 1024) return String.format("%.2f KB", size / 1024.0);
        if (size < 1024 * 1024 * 1024) return String.format("%.2f MB", size / (1024.0 * 1024));
        return String.format("%.2f GB", size / (1024.0 * 1024 * 1024));
    }
}