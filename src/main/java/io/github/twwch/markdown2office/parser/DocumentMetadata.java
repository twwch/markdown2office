package io.github.twwch.markdown2office.parser;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Document metadata containing detailed information about the parsed file
 */
public class DocumentMetadata {
    
    // Basic file information
    private String fileName;
    private String filePath;
    private Long fileSize; // in bytes
    private String mimeType;
    private ParsedDocument.FileType fileType;
    
    // Document properties
    private String title;
    private String author;
    private String subject;
    private String keywords;
    private String description;
    private String creator;
    private String producer;
    private String language;
    
    // Time information
    private Date creationDate;
    private Date modificationDate;
    private Date lastAccessDate;
    
    // Page/Structure information
    private Integer totalPages;
    private Integer totalWords;
    private Integer totalCharacters;
    private Integer totalCharactersWithSpaces;
    private Integer totalParagraphs;
    private Integer totalLines;
    private Integer totalTables;
    private Integer totalImages;
    private Integer totalSections;
    private Integer totalSheets; // for Excel
    private Integer totalSlides; // for PowerPoint
    
    // Format information
    private String documentFormat; // PDF version, Word version, etc.
    private Boolean isEncrypted;
    private Boolean isPasswordProtected;
    private Boolean hasDigitalSignature;
    private Boolean hasComments;
    private Boolean hasFormFields;
    
    // Custom properties
    private Map<String, String> customProperties;
    
    public DocumentMetadata() {
        this.customProperties = new HashMap<>();
    }
    
    // Getters and Setters
    public String getFileName() {
        return fileName;
    }
    
    public void setFileName(String fileName) {
        this.fileName = fileName;
    }
    
    public String getFilePath() {
        return filePath;
    }
    
    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
    
    public Long getFileSize() {
        return fileSize;
    }
    
    public void setFileSize(Long fileSize) {
        this.fileSize = fileSize;
    }
    
    public String getMimeType() {
        return mimeType;
    }
    
    public void setMimeType(String mimeType) {
        this.mimeType = mimeType;
    }
    
    public ParsedDocument.FileType getFileType() {
        return fileType;
    }
    
    public void setFileType(ParsedDocument.FileType fileType) {
        this.fileType = fileType;
    }
    
    public String getTitle() {
        return title;
    }
    
    public void setTitle(String title) {
        this.title = title;
    }
    
    public String getAuthor() {
        return author;
    }
    
    public void setAuthor(String author) {
        this.author = author;
    }
    
    public String getSubject() {
        return subject;
    }
    
    public void setSubject(String subject) {
        this.subject = subject;
    }
    
    public String getKeywords() {
        return keywords;
    }
    
    public void setKeywords(String keywords) {
        this.keywords = keywords;
    }
    
    public String getDescription() {
        return description;
    }
    
    public void setDescription(String description) {
        this.description = description;
    }
    
    public String getCreator() {
        return creator;
    }
    
    public void setCreator(String creator) {
        this.creator = creator;
    }
    
    public String getProducer() {
        return producer;
    }
    
    public void setProducer(String producer) {
        this.producer = producer;
    }
    
    public String getLanguage() {
        return language;
    }
    
    public void setLanguage(String language) {
        this.language = language;
    }
    
    public Date getCreationDate() {
        return creationDate;
    }
    
    public void setCreationDate(Date creationDate) {
        this.creationDate = creationDate;
    }
    
    public Date getModificationDate() {
        return modificationDate;
    }
    
    public void setModificationDate(Date modificationDate) {
        this.modificationDate = modificationDate;
    }
    
    public Date getLastAccessDate() {
        return lastAccessDate;
    }
    
    public void setLastAccessDate(Date lastAccessDate) {
        this.lastAccessDate = lastAccessDate;
    }
    
    public Integer getTotalPages() {
        return totalPages;
    }
    
    public void setTotalPages(Integer totalPages) {
        this.totalPages = totalPages;
    }
    
    public Integer getTotalWords() {
        return totalWords;
    }
    
    public void setTotalWords(Integer totalWords) {
        this.totalWords = totalWords;
    }
    
    public Integer getTotalCharacters() {
        return totalCharacters;
    }
    
    public void setTotalCharacters(Integer totalCharacters) {
        this.totalCharacters = totalCharacters;
    }
    
    public Integer getTotalCharactersWithSpaces() {
        return totalCharactersWithSpaces;
    }
    
    public void setTotalCharactersWithSpaces(Integer totalCharactersWithSpaces) {
        this.totalCharactersWithSpaces = totalCharactersWithSpaces;
    }
    
    public Integer getTotalParagraphs() {
        return totalParagraphs;
    }
    
    public void setTotalParagraphs(Integer totalParagraphs) {
        this.totalParagraphs = totalParagraphs;
    }
    
    public Integer getTotalLines() {
        return totalLines;
    }
    
    public void setTotalLines(Integer totalLines) {
        this.totalLines = totalLines;
    }
    
    public Integer getTotalTables() {
        return totalTables;
    }
    
    public void setTotalTables(Integer totalTables) {
        this.totalTables = totalTables;
    }
    
    public Integer getTotalImages() {
        return totalImages;
    }
    
    public void setTotalImages(Integer totalImages) {
        this.totalImages = totalImages;
    }
    
    public Integer getTotalSections() {
        return totalSections;
    }
    
    public void setTotalSections(Integer totalSections) {
        this.totalSections = totalSections;
    }
    
    public Integer getTotalSheets() {
        return totalSheets;
    }
    
    public void setTotalSheets(Integer totalSheets) {
        this.totalSheets = totalSheets;
    }
    
    public Integer getTotalSlides() {
        return totalSlides;
    }
    
    public void setTotalSlides(Integer totalSlides) {
        this.totalSlides = totalSlides;
    }
    
    public String getDocumentFormat() {
        return documentFormat;
    }
    
    public void setDocumentFormat(String documentFormat) {
        this.documentFormat = documentFormat;
    }
    
    public Boolean getIsEncrypted() {
        return isEncrypted;
    }
    
    public void setIsEncrypted(Boolean isEncrypted) {
        this.isEncrypted = isEncrypted;
    }
    
    public Boolean getIsPasswordProtected() {
        return isPasswordProtected;
    }
    
    public void setIsPasswordProtected(Boolean isPasswordProtected) {
        this.isPasswordProtected = isPasswordProtected;
    }
    
    public Boolean getHasDigitalSignature() {
        return hasDigitalSignature;
    }
    
    public void setHasDigitalSignature(Boolean hasDigitalSignature) {
        this.hasDigitalSignature = hasDigitalSignature;
    }
    
    public Boolean getHasComments() {
        return hasComments;
    }
    
    public void setHasComments(Boolean hasComments) {
        this.hasComments = hasComments;
    }
    
    public Boolean getHasFormFields() {
        return hasFormFields;
    }
    
    public void setHasFormFields(Boolean hasFormFields) {
        this.hasFormFields = hasFormFields;
    }
    
    public Map<String, String> getCustomProperties() {
        return customProperties;
    }
    
    public void setCustomProperties(Map<String, String> customProperties) {
        this.customProperties = customProperties;
    }
    
    public void addCustomProperty(String key, String value) {
        this.customProperties.put(key, value);
    }
    
    /**
     * Convert metadata to a readable string format
     */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("=== Document Metadata ===\n");
        
        if (fileName != null) sb.append("File Name: ").append(fileName).append("\n");
        if (fileSize != null) sb.append("File Size: ").append(formatFileSize(fileSize)).append("\n");
        if (fileType != null) sb.append("File Type: ").append(fileType).append("\n");
        if (documentFormat != null) sb.append("Format: ").append(documentFormat).append("\n");
        
        if (title != null) sb.append("Title: ").append(title).append("\n");
        if (author != null) sb.append("Author: ").append(author).append("\n");
        if (subject != null) sb.append("Subject: ").append(subject).append("\n");
        
        if (totalPages != null) sb.append("Total Pages: ").append(totalPages).append("\n");
        if (totalWords != null) sb.append("Total Words: ").append(totalWords).append("\n");
        if (totalCharacters != null) sb.append("Total Characters: ").append(totalCharacters).append("\n");
        if (totalTables != null && totalTables > 0) sb.append("Total Tables: ").append(totalTables).append("\n");
        if (totalImages != null && totalImages > 0) sb.append("Total Images: ").append(totalImages).append("\n");
        if (totalSheets != null) sb.append("Total Sheets: ").append(totalSheets).append("\n");
        if (totalSlides != null) sb.append("Total Slides: ").append(totalSlides).append("\n");
        
        if (creationDate != null) sb.append("Created: ").append(creationDate).append("\n");
        if (modificationDate != null) sb.append("Modified: ").append(modificationDate).append("\n");
        
        return sb.toString();
    }
    
    /**
     * Format file size in human readable format
     */
    private String formatFileSize(long size) {
        if (size < 1024) return size + " B";
        if (size < 1024 * 1024) return String.format("%.2f KB", size / 1024.0);
        if (size < 1024 * 1024 * 1024) return String.format("%.2f MB", size / (1024.0 * 1024));
        return String.format("%.2f GB", size / (1024.0 * 1024 * 1024));
    }
}