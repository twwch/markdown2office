package io.github.twwch.markdown2office.parser.impl;

import io.github.twwch.markdown2office.parser.DocumentMetadata;
import io.github.twwch.markdown2office.parser.FileParser;
import io.github.twwch.markdown2office.parser.PageContent;
import io.github.twwch.markdown2office.parser.ParsedDocument;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.PDXObject;
import org.apache.pdfbox.pdmodel.graphics.form.PDFormXObject;
import org.apache.pdfbox.pdmodel.graphics.state.PDExtendedGraphicsState;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotation;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Parser for PDF files with optional hidden layer filtering
 */
public class PdfFileParser implements FileParser {
    
    // Configuration field to control whether to read hidden layers
    private boolean includeHiddenLayers = false;
    
    /**
     * Default constructor - hidden layers are excluded by default
     */
    public PdfFileParser() {
        this(false);
    }
    
    /**
     * Constructor with configuration
     * @param includeHiddenLayers if true, hidden layers will be included in extraction
     */
    public PdfFileParser(boolean includeHiddenLayers) {
        this.includeHiddenLayers = includeHiddenLayers;
    }
    
    /**
     * Set whether to include hidden layers in extraction
     * @param includeHiddenLayers true to include hidden layers, false to exclude them
     */
    public void setIncludeHiddenLayers(boolean includeHiddenLayers) {
        this.includeHiddenLayers = includeHiddenLayers;
    }
    
    /**
     * Get whether hidden layers are included in extraction
     * @return true if hidden layers are included, false otherwise
     */
    public boolean isIncludeHiddenLayers() {
        return includeHiddenLayers;
    }
    
    @Override
    public ParsedDocument parse(String filePath) throws IOException {
        return parse(new File(filePath));
    }
    
    @Override
    public ParsedDocument parse(File file) throws IOException {
        try (PDDocument document = PDDocument.load(file)) {
            // Remove hidden content if configured to do so
            if (!includeHiddenLayers) {
                removeHiddenContent(document);
            }
            
            ParsedDocument parsedDoc = extractContent(document, file.getName());
            // Set file size if available
            if (parsedDoc.getDocumentMetadata() != null) {
                parsedDoc.getDocumentMetadata().setFileSize(file.length());
            }
            return parsedDoc;
        }
    }
    
    @Override
    public ParsedDocument parse(InputStream inputStream, String fileName) throws IOException {
        try (PDDocument document = PDDocument.load(inputStream)) {
            // Remove hidden content if configured to do so
            if (!includeHiddenLayers) {
                removeHiddenContent(document);
            }
            
            return extractContent(document, fileName);
        }
    }
    
    @Override
    public boolean supports(String fileName) {
        return fileName != null && fileName.toLowerCase().endsWith(".pdf");
    }
    
    /**
     * Remove hidden content from PDF document
     */
    private void removeHiddenContent(PDDocument document) {
        for (PDPage page : document.getPages()) {
            try {
                // Remove hidden annotations
                List<PDAnnotation> annotations = page.getAnnotations();
                if (annotations != null) {
                    annotations.removeIf(annotation -> {
                        try {
                            // Remove hidden or watermark annotations
                            return annotation.isHidden() || 
                                   annotation.isNoView() ||
                                   (annotation.getAppearanceState() != null && 
                                    annotation.getAppearanceState().getName().toLowerCase().contains("watermark"));
                        } catch (Exception e) {
                            return false;
                        }
                    });
                }
                
                // Check and remove watermark layers from resources
                PDResources resources = page.getResources();
                if (resources != null) {
                    removeWatermarkFromResources(resources);
                }
                
            } catch (Exception e) {
                // Continue processing even if one page fails
                // Silently ignore to avoid disrupting the parsing process
            }
        }
    }
    
    /**
     * Remove watermark XObjects from resources
     */
    private void removeWatermarkFromResources(PDResources resources) throws IOException {
        if (resources == null) {
            return;
        }
        
        Set<COSName> watermarkNames = new HashSet<>();
        
        // Find XObjects that might be watermarks
        for (COSName name : resources.getXObjectNames()) {
            PDXObject xobject = resources.getXObject(name);
            
            if (xobject instanceof PDFormXObject) {
                PDFormXObject form = (PDFormXObject) xobject;
                
                // Check if this form has transparency or is marked as watermark
                PDResources formResources = form.getResources();
                if (formResources != null) {
                    // Check for transparency in extended graphics state
                    for (COSName gsName : formResources.getExtGStateNames()) {
                        PDExtendedGraphicsState gs = formResources.getExtGState(gsName);
                        if (gs != null) {
                            // If it has low opacity, it might be a watermark
                            Float ca = gs.getStrokingAlphaConstant();
                            Float ca2 = gs.getNonStrokingAlphaConstant();
                            
                            if ((ca != null && ca < 0.5f) || (ca2 != null && ca2 < 0.5f)) {
                                watermarkNames.add(name);
                                break;
                            }
                        }
                    }
                }
                
                // Check if the name suggests it's a watermark
                String nameStr = name.getName().toLowerCase();
                if (nameStr.contains("watermark") || 
                    nameStr.contains("wm") || 
                    nameStr.contains("background")) {
                    watermarkNames.add(name);
                }
            }
        }
        
        // Remove identified watermarks
        for (COSName name : watermarkNames) {
            resources.getCOSObject().removeItem(name);
        }
    }
    
    /**
     * Custom PDFTextStripper that filters invisible text when configured
     */
    private class FilteredTextStripper extends PDFTextStripper {
        
        public FilteredTextStripper() throws IOException {
            super();
            setSuppressDuplicateOverlappingText(true);
            setSortByPosition(true);
        }
        
        @Override
        protected void processTextPosition(TextPosition text) {
            // If including hidden layers, process everything
            if (includeHiddenLayers) {
                super.processTextPosition(text);
                return;
            }
            
            // Otherwise, filter out hidden content
            RenderingMode renderingMode = getGraphicsState().getTextState().getRenderingMode();
            
            // Skip invisible text (rendering mode 3)
            if (renderingMode == RenderingMode.NEITHER) {
                return;
            }
            
            // Check transparency - skip nearly transparent text
            float strokeAlpha = (float) getGraphicsState().getAlphaConstant();
            float fillAlpha = (float) getGraphicsState().getNonStrokeAlphaConstant();
            
            if (strokeAlpha < 0.3f || fillAlpha < 0.3f) {
                return;
            }
            
            // Skip white or very light colored text (often invisible on white background)
            float[] rgb = getGraphicsState().getNonStrokingColor().getComponents();
            if (rgb != null && rgb.length >= 3) {
                // If all color components are very high (close to white)
                if (rgb[0] > 0.95f && rgb[1] > 0.95f && rgb[2] > 0.95f) {
                    return;
                }
            }
            
            // Process normally if all checks pass
            super.processTextPosition(text);
        }
    }
    
    private ParsedDocument extractContent(PDDocument document, String fileName) throws IOException {
        ParsedDocument parsedDoc = new ParsedDocument();
        parsedDoc.setFileType(ParsedDocument.FileType.PDF);
        
        // Create and populate enhanced metadata
        DocumentMetadata metadata = new DocumentMetadata();
        metadata.setFileName(fileName);
        metadata.setFileType(ParsedDocument.FileType.PDF);
        metadata.setTotalPages(document.getNumberOfPages());
        
        // Extract PDF metadata
        PDDocumentInformation info = document.getDocumentInformation();
        if (info != null) {
            if (info.getTitle() != null) {
                metadata.setTitle(info.getTitle());
                parsedDoc.setTitle(info.getTitle());
            }
            if (info.getAuthor() != null) {
                metadata.setAuthor(info.getAuthor());
                parsedDoc.setAuthor(info.getAuthor());
            }
            if (info.getSubject() != null) {
                metadata.setSubject(info.getSubject());
            }
            if (info.getKeywords() != null) {
                metadata.setKeywords(info.getKeywords());
            }
            if (info.getCreator() != null) {
                metadata.setCreator(info.getCreator());
            }
            if (info.getProducer() != null) {
                metadata.setProducer(info.getProducer());
            }
            if (info.getCreationDate() != null) {
                metadata.setCreationDate(info.getCreationDate().getTime());
            }
            if (info.getModificationDate() != null) {
                metadata.setModificationDate(info.getModificationDate().getTime());
            }
        }
        
        // Extract content page by page for better structure preservation
        StringBuilder allContent = new StringBuilder();
        StringBuilder allMarkdown = new StringBuilder();
        int totalWords = 0;
        int totalChars = 0;
        
        // Use filtered text stripper if configured to exclude hidden layers
        PDFTextStripper textStripper = includeHiddenLayers ? 
            new PDFTextStripper() : new FilteredTextStripper();
        
        for (int pageNum = 1; pageNum <= document.getNumberOfPages(); pageNum++) {
            // Extract text for this specific page
            textStripper.setStartPage(pageNum);
            textStripper.setEndPage(pageNum);
            String pageText = textStripper.getText(document);
            
            if (pageText != null && !pageText.trim().isEmpty()) {
                PageContent pageContent = new PageContent(pageNum);
                pageContent.setRawText(pageText);
                
                // Convert page text to markdown with better formatting
                String pageMarkdown = convertPageToMarkdown(pageText, pageNum);
                pageContent.setMarkdownContent(pageMarkdown);
                
                // Extract structured content from the page
                extractPageStructure(pageText, pageContent);
                
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
        }
        
        // Update metadata with calculated statistics
        metadata.setTotalWords(totalWords);
        metadata.setTotalCharacters(totalChars);
        metadata.setTotalCharactersWithSpaces(allContent.toString().length());
        
        parsedDoc.setDocumentMetadata(metadata);
        
        // Set legacy content for backward compatibility
        parsedDoc.setContent(allContent.toString());
        parsedDoc.setMarkdownContent(allMarkdown.toString());
        
        // Add legacy metadata
        parsedDoc.addMetadata("Page Count", String.valueOf(document.getNumberOfPages()));
        parsedDoc.addMetadata("Word Count", String.valueOf(totalWords));
        parsedDoc.addMetadata("Character Count", String.valueOf(totalChars));
        
        return parsedDoc;
    }
    
    private String convertPageToMarkdown(String text, int pageNum) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        
        StringBuilder markdown = new StringBuilder();
        
        // Add page separator (except for first page)
        if (pageNum > 1) {
            markdown.append("---\n\n");
        }
        
        String[] lines = text.split("\n");
        boolean inParagraph = false;
        
        for (int i = 0; i < lines.length; i++) {
            String line = lines[i].trim();
            
            if (line.isEmpty()) {
                if (inParagraph) {
                    markdown.append("\n\n");
                    inParagraph = false;
                } else {
                    markdown.append("\n");
                }
                continue;
            }
            
            // Detect different types of content
            if (isHeading(line, i, lines)) {
                if (inParagraph) {
                    markdown.append("\n\n");
                    inParagraph = false;
                }
                
                int level = detectHeadingLevel(line);
                markdown.append("#".repeat(level)).append(" ").append(line).append("\n\n");
            } else if (isBulletPoint(line)) {
                if (inParagraph) {
                    markdown.append("\n\n");
                    inParagraph = false;
                }
                markdown.append("- ").append(cleanBulletPoint(line)).append("\n");
            } else if (isNumberedList(line)) {
                if (inParagraph) {
                    markdown.append("\n\n");
                    inParagraph = false;
                }
                markdown.append(line).append("\n");
            } else {
                // Regular paragraph text
                if (!inParagraph) {
                    inParagraph = true;
                } else {
                    markdown.append(" ");
                }
                markdown.append(line);
            }
        }
        
        if (inParagraph) {
            markdown.append("\n\n");
        }
        
        return markdown.toString();
    }
    
    private void extractPageStructure(String pageText, PageContent pageContent) {
        String[] lines = pageText.split("\n");
        
        for (int i = 0; i < lines.length; i++) {
            String line = lines[i].trim();
            
            if (line.isEmpty()) continue;
            
            if (isHeading(line, i, lines)) {
                pageContent.addHeading(line);
            } else if (isBulletPoint(line) || isNumberedList(line)) {
                pageContent.addList(line);
            } else if (line.length() > 20) { // Assume longer lines are paragraphs
                pageContent.addParagraph(line);
            }
        }
    }
    
    private boolean isHeading(String line, int lineIndex, String[] allLines) {
        // Multiple criteria for heading detection
        return (line.length() < 80 && // Not too long
               line.equals(line.toUpperCase()) && // All caps
               !line.matches(".*\\d{4}.*") && // No years
               !line.matches(".*\\$\\d+.*") && // No prices
               line.length() > 3) || // Not too short
               (line.length() < 60 &&
                lineIndex + 1 < allLines.length &&
                allLines[lineIndex + 1].trim().isEmpty()); // Followed by empty line
    }
    
    private int detectHeadingLevel(String line) {
        if (line.length() < 30 && line.equals(line.toUpperCase())) {
            return 2; // Major heading
        }
        return 3; // Minor heading
    }
    
    private boolean isBulletPoint(String line) {
        return line.matches("^[•\\*\\-\\+]\\s+.*") ||
               line.matches("^\\s*[•\\*\\-\\+]\\s+.*") ||
               line.matches("^\\s*[→►▪▫◦‣⁃]\\s+.*");
    }
    
    private boolean isNumberedList(String line) {
        return line.matches("^\\d+[.)]\\s+.*") ||
               line.matches("^\\s*\\d+[.)]\\s+.*") ||
               line.matches("^[a-zA-Z][.)]\\s+.*");
    }
    
    private String cleanBulletPoint(String line) {
        return line.replaceFirst("^[•\\*\\-\\+→►▪▫◦‣⁃]\\s*", "")
                  .replaceFirst("^\\s*[•\\*\\-\\+→►▪▫◦‣⁃]\\s*", "");
    }
}