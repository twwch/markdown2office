package io.github.twwch.markdown2office.converter;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import io.github.twwch.markdown2office.parser.MarkdownParser;
import org.commonmark.ext.gfm.tables.*;
import org.commonmark.node.*;
import org.commonmark.node.ListItem;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

public class PdfConverter implements Converter {
    
    private final MarkdownParser parser;
    private com.itextpdf.text.Document pdfDocument;
    private Font normalFont;
    private Font boldFont;
    private Font italicFont;
    private Font codeFont;
    private Font[] headingFonts;
    private int listLevel;
    private BaseFont chineseFont;
    
    public PdfConverter() {
        this.parser = new MarkdownParser();
        this.listLevel = 0;
        initFonts();
    }
    
    private void initFonts() {
        try {
            // Try to create a font that supports Chinese characters
            // Use iText's built-in Asian font support
            chineseFont = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        } catch (Exception e1) {
            try {
                // Fallback: try system font path on macOS
                chineseFont = BaseFont.createFont("/System/Library/Fonts/PingFang.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            } catch (Exception e2) {
                try {
                    // Second fallback: try Helvetica World font (includes more Unicode)
                    chineseFont = BaseFont.createFont("/System/Library/Fonts/Helvetica.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                } catch (Exception e3) {
                    // Final fallback: use default font
                    chineseFont = null;
                }
            }
        }
        
        if (chineseFont != null) {
            normalFont = new Font(chineseFont, 12, Font.NORMAL);
            boldFont = new Font(chineseFont, 12, Font.BOLD);
            italicFont = new Font(chineseFont, 12, Font.ITALIC);
            codeFont = new Font(chineseFont, 11, Font.NORMAL, BaseColor.DARK_GRAY);
            
            headingFonts = new Font[6];
            headingFonts[0] = new Font(chineseFont, 24, Font.BOLD);
            headingFonts[1] = new Font(chineseFont, 20, Font.BOLD);
            headingFonts[2] = new Font(chineseFont, 18, Font.BOLD);
            headingFonts[3] = new Font(chineseFont, 16, Font.BOLD);
            headingFonts[4] = new Font(chineseFont, 14, Font.BOLD);
            headingFonts[5] = new Font(chineseFont, 13, Font.BOLD);
        } else {
            // Use default fonts if Chinese font is not available
            normalFont = new Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL);
            boldFont = new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD);
            italicFont = new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC);
            codeFont = new Font(Font.FontFamily.COURIER, 11, Font.NORMAL, BaseColor.DARK_GRAY);
            
            headingFonts = new Font[6];
            headingFonts[0] = new Font(Font.FontFamily.HELVETICA, 24, Font.BOLD);
            headingFonts[1] = new Font(Font.FontFamily.HELVETICA, 20, Font.BOLD);
            headingFonts[2] = new Font(Font.FontFamily.HELVETICA, 18, Font.BOLD);
            headingFonts[3] = new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD);
            headingFonts[4] = new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD);
            headingFonts[5] = new Font(Font.FontFamily.HELVETICA, 13, Font.BOLD);
        }
    }
    
    @Override
    public void convert(String markdown, OutputStream outputStream) throws IOException {
        try {
            pdfDocument = new com.itextpdf.text.Document(PageSize.A4, 50, 50, 50, 50);
            PdfWriter.getInstance(pdfDocument, outputStream);
            pdfDocument.open();
            
            Node node = parser.parse(markdown);
            processNode(node);
            
            pdfDocument.close();
        } catch (DocumentException e) {
            throw new IOException("Error creating PDF document", e);
        }
    }
    
    private void processNode(Node node) throws DocumentException {
        if (node instanceof org.commonmark.node.Document) {
            Node child = node.getFirstChild();
            while (child != null) {
                processNode(child);
                child = child.getNext();
            }
        } else if (node instanceof Heading) {
            processHeading((Heading) node);
        } else if (node instanceof org.commonmark.node.Paragraph) {
            processParagraph((org.commonmark.node.Paragraph) node);
        } else if (node instanceof BulletList) {
            processList((BulletList) node, false);
        } else if (node instanceof OrderedList) {
            processList((OrderedList) node, true);
        } else if (node instanceof BlockQuote) {
            processBlockQuote((BlockQuote) node);
        } else if (node instanceof FencedCodeBlock) {
            processCodeBlock((FencedCodeBlock) node);
        } else if (node instanceof IndentedCodeBlock) {
            processCodeBlock((IndentedCodeBlock) node);
        } else if (node instanceof TableBlock) {
            processTable((TableBlock) node);
        } else if (node instanceof ThematicBreak) {
            pdfDocument.add(new com.itextpdf.text.pdf.draw.LineSeparator());
        } else if (node instanceof HtmlBlock) {
            com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph(((HtmlBlock) node).getLiteral(), normalFont);
            pdfDocument.add(p);
        }
    }
    
    private void processHeading(Heading heading) throws DocumentException {
        int level = heading.getLevel() - 1;
        if (level < 0) level = 0;
        if (level > 5) level = 5;
        
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph();
        p.setFont(headingFonts[level]);
        p.setSpacingBefore(10);
        p.setSpacingAfter(10);
        p.setLeading(headingFonts[level].getSize() * 1.5f); // Set line height
        
        addInlineContent(heading.getFirstChild(), p, headingFonts[level]);
        pdfDocument.add(p);
    }
    
    private void processParagraph(org.commonmark.node.Paragraph paragraph) throws DocumentException {
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph();
        p.setSpacingAfter(10);
        p.setLeading(18f); // Set line height for better readability
        p.setAlignment(com.itextpdf.text.Element.ALIGN_JUSTIFIED_ALL); // Justify text for better spacing
        addInlineContent(paragraph.getFirstChild(), p, normalFont);
        pdfDocument.add(p);
    }
    
    private void processList(ListBlock listBlock, boolean ordered) throws DocumentException {
        listLevel++;
        com.itextpdf.text.List list = new com.itextpdf.text.List(ordered);
        list.setIndentationLeft(20 * listLevel);
        
        Node item = listBlock.getFirstChild();
        while (item != null) {
            if (item instanceof ListItem) {
                StringBuilder itemText = new StringBuilder();
                extractText(item.getFirstChild(), itemText);
                
                com.itextpdf.text.ListItem listItem = new com.itextpdf.text.ListItem(itemText.toString(), normalFont);
                listItem.setLeading(18f); // Set line height for list items
                list.add(listItem);
                
                Node child = item.getFirstChild();
                while (child != null) {
                    if (child instanceof ListBlock) {
                        pdfDocument.add(list);
                        processNode(child);
                        list = new com.itextpdf.text.List(ordered);
                        list.setIndentationLeft(20 * listLevel);
                    }
                    child = child.getNext();
                }
            }
            item = item.getNext();
        }
        
        pdfDocument.add(list);
        listLevel--;
    }
    
    private void processBlockQuote(BlockQuote blockQuote) throws DocumentException {
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph();
        p.setIndentationLeft(30);
        p.setSpacingBefore(10);
        p.setSpacingAfter(10);
        p.setLeading(18f); // Set line height for block quotes
        
        Font quoteFont = chineseFont != null ? new Font(chineseFont, 12, Font.NORMAL, BaseColor.GRAY) : new Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL, BaseColor.GRAY);
        Chunk quoteBar = new Chunk("│ ", quoteFont);
        p.add(quoteBar);
        
        Node child = blockQuote.getFirstChild();
        while (child != null) {
            if (child instanceof org.commonmark.node.Paragraph) {
                addInlineContent(child.getFirstChild(), p, italicFont);
            }
            child = child.getNext();
        }
        
        pdfDocument.add(p);
    }
    
    private void processCodeBlock(FencedCodeBlock codeBlock) throws DocumentException {
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph();
        p.setIndentationLeft(20);
        p.setSpacingBefore(10);
        p.setSpacingAfter(10);
        p.setLeading(16f); // Set line height for code blocks
        
        // Add code with character spacing
        Chunk codeChunk = new Chunk(codeBlock.getLiteral(), codeFont);
        codeChunk.setCharacterSpacing(0.15f); // Add character spacing for code
        p.add(codeChunk);
        
        pdfDocument.add(p);
    }
    
    private void processCodeBlock(IndentedCodeBlock codeBlock) throws DocumentException {
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph();
        p.setIndentationLeft(20);
        p.setSpacingBefore(10);
        p.setSpacingAfter(10);
        p.setLeading(16f); // Set line height for code blocks
        
        // Add code with character spacing
        Chunk codeChunk = new Chunk(codeBlock.getLiteral(), codeFont);
        codeChunk.setCharacterSpacing(0.15f); // Add character spacing for code
        p.add(codeChunk);
        
        pdfDocument.add(p);
    }
    
    private void processTable(TableBlock tableBlock) throws DocumentException {
        java.util.List<java.util.List<String>> tableData = new ArrayList<>();
        boolean hasHeader = false;
        
        Node child = tableBlock.getFirstChild();
        while (child != null) {
            if (child instanceof TableHead) {
                hasHeader = true;
                processTableSection(child, tableData);
            } else if (child instanceof TableBody) {
                processTableSection(child, tableData);
            }
            child = child.getNext();
        }
        
        if (!tableData.isEmpty()) {
            PdfPTable table = new PdfPTable(tableData.get(0).size());
            table.setWidthPercentage(100);
            table.setSpacingBefore(10);
            table.setSpacingAfter(10);
            
            for (int i = 0; i < tableData.size(); i++) {
                java.util.List<String> rowData = tableData.get(i);
                for (String cellText : rowData) {
                    // Create phrase with character spacing
                    Phrase phrase = new Phrase();
                    Chunk chunk = new Chunk(cellText, (hasHeader && i == 0) ? boldFont : normalFont);
                    chunk.setCharacterSpacing(0.2f); // Add character spacing in tables
                    phrase.add(chunk);
                    
                    PdfPCell cell = new PdfPCell(phrase);
                    cell.setPadding(5);
                    
                    if (hasHeader && i == 0) {
                        cell.setBackgroundColor(BaseColor.LIGHT_GRAY);
                    }
                    
                    table.addCell(cell);
                }
            }
            
            pdfDocument.add(table);
        }
    }
    
    private void processTableSection(Node section, java.util.List<java.util.List<String>> tableData) {
        Node row = section.getFirstChild();
        while (row != null) {
            if (row instanceof TableRow) {
                java.util.List<String> rowData = new ArrayList<>();
                Node cell = row.getFirstChild();
                while (cell != null) {
                    if (cell instanceof TableCell) {
                        StringBuilder cellText = new StringBuilder();
                        extractText(cell, cellText);
                        rowData.add(cellText.toString().trim());
                    }
                    cell = cell.getNext();
                }
                tableData.add(rowData);
            }
            row = row.getNext();
        }
    }
    
    private void addInlineContent(Node node, com.itextpdf.text.Paragraph paragraph, Font defaultFont) {
        while (node != null) {
            if (node instanceof Text) {
                String text = ((Text) node).getLiteral();
                // Add small space between characters for better readability
                Chunk chunk = new Chunk(text, defaultFont);
                chunk.setCharacterSpacing(0.2f); // Add character spacing
                paragraph.add(chunk);
            } else if (node instanceof Emphasis) {
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                Chunk chunk = new Chunk(text.toString(), italicFont);
                chunk.setCharacterSpacing(0.2f);
                paragraph.add(chunk);
            } else if (node instanceof StrongEmphasis) {
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                Chunk chunk = new Chunk(text.toString(), boldFont);
                chunk.setCharacterSpacing(0.2f);
                paragraph.add(chunk);
            } else if (node instanceof Code) {
                Chunk chunk = new Chunk(((Code) node).getLiteral(), codeFont);
                chunk.setCharacterSpacing(0.15f); // Slightly less spacing for code
                paragraph.add(chunk);
            } else if (node instanceof Link) {
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                Font linkFont = chineseFont != null ? new Font(chineseFont, 12, Font.UNDERLINE, BaseColor.BLUE) : new Font(Font.FontFamily.HELVETICA, 12, Font.UNDERLINE, BaseColor.BLUE);
                Chunk linkChunk = new Chunk(text.toString(), linkFont);
                linkChunk.setAnchor(((Link) node).getDestination());
                linkChunk.setCharacterSpacing(0.2f);
                paragraph.add(linkChunk);
            } else if (node instanceof org.commonmark.node.Image) {
                paragraph.add(new Chunk("[Image: " + ((org.commonmark.node.Image) node).getTitle() + "]", defaultFont));
            } else if (node instanceof HardLineBreak || node instanceof SoftLineBreak) {
                paragraph.add(Chunk.NEWLINE);
            } else {
                addInlineContent(node.getFirstChild(), paragraph, defaultFont);
            }
            node = node.getNext();
        }
    }
    
    private void extractText(Node node, StringBuilder text) {
        if (node == null) return;
        
        if (node instanceof Text) {
            text.append(((Text) node).getLiteral());
        } else if (node instanceof Code) {
            text.append(((Code) node).getLiteral());
        } else if (node instanceof HardLineBreak || node instanceof SoftLineBreak) {
            text.append("\n");
        }
        
        Node child = node.getFirstChild();
        while (child != null) {
            extractText(child, text);
            child = child.getNext();
        }
    }
}