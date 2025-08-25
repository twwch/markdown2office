package io.github.twwch.markdown2office.converter;

import com.itextpdf.text.*;
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
    
    public PdfConverter() {
        this.parser = new MarkdownParser();
        this.listLevel = 0;
        initFonts();
    }
    
    private void initFonts() {
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
        
        addInlineContent(heading.getFirstChild(), p, headingFonts[level]);
        pdfDocument.add(p);
    }
    
    private void processParagraph(org.commonmark.node.Paragraph paragraph) throws DocumentException {
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph();
        p.setSpacingAfter(10);
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
        
        Chunk quoteBar = new Chunk("│ ", new Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL, BaseColor.GRAY));
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
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph(codeBlock.getLiteral(), codeFont);
        p.setIndentationLeft(20);
        p.setSpacingBefore(10);
        p.setSpacingAfter(10);
        pdfDocument.add(p);
    }
    
    private void processCodeBlock(IndentedCodeBlock codeBlock) throws DocumentException {
        com.itextpdf.text.Paragraph p = new com.itextpdf.text.Paragraph(codeBlock.getLiteral(), codeFont);
        p.setIndentationLeft(20);
        p.setSpacingBefore(10);
        p.setSpacingAfter(10);
        // Background color not directly supported for paragraphs in iText5
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
                    PdfPCell cell = new PdfPCell(new Phrase(cellText, (hasHeader && i == 0) ? boldFont : normalFont));
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
                paragraph.add(new Chunk(((Text) node).getLiteral(), defaultFont));
            } else if (node instanceof Emphasis) {
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                paragraph.add(new Chunk(text.toString(), italicFont));
            } else if (node instanceof StrongEmphasis) {
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                paragraph.add(new Chunk(text.toString(), boldFont));
            } else if (node instanceof Code) {
                paragraph.add(new Chunk(((Code) node).getLiteral(), codeFont));
            } else if (node instanceof Link) {
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                Chunk linkChunk = new Chunk(text.toString(), new Font(Font.FontFamily.HELVETICA, 12, Font.UNDERLINE, BaseColor.BLUE));
                linkChunk.setAnchor(((Link) node).getDestination());
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