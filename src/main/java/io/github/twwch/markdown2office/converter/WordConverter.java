package io.github.twwch.markdown2office.converter;

import io.github.twwch.markdown2office.parser.MarkdownParser;
import org.apache.poi.xwpf.usermodel.*;
import org.commonmark.ext.gfm.tables.TableBlock;
import org.commonmark.ext.gfm.tables.TableBody;
import org.commonmark.ext.gfm.tables.TableCell;
import org.commonmark.ext.gfm.tables.TableHead;
import org.commonmark.ext.gfm.tables.TableRow;
import org.commonmark.node.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class WordConverter implements Converter {
    
    private final MarkdownParser parser;
    private XWPFDocument document;
    private XWPFParagraph currentParagraph;
    private int listLevel;
    
    public WordConverter() {
        this.parser = new MarkdownParser();
        this.listLevel = 0;
    }
    
    @Override
    public void convert(String markdown, OutputStream outputStream) throws IOException {
        document = new XWPFDocument();
        Node node = parser.parse(markdown);
        processNode(node);
        document.write(outputStream);
        outputStream.close();
    }
    
    private void processNode(Node node) {
        if (node instanceof org.commonmark.node.Document) {
            Node child = node.getFirstChild();
            while (child != null) {
                processNode(child);
                child = child.getNext();
            }
        } else if (node instanceof Heading) {
            processHeading((Heading) node);
        } else if (node instanceof Paragraph) {
            processParagraph((Paragraph) node);
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
            currentParagraph = document.createParagraph();
            currentParagraph.createRun().addBreak();
        } else if (node instanceof HtmlBlock) {
            processHtmlBlock((HtmlBlock) node);
        } else {
            Node child = node.getFirstChild();
            while (child != null) {
                processNode(child);
                child = child.getNext();
            }
        }
    }
    
    private void processHeading(Heading heading) {
        currentParagraph = document.createParagraph();
        int level = heading.getLevel();
        
        switch (level) {
            case 1:
                currentParagraph.setStyle("Heading1");
                break;
            case 2:
                currentParagraph.setStyle("Heading2");
                break;
            case 3:
                currentParagraph.setStyle("Heading3");
                break;
            case 4:
                currentParagraph.setStyle("Heading4");
                break;
            case 5:
                currentParagraph.setStyle("Heading5");
                break;
            case 6:
                currentParagraph.setStyle("Heading6");
                break;
        }
        
        processInlineContent(heading.getFirstChild());
    }
    
    private void processParagraph(Paragraph paragraph) {
        currentParagraph = document.createParagraph();
        processInlineContent(paragraph.getFirstChild());
    }
    
    private void processList(ListBlock listBlock, boolean ordered) {
        listLevel++;
        
        Node item = listBlock.getFirstChild();
        while (item != null) {
            if (item instanceof ListItem) {
                currentParagraph = document.createParagraph();
                currentParagraph.setIndentationLeft(400 * listLevel);
                
                if (ordered) {
                    currentParagraph.setNumID(BigInteger.valueOf(1));
                } else {
                    XWPFRun run = currentParagraph.createRun();
                    run.setText("• ");
                }
                
                Node child = item.getFirstChild();
                while (child != null) {
                    if (child instanceof Paragraph) {
                        processInlineContent(child.getFirstChild());
                    } else {
                        processNode(child);
                    }
                    child = child.getNext();
                }
            }
            item = item.getNext();
        }
        
        listLevel--;
    }
    
    private void processBlockQuote(BlockQuote blockQuote) {
        currentParagraph = document.createParagraph();
        currentParagraph.setIndentationLeft(720);
        currentParagraph.setBorderLeft(Borders.SINGLE);
        
        Node child = blockQuote.getFirstChild();
        while (child != null) {
            if (child instanceof Paragraph) {
                processInlineContent(child.getFirstChild());
            } else {
                processNode(child);
            }
            child = child.getNext();
        }
    }
    
    private void processCodeBlock(FencedCodeBlock codeBlock) {
        currentParagraph = document.createParagraph();
        currentParagraph.setBorderTop(Borders.SINGLE);
        currentParagraph.setBorderBottom(Borders.SINGLE);
        currentParagraph.setBorderLeft(Borders.SINGLE);
        currentParagraph.setBorderRight(Borders.SINGLE);
        
        XWPFRun run = currentParagraph.createRun();
        run.setFontFamily("Courier New");
        run.setFontSize(10);
        run.setText(codeBlock.getLiteral());
    }
    
    private void processCodeBlock(IndentedCodeBlock codeBlock) {
        currentParagraph = document.createParagraph();
        currentParagraph.setBorderTop(Borders.SINGLE);
        currentParagraph.setBorderBottom(Borders.SINGLE);
        currentParagraph.setBorderLeft(Borders.SINGLE);
        currentParagraph.setBorderRight(Borders.SINGLE);
        
        XWPFRun run = currentParagraph.createRun();
        run.setFontFamily("Courier New");
        run.setFontSize(10);
        run.setText(codeBlock.getLiteral());
    }
    
    private void processTable(TableBlock tableBlock) {
        List<List<String>> tableData = new ArrayList<>();
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
            XWPFTable table = document.createTable(tableData.size(), tableData.get(0).size());
            CTTblWidth width = table.getCTTbl().addNewTblPr().addNewTblW();
            width.setType(STTblWidth.DXA);
            width.setW(BigInteger.valueOf(9072));
            
            for (int i = 0; i < tableData.size(); i++) {
                XWPFTableRow row = table.getRow(i);
                List<String> rowData = tableData.get(i);
                
                for (int j = 0; j < rowData.size(); j++) {
                    XWPFTableCell cell = row.getCell(j);
                    cell.setText(rowData.get(j));
                    
                    if (hasHeader && i == 0) {
                        XWPFParagraph p = cell.getParagraphs().get(0);
                        XWPFRun r = p.getRuns().get(0);
                        r.setBold(true);
                    }
                }
            }
        }
    }
    
    private void processTableSection(Node section, List<List<String>> tableData) {
        Node row = section.getFirstChild();
        while (row != null) {
            if (row instanceof TableRow) {
                List<String> rowData = new ArrayList<>();
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
    
    private void processHtmlBlock(HtmlBlock htmlBlock) {
        currentParagraph = document.createParagraph();
        XWPFRun run = currentParagraph.createRun();
        run.setText(htmlBlock.getLiteral());
    }
    
    private void processInlineContent(Node node) {
        while (node != null) {
            if (node instanceof Text) {
                XWPFRun run = currentParagraph.createRun();
                run.setText(((Text) node).getLiteral());
            } else if (node instanceof Emphasis) {
                XWPFRun run = currentParagraph.createRun();
                run.setItalic(true);
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                run.setText(text.toString());
            } else if (node instanceof StrongEmphasis) {
                XWPFRun run = currentParagraph.createRun();
                run.setBold(true);
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                run.setText(text.toString());
            } else if (node instanceof Code) {
                XWPFRun run = currentParagraph.createRun();
                run.setFontFamily("Courier New");
                run.setText(((Code) node).getLiteral());
            } else if (node instanceof Link) {
                XWPFRun run = currentParagraph.createRun();
                run.setUnderline(UnderlinePatterns.SINGLE);
                run.setColor("0000FF");
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                run.setText(text.toString());
            } else if (node instanceof Image) {
                XWPFRun run = currentParagraph.createRun();
                run.setText("[Image: " + ((Image) node).getTitle() + "]");
            } else if (node instanceof HardLineBreak || node instanceof SoftLineBreak) {
                currentParagraph.createRun().addBreak();
            } else {
                processInlineContent(node.getFirstChild());
            }
            node = node.getNext();
        }
    }
    
    private void extractText(Node node, StringBuilder text) {
        if (node instanceof Text) {
            text.append(((Text) node).getLiteral());
        }
        Node child = node.getFirstChild();
        while (child != null) {
            extractText(child, text);
            child = child.getNext();
        }
    }
}