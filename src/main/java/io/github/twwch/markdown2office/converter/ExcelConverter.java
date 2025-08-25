package io.github.twwch.markdown2office.converter;

import io.github.twwch.markdown2office.parser.MarkdownParser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.commonmark.ext.gfm.tables.*;
import org.commonmark.node.*;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelConverter implements Converter {
    
    private final MarkdownParser parser;
    private Workbook workbook;
    private Sheet sheet;
    private int currentRow;
    private CellStyle headerStyle;
    private CellStyle codeStyle;
    private CellStyle boldStyle;
    private CellStyle italicStyle;
    
    public ExcelConverter() {
        this.parser = new MarkdownParser();
    }
    
    @Override
    public void convert(String markdown, OutputStream outputStream) throws IOException {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Markdown Content");
        currentRow = 0;
        
        initStyles();
        
        Node node = parser.parse(markdown);
        processNode(node, 0);
        
        autoSizeColumns();
        
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
    
    private void initStyles() {
        headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerStyle.setFont(headerFont);
        
        codeStyle = workbook.createCellStyle();
        Font codeFont = workbook.createFont();
        codeFont.setFontName("Courier New");
        codeStyle.setFont(codeFont);
        codeStyle.setWrapText(true);
        
        boldStyle = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);
        
        italicStyle = workbook.createCellStyle();
        Font italicFont = workbook.createFont();
        italicFont.setItalic(true);
        italicStyle.setFont(italicFont);
    }
    
    private void processNode(Node node, int indent) {
        if (node instanceof Document) {
            Node child = node.getFirstChild();
            while (child != null) {
                processNode(child, indent);
                child = child.getNext();
            }
        } else if (node instanceof Heading) {
            processHeading((Heading) node, indent);
        } else if (node instanceof Paragraph) {
            processParagraph((Paragraph) node, indent);
        } else if (node instanceof BulletList) {
            processList((BulletList) node, indent, false);
        } else if (node instanceof OrderedList) {
            processList((OrderedList) node, indent, true);
        } else if (node instanceof BlockQuote) {
            processBlockQuote((BlockQuote) node, indent);
        } else if (node instanceof FencedCodeBlock) {
            processCodeBlock((FencedCodeBlock) node, indent);
        } else if (node instanceof IndentedCodeBlock) {
            processCodeBlock((IndentedCodeBlock) node, indent);
        } else if (node instanceof TableBlock) {
            processTable((TableBlock) node, indent);
        } else if (node instanceof ThematicBreak) {
            Row row = sheet.createRow(currentRow++);
            Cell cell = row.createCell(indent);
            cell.setCellValue("---");
        } else if (node instanceof HtmlBlock) {
            Row row = sheet.createRow(currentRow++);
            Cell cell = row.createCell(indent);
            cell.setCellValue(((HtmlBlock) node).getLiteral());
        }
    }
    
    private void processHeading(Heading heading, int indent) {
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(indent);
        
        StringBuilder text = new StringBuilder();
        for (int i = 0; i < heading.getLevel(); i++) {
            text.append("#");
        }
        text.append(" ");
        extractText(heading.getFirstChild(), text);
        
        cell.setCellValue(text.toString());
        cell.setCellStyle(headerStyle);
    }
    
    private void processParagraph(Paragraph paragraph, int indent) {
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(indent);
        
        StringBuilder text = new StringBuilder();
        extractText(paragraph.getFirstChild(), text);
        cell.setCellValue(text.toString());
    }
    
    private void processList(ListBlock listBlock, int indent, boolean ordered) {
        int itemNumber = 1;
        Node item = listBlock.getFirstChild();
        
        while (item != null) {
            if (item instanceof ListItem) {
                Row row = sheet.createRow(currentRow++);
                Cell cell = row.createCell(indent);
                
                StringBuilder text = new StringBuilder();
                if (ordered) {
                    text.append(itemNumber).append(". ");
                } else {
                    text.append("• ");
                }
                
                Node child = item.getFirstChild();
                while (child != null) {
                    if (child instanceof Paragraph) {
                        extractText(child.getFirstChild(), text);
                    } else if (child instanceof ListBlock) {
                        cell.setCellValue(text.toString());
                        processNode(child, indent + 1);
                        text = new StringBuilder();
                    }
                    child = child.getNext();
                }
                
                if (text.length() > 0) {
                    cell.setCellValue(text.toString());
                }
                itemNumber++;
            }
            item = item.getNext();
        }
    }
    
    private void processBlockQuote(BlockQuote blockQuote, int indent) {
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(indent);
        
        StringBuilder text = new StringBuilder("> ");
        Node child = blockQuote.getFirstChild();
        while (child != null) {
            if (child instanceof Paragraph) {
                extractText(child.getFirstChild(), text);
            }
            child = child.getNext();
        }
        
        cell.setCellValue(text.toString());
        cell.setCellStyle(italicStyle);
    }
    
    private void processCodeBlock(FencedCodeBlock codeBlock, int indent) {
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(indent);
        cell.setCellValue(codeBlock.getLiteral());
        cell.setCellStyle(codeStyle);
    }
    
    private void processCodeBlock(IndentedCodeBlock codeBlock, int indent) {
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(indent);
        cell.setCellValue(codeBlock.getLiteral());
        cell.setCellStyle(codeStyle);
    }
    
    private void processTable(TableBlock tableBlock, int indent) {
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
        
        for (int i = 0; i < tableData.size(); i++) {
            Row row = sheet.createRow(currentRow++);
            List<String> rowData = tableData.get(i);
            
            for (int j = 0; j < rowData.size(); j++) {
                Cell cell = row.createCell(indent + j);
                cell.setCellValue(rowData.get(j));
                
                if (hasHeader && i == 0) {
                    cell.setCellStyle(headerStyle);
                }
            }
        }
        
        currentRow++;
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
    
    private void autoSizeColumns() {
        int maxColumn = 0;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getLastCellNum() > maxColumn) {
                maxColumn = row.getLastCellNum();
            }
        }
        
        for (int i = 0; i < maxColumn; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}