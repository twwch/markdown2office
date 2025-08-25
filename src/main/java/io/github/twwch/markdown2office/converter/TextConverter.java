package io.github.twwch.markdown2office.converter;

import io.github.twwch.markdown2office.parser.MarkdownParser;
import org.commonmark.ext.gfm.tables.*;
import org.commonmark.node.*;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

public class TextConverter implements Converter {
    
    private final MarkdownParser parser;
    private Writer writer;
    private int listLevel;
    private int listItemNumber;
    
    public TextConverter() {
        this.parser = new MarkdownParser();
        this.listLevel = 0;
    }
    
    @Override
    public void convert(String markdown, OutputStream outputStream) throws IOException {
        writer = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8);
        Node node = parser.parse(markdown);
        processNode(node);
        writer.flush();
        writer.close();
    }
    
    private void processNode(Node node) throws IOException {
        if (node instanceof Document) {
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
            writer.write("\n---\n\n");
        } else if (node instanceof HtmlBlock) {
            writer.write(((HtmlBlock) node).getLiteral());
            writer.write("\n\n");
        }
    }
    
    private void processHeading(Heading heading) throws IOException {
        StringBuilder text = new StringBuilder();
        extractText(heading.getFirstChild(), text);
        
        String headingText = text.toString().toUpperCase();
        writer.write("\n");
        writer.write(headingText);
        writer.write("\n");
        
        for (int i = 0; i < headingText.length(); i++) {
            writer.write(heading.getLevel() == 1 ? "=" : "-");
        }
        writer.write("\n\n");
    }
    
    private void processParagraph(Paragraph paragraph) throws IOException {
        StringBuilder text = new StringBuilder();
        Node child = paragraph.getFirstChild();
        while (child != null) {
            extractText(child, text);
            child = child.getNext();
        }
        writer.write(text.toString());
        writer.write("\n\n");
    }
    
    private void processList(ListBlock listBlock, boolean ordered) throws IOException {
        listLevel++;
        listItemNumber = 1;
        
        Node item = listBlock.getFirstChild();
        while (item != null) {
            if (item instanceof ListItem) {
                writeIndent();
                
                if (ordered) {
                    writer.write(String.valueOf(listItemNumber++));
                    writer.write(". ");
                } else {
                    writer.write("* ");
                }
                
                Node child = item.getFirstChild();
                boolean first = true;
                while (child != null) {
                    if (child instanceof Paragraph) {
                        if (!first) {
                            writeIndent();
                            writer.write("  ");
                        }
                        StringBuilder text = new StringBuilder();
                        extractText(child.getFirstChild(), text);
                        writer.write(text.toString().trim());
                        first = false;
                    } else if (child instanceof ListBlock) {
                        writer.write("\n");
                        processNode(child);
                    }
                    child = child.getNext();
                }
                writer.write("\n");
            }
            item = item.getNext();
        }
        
        listLevel--;
        if (listLevel == 0) {
            writer.write("\n");
        }
    }
    
    private void processBlockQuote(BlockQuote blockQuote) throws IOException {
        Node child = blockQuote.getFirstChild();
        while (child != null) {
            writer.write("> ");
            if (child instanceof Paragraph) {
                StringBuilder text = new StringBuilder();
                extractText(child.getFirstChild(), text);
                writer.write(text.toString().trim());
            }
            writer.write("\n");
            child = child.getNext();
        }
        writer.write("\n");
    }
    
    private void processCodeBlock(FencedCodeBlock codeBlock) throws IOException {
        String[] lines = codeBlock.getLiteral().split("\n");
        for (String line : lines) {
            writer.write("    ");
            writer.write(line);
            writer.write("\n");
        }
        writer.write("\n");
    }
    
    private void processCodeBlock(IndentedCodeBlock codeBlock) throws IOException {
        String[] lines = codeBlock.getLiteral().split("\n");
        for (String line : lines) {
            writer.write("    ");
            writer.write(line);
            writer.write("\n");
        }
        writer.write("\n");
    }
    
    private void processTable(TableBlock tableBlock) throws IOException {
        List<List<String>> tableData = new ArrayList<>();
        List<Integer> columnWidths = new ArrayList<>();
        
        Node child = tableBlock.getFirstChild();
        while (child != null) {
            if (child instanceof TableHead || child instanceof TableBody) {
                processTableSection(child, tableData);
            }
            child = child.getNext();
        }
        
        if (!tableData.isEmpty()) {
            for (List<String> row : tableData) {
                for (int i = 0; i < row.size(); i++) {
                    if (columnWidths.size() <= i) {
                        columnWidths.add(0);
                    }
                    columnWidths.set(i, Math.max(columnWidths.get(i), row.get(i).length()));
                }
            }
            
            for (int rowIndex = 0; rowIndex < tableData.size(); rowIndex++) {
                List<String> row = tableData.get(rowIndex);
                writer.write("| ");
                for (int i = 0; i < row.size(); i++) {
                    writer.write(padRight(row.get(i), columnWidths.get(i)));
                    writer.write(" | ");
                }
                writer.write("\n");
                
                if (rowIndex == 0 && child instanceof TableHead) {
                    writer.write("|");
                    for (int width : columnWidths) {
                        writer.write("-");
                        for (int i = 0; i < width; i++) {
                            writer.write("-");
                        }
                        writer.write("-|");
                    }
                    writer.write("\n");
                }
            }
            writer.write("\n");
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
    
    private void writeIndent() throws IOException {
        for (int i = 0; i < listLevel - 1; i++) {
            writer.write("  ");
        }
    }
    
    private String padRight(String text, int length) {
        StringBuilder result = new StringBuilder(text);
        while (result.length() < length) {
            result.append(" ");
        }
        return result.toString();
    }
    
    private void extractText(Node node, StringBuilder text) {
        if (node == null) return;
        
        if (node instanceof Text) {
            text.append(((Text) node).getLiteral());
        } else if (node instanceof Code) {
            text.append("`").append(((Code) node).getLiteral()).append("`");
        } else if (node instanceof Emphasis) {
            text.append("*");
            Node child = node.getFirstChild();
            while (child != null) {
                extractText(child, text);
                child = child.getNext();
            }
            text.append("*");
        } else if (node instanceof StrongEmphasis) {
            text.append("**");
            Node child = node.getFirstChild();
            while (child != null) {
                extractText(child, text);
                child = child.getNext();
            }
            text.append("**");
        } else if (node instanceof Link) {
            text.append("[");
            Node child = node.getFirstChild();
            while (child != null) {
                extractText(child, text);
                child = child.getNext();
            }
            text.append("](").append(((Link) node).getDestination()).append(")");
        } else if (node instanceof Image) {
            text.append("![").append(((Image) node).getTitle() != null ? ((Image) node).getTitle() : "")
                .append("](").append(((Image) node).getDestination()).append(")");
        } else if (node instanceof HardLineBreak || node instanceof SoftLineBreak) {
            text.append("\n");
        } else {
            Node child = node.getFirstChild();
            while (child != null) {
                extractText(child, text);
                child = child.getNext();
            }
        }
    }
}