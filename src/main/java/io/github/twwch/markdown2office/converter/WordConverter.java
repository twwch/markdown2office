package io.github.twwch.markdown2office.converter;

import io.github.twwch.markdown2office.parser.MarkdownParser;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.commonmark.ext.gfm.strikethrough.Strikethrough;
import org.commonmark.ext.gfm.tables.*;
import org.commonmark.ext.task.list.items.TaskListItemMarker;
import org.commonmark.node.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.net.HttpURLConnection;
import java.net.URL;
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
        createHeadingStyles();
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
        
        // Apply heading formatting directly instead of relying on styles
        XWPFRun run = currentParagraph.createRun();
        run.setBold(true);
        
        // Set font size based on heading level
        switch (level) {
            case 1:
                run.setFontSize(24);
                currentParagraph.setSpacingAfter(300); // 15pt spacing after
                currentParagraph.setSpacingBefore(240); // 12pt spacing before
                break;
            case 2:
                run.setFontSize(20);
                currentParagraph.setSpacingAfter(260); // 13pt spacing after
                currentParagraph.setSpacingBefore(200); // 10pt spacing before
                break;
            case 3:
                run.setFontSize(18);
                currentParagraph.setSpacingAfter(240); // 12pt spacing after
                currentParagraph.setSpacingBefore(160); // 8pt spacing before
                break;
            case 4:
                run.setFontSize(16);
                currentParagraph.setSpacingAfter(200); // 10pt spacing after
                currentParagraph.setSpacingBefore(140); // 7pt spacing before
                break;
            case 5:
                run.setFontSize(14);
                currentParagraph.setSpacingAfter(160); // 8pt spacing after
                currentParagraph.setSpacingBefore(120); // 6pt spacing before
                break;
            case 6:
                run.setFontSize(13);
                currentParagraph.setSpacingAfter(140); // 7pt spacing after
                currentParagraph.setSpacingBefore(100); // 5pt spacing before
                break;
            default:
                run.setFontSize(12);
                break;
        }
        
        // Extract text from heading node and add to the run
        StringBuilder headingText = new StringBuilder();
        extractTextFromNode(heading.getFirstChild(), headingText);
        run.setText(headingText.toString());
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
                
                // Check for task list items
                Node firstChild = item.getFirstChild();
                boolean isTaskItem = false;
                if (firstChild instanceof TaskListItemMarker) {
                    TaskListItemMarker marker = (TaskListItemMarker) firstChild;
                    XWPFRun checkRun = currentParagraph.createRun();
                    checkRun.setText(marker.isChecked() ? "[✓] " : "[ ] ");
                    isTaskItem = true;
                }
                
                if (!isTaskItem) {
                    if (ordered) {
                        currentParagraph.setNumID(BigInteger.valueOf(1));
                    } else {
                        XWPFRun run = currentParagraph.createRun();
                        run.setText("• ");
                    }
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
            } else if (node instanceof Strikethrough) {
                XWPFRun run = currentParagraph.createRun();
                run.setStrikeThrough(true);
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                run.setText(text.toString());
            } else if (node instanceof Link) {
                XWPFRun run = currentParagraph.createRun();
                run.setUnderline(UnderlinePatterns.SINGLE);
                run.setColor("0000FF");
                StringBuilder text = new StringBuilder();
                extractText(node, text);
                run.setText(text.toString());
            } else if (node instanceof Image) {
                Image imageNode = (Image) node;
                String imageUrl = imageNode.getDestination();
                XWPFRun run = currentParagraph.createRun();
                
                try {
                    InputStream imageStream = null;
                    String imageType = null;
                    int width = 0;
                    int height = 0;
                    
                    if (imageUrl != null && !imageUrl.isEmpty()) {
                        if (imageUrl.startsWith("http://") || imageUrl.startsWith("https://")) {
                            // Load from URL
                            URL url = new URL(imageUrl);
                            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                            conn.setRequestMethod("GET");
                            conn.setConnectTimeout(5000);
                            conn.setReadTimeout(5000);
                            imageStream = conn.getInputStream();
                            
                            // Determine image type from URL
                            String urlLower = imageUrl.toLowerCase();
                            if (urlLower.endsWith(".png")) {
                                imageType = "png";
                            } else if (urlLower.endsWith(".jpg") || urlLower.endsWith(".jpeg")) {
                                imageType = "jpeg";
                            } else if (urlLower.endsWith(".gif")) {
                                imageType = "gif";
                            } else {
                                imageType = "png"; // default
                            }
                            
                            // Get image dimensions
                            BufferedImage bimg = ImageIO.read(url);
                            if (bimg != null) {
                                width = bimg.getWidth();
                                height = bimg.getHeight();
                            }
                        } else {
                            // Try as local file
                            java.io.File imageFile = new java.io.File(imageUrl);
                            if (imageFile.exists()) {
                                imageStream = new java.io.FileInputStream(imageFile);
                                
                                // Determine image type from file extension
                                String fileName = imageFile.getName().toLowerCase();
                                if (fileName.endsWith(".png")) {
                                    imageType = "png";
                                } else if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) {
                                    imageType = "jpeg";
                                } else if (fileName.endsWith(".gif")) {
                                    imageType = "gif";
                                } else {
                                    imageType = "png"; // default
                                }
                                
                                // Get image dimensions
                                BufferedImage bimg = ImageIO.read(imageFile);
                                if (bimg != null) {
                                    width = bimg.getWidth();
                                    height = bimg.getHeight();
                                }
                            }
                        }
                    }
                    
                    if (imageStream != null && imageType != null) {
                        // Determine picture type
                        int pictureType = XWPFDocument.PICTURE_TYPE_PNG;
                        if ("jpeg".equals(imageType)) {
                            pictureType = XWPFDocument.PICTURE_TYPE_JPEG;
                        } else if ("gif".equals(imageType)) {
                            pictureType = XWPFDocument.PICTURE_TYPE_GIF;
                        }
                        
                        // Calculate scaled dimensions to fit page (max width 500px)
                        int maxWidth = 500;
                        int scaledWidth = width;
                        int scaledHeight = height;
                        
                        if (width > maxWidth) {
                            scaledWidth = maxWidth;
                            scaledHeight = (height * maxWidth) / width;
                        }
                        
                        // Add picture to document
                        run.addPicture(imageStream, pictureType, imageUrl,
                                      Units.toEMU(scaledWidth), Units.toEMU(scaledHeight));
                        imageStream.close();
                    } else {
                        // Fallback if image cannot be loaded
                        String altText = imageNode.getTitle() != null ? imageNode.getTitle() : "Image";
                        run.setText("[" + altText + "]");
                    }
                } catch (Exception e) {
                    // If image loading fails, show alt text
                    String altText = imageNode.getTitle() != null ? imageNode.getTitle() : "Image";
                    run.setText("[" + altText + ": " + imageUrl + "]");
                }
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
    
    private void createHeadingStyles() {
        // This method creates proper heading styles in the document
        // In Apache POI, styles are created automatically when used, but we can set up
        // default paragraph formatting to ensure consistent appearance
    }
    
    private void extractTextFromNode(Node node, StringBuilder text) {
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
            extractTextFromNode(child, text);
            child = child.getNext();
        }
    }
}