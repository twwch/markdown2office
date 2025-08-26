package io.github.twwch.markdown2office.parser.example;

import io.github.twwch.markdown2office.parser.ParsedDocument;
import io.github.twwch.markdown2office.parser.UniversalFileParser;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

/**
 * Example demonstrating how to use the file parsers
 */
public class FileParserExample {
    
    public static void main(String[] args) {
        UniversalFileParser parser = new UniversalFileParser();
        
        System.out.println("=== File Parser Example ===\n");
        
        // Show supported extensions
        System.out.println("Supported file extensions:");
        String[] extensions = parser.getSupportedExtensions();
        for (String ext : extensions) {
            System.out.println("  " + ext);
        }
        System.out.println();
        
        // Parse different file types from streams
        try {
            // Parse Markdown
            parseMarkdownExample(parser);
            System.out.println();
            
            // Parse CSV
            parseCsvExample(parser);
            System.out.println();
            
            // Parse Plain Text
            parseTextExample(parser);
            System.out.println();
            
            // Parse HTML
            parseHtmlExample(parser);
            System.out.println();
            
        } catch (IOException e) {
            System.err.println("Error parsing files: " + e.getMessage());
        }
        
        // Show parser information
        System.out.println("Available parsers:");
        System.out.println(parser.getParserInfo());
    }
    
    private static void parseMarkdownExample(UniversalFileParser parser) throws IOException {
        System.out.println("=== Parsing Markdown ===");
        String markdown = "# Sample Document\n\n" +
                         "This is a **sample** document with *formatting*.\n\n" +
                         "## Features\n\n" +
                         "- List item 1\n" +
                         "- List item 2\n" +
                         "- List item 3\n\n" +
                         "### Table Example\n\n" +
                         "| Name | Age | City |\n" +
                         "|------|-----|------|\n" +
                         "| John | 30  | NYC  |\n" +
                         "| Jane | 25  | LA   |";
        
        ByteArrayInputStream stream = new ByteArrayInputStream(markdown.getBytes(StandardCharsets.UTF_8));
        ParsedDocument doc = parser.parse(stream, "sample.md");
        
        System.out.println("Title: " + doc.getTitle());
        System.out.println("File Type: " + doc.getFileType());
        System.out.println("Tables found: " + doc.getTables().size());
        System.out.println("Parser used: " + doc.getMetadata().get("Parser Used"));
        System.out.println("\nMarkdown content preview:");
        System.out.println(doc.getMarkdownContent().substring(0, Math.min(200, doc.getMarkdownContent().length())) + "...");
    }
    
    private static void parseCsvExample(UniversalFileParser parser) throws IOException {
        System.out.println("=== Parsing CSV ===");
        String csv = "Product,Price,Category,In Stock\n" +
                    "Laptop,999.99,Electronics,true\n" +
                    "Book,19.99,Education,true\n" +
                    "Desk Chair,149.99,Furniture,false\n" +
                    "Smartphone,699.99,Electronics,true";
        
        ByteArrayInputStream stream = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));
        ParsedDocument doc = parser.parse(stream, "products.csv");
        
        System.out.println("Title: " + doc.getTitle());
        System.out.println("File Type: " + doc.getFileType());
        System.out.println("Tables found: " + doc.getTables().size());
        System.out.println("Parser used: " + doc.getMetadata().get("Parser Used"));
        
        if (!doc.getTables().isEmpty()) {
            ParsedDocument.ParsedTable table = doc.getTables().get(0);
            System.out.println("Table headers: " + table.getHeaders());
            System.out.println("Table rows: " + table.getData().size());
            System.out.println("First row data: " + table.getData().get(0));
        }
    }
    
    private static void parseTextExample(UniversalFileParser parser) throws IOException {
        System.out.println("=== Parsing Plain Text ===");
        String text = "Sample Text Document\n\n" +
                     "This is a plain text file with some content.\n" +
                     "It has multiple lines and paragraphs.\n\n" +
                     "SECTION HEADER\n" +
                     "This section contains important information.\n\n" +
                     "    // This looks like code\n" +
                     "    function example() {\n" +
                     "        return 'hello world';\n" +
                     "    }";
        
        ByteArrayInputStream stream = new ByteArrayInputStream(text.getBytes(StandardCharsets.UTF_8));
        ParsedDocument doc = parser.parse(stream, "sample.txt");
        
        System.out.println("Title: " + doc.getTitle());
        System.out.println("File Type: " + doc.getFileType());
        System.out.println("Parser used: " + doc.getMetadata().get("Parser Used"));
        System.out.println("Line count: " + doc.getMetadata().get("Line Count"));
        System.out.println("\nContent preview:");
        System.out.println(doc.getContent().substring(0, Math.min(150, doc.getContent().length())) + "...");
    }
    
    private static void parseHtmlExample(UniversalFileParser parser) throws IOException {
        System.out.println("=== Parsing HTML (using Tika) ===");
        String html = "<!DOCTYPE html>\n" +
                     "<html>\n" +
                     "<head><title>Sample HTML</title></head>\n" +
                     "<body>\n" +
                     "<h1>Welcome to HTML Parsing</h1>\n" +
                     "<p>This is a <strong>sample</strong> HTML document.</p>\n" +
                     "<ul>\n" +
                     "<li>Item 1</li>\n" +
                     "<li>Item 2</li>\n" +
                     "</ul>\n" +
                     "</body>\n" +
                     "</html>";
        
        ByteArrayInputStream stream = new ByteArrayInputStream(html.getBytes(StandardCharsets.UTF_8));
        ParsedDocument doc = parser.parse(stream, "sample.html");
        
        System.out.println("Title: " + doc.getTitle());
        System.out.println("File Type: " + doc.getFileType());
        System.out.println("Parser used: " + doc.getMetadata().get("Parser Used"));
        System.out.println("Content Type: " + doc.getMetadata().get("Content Type"));
        System.out.println("\nContent preview:");
        System.out.println(doc.getContent().substring(0, Math.min(100, doc.getContent().length())) + "...");
    }
}