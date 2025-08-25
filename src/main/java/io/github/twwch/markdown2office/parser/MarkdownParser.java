package io.github.twwch.markdown2office.parser;

import org.commonmark.Extension;
import org.commonmark.ext.gfm.tables.TablesExtension;
import org.commonmark.ext.task.list.items.TaskListItemsExtension;
import org.commonmark.node.Node;
import org.commonmark.parser.Parser;

import java.util.Arrays;
import java.util.List;

public class MarkdownParser {
    
    private final Parser parser;
    
    public MarkdownParser() {
        List<Extension> extensions = Arrays.asList(
            TablesExtension.create(),
            TaskListItemsExtension.create()
        );
        
        this.parser = Parser.builder()
                .extensions(extensions)
                .build();
    }
    
    public Node parse(String markdown) {
        if (markdown == null) {
            throw new IllegalArgumentException("Markdown content cannot be null");
        }
        return parser.parse(markdown);
    }
}