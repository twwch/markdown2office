package io.github.twwch.markdown2office.converter;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;

public class MarkdownConverter implements Converter {
    
    @Override
    public void convert(String markdown, OutputStream outputStream) throws IOException {
        Writer writer = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8);
        writer.write(markdown);
        writer.flush();
        writer.close();
    }
}