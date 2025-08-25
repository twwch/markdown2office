package io.github.twwch.markdown2office.converter;

import java.io.IOException;
import java.io.OutputStream;

public interface Converter {
    void convert(String markdown, OutputStream outputStream) throws IOException;
}