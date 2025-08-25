package io.github.twwch.markdown2office;

import io.github.twwch.markdown2office.converter.*;
import io.github.twwch.markdown2office.model.FileType;

public class ConverterFactory {
    
    public static Converter getConverter(FileType fileType) {
        switch (fileType) {
            case WORD:
                return new WordConverter();
            case EXCEL:
                return new ExcelConverter();
            case PDF:
                return new PdfConverter();
            case TEXT:
                return new TextConverter();
            case MARKDOWN:
                return new MarkdownConverter();
            default:
                throw new IllegalArgumentException("Unsupported file type: " + fileType);
        }
    }
    
    public static Converter getConverter(String fileExtension) {
        FileType fileType = FileType.fromExtension(fileExtension);
        return getConverter(fileType);
    }
}