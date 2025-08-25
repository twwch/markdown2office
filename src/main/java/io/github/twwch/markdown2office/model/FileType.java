package io.github.twwch.markdown2office.model;

public enum FileType {
    WORD("docx"),
    EXCEL("xlsx"),
    PDF("pdf"),
    MARKDOWN("md"),
    TEXT("txt");

    private final String extension;

    FileType(String extension) {
        this.extension = extension;
    }

    public String getExtension() {
        return extension;
    }

    public static FileType fromExtension(String extension) {
        for (FileType type : values()) {
            if (type.extension.equalsIgnoreCase(extension)) {
                return type;
            }
        }
        throw new IllegalArgumentException("Unsupported file extension: " + extension);
    }
}