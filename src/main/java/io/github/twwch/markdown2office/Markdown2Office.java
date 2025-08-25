package io.github.twwch.markdown2office;

import io.github.twwch.markdown2office.converter.Converter;
import io.github.twwch.markdown2office.model.FileType;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Markdown2Office {
    
    public void convert(String markdown, FileType fileType, OutputStream outputStream) throws IOException {
        if (markdown == null || markdown.trim().isEmpty()) {
            throw new IllegalArgumentException("Markdown content cannot be null or empty");
        }
        
        if (fileType == null) {
            throw new IllegalArgumentException("File type cannot be null");
        }
        
        if (outputStream == null) {
            throw new IllegalArgumentException("Output stream cannot be null");
        }
        
        Converter converter = ConverterFactory.getConverter(fileType);
        converter.convert(markdown, outputStream);
    }
    
    public void convert(String markdown, FileType fileType, String outputPath) throws IOException {
        Path path = Paths.get(outputPath);
        Path parentDir = path.getParent();
        
        if (parentDir != null && !Files.exists(parentDir)) {
            Files.createDirectories(parentDir);
        }
        
        try (OutputStream outputStream = new FileOutputStream(outputPath)) {
            convert(markdown, fileType, outputStream);
        }
    }
    
    public void convert(String markdown, FileType fileType, File outputFile) throws IOException {
        File parentDir = outputFile.getParentFile();
        
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs();
        }
        
        try (OutputStream outputStream = new FileOutputStream(outputFile)) {
            convert(markdown, fileType, outputStream);
        }
    }
    
    public void convertFile(String inputPath, FileType fileType, String outputPath) throws IOException {
        String markdown = readFile(inputPath);
        convert(markdown, fileType, outputPath);
    }
    
    public void convertFile(File inputFile, FileType fileType, File outputFile) throws IOException {
        String markdown = readFile(inputFile);
        convert(markdown, fileType, outputFile);
    }
    
    public void convertFile(String inputPath, String outputPath) throws IOException {
        String extension = getFileExtension(outputPath);
        FileType fileType = FileType.fromExtension(extension);
        convertFile(inputPath, fileType, outputPath);
    }
    
    public byte[] convertToBytes(String markdown, FileType fileType) throws IOException {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            convert(markdown, fileType, outputStream);
            return outputStream.toByteArray();
        }
    }
    
    private String readFile(String filePath) throws IOException {
        return new String(Files.readAllBytes(Paths.get(filePath)), StandardCharsets.UTF_8);
    }
    
    private String readFile(File file) throws IOException {
        return new String(Files.readAllBytes(file.toPath()), StandardCharsets.UTF_8);
    }
    
    private String getFileExtension(String filePath) {
        int lastDotIndex = filePath.lastIndexOf('.');
        if (lastDotIndex > 0 && lastDotIndex < filePath.length() - 1) {
            return filePath.substring(lastDotIndex + 1).toLowerCase();
        }
        throw new IllegalArgumentException("Cannot determine file extension from path: " + filePath);
    }
    
    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("Usage: java -jar markdown2office.jar <input.md> <output.ext>");
            System.out.println("Supported output formats: .docx, .xlsx, .pdf, .txt, .md");
            System.exit(1);
        }
        
        try {
            Markdown2Office converter = new Markdown2Office();
            converter.convertFile(args[0], args[1]);
            System.out.println("Conversion successful: " + args[0] + " -> " + args[1]);
        } catch (Exception e) {
            System.err.println("Conversion failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}