package io.github.twwch.markdown2office.parser;

import io.github.twwch.markdown2office.parser.impl.WordFileParser;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Test for Word document parsing (both DOC and DOCX formats)
 */
public class WordDocParsingTest {
    
    @Test
    public void testSupportsDocAndDocx() {
        WordFileParser parser = new WordFileParser();
        
        // Test DOC format
        assertTrue(parser.supports("document.doc"));
        assertTrue(parser.supports("Document.DOC"));
        assertTrue(parser.supports("test.doc"));
        
        // Test DOCX format
        assertTrue(parser.supports("document.docx"));
        assertTrue(parser.supports("Document.DOCX"));
        assertTrue(parser.supports("test.docx"));
        
        // Test non-Word formats
        assertFalse(parser.supports("document.pdf"));
        assertFalse(parser.supports("document.txt"));
        assertFalse(parser.supports(null));
    }
    
    @Test
    public void testUnsupportedFormatError() {
        WordFileParser parser = new WordFileParser();
        
        // Create a simple text file content (not OLE2 or OOXML)
        byte[] textContent = "This is plain text".getBytes();
        ByteArrayInputStream inputStream = new ByteArrayInputStream(textContent);
        
        // Should throw IOException for unsupported format
        assertThrows(IOException.class, () -> {
            parser.parse(inputStream, "test.doc");
        });
    }
    
    /**
     * Note: To fully test DOC and DOCX parsing, you would need actual DOC and DOCX files.
     * This test verifies the basic structure and format detection logic.
     * 
     * For production testing, you should:
     * 1. Create sample DOC and DOCX files with known content
     * 2. Place them in src/test/resources
     * 3. Test parsing and verify the extracted content
     */
}