package io.github.twwch.markdown2office;

import io.github.twwch.markdown2office.model.FileType;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;

public class Markdown2OfficeTest {
    
    private Markdown2Office converter;
    private String sampleMarkdown;
    
    @BeforeEach
    public void setUp() {
        converter = new Markdown2Office();
        sampleMarkdown = "# 二战的详细经过\n" +
                "\n" +
                "## Document: 二战的经过 (The Process of World War II)\n" +
                "\n" +
                "### 目录 (Table of Contents)\n" +
                "\n" +
                "1. 引言 (Introduction)\n" +
                "2. 战争的起因 (Causes of the War)\n" +
                "3. 主要战役与事件 (Key Battles and Events)\n" +
                "   - 3.1 1939-1941年 (1939-1941)\n" +
                "   - 3.2 1942年 (1942)\n" +
                "   - 3.3 1943年 (1943)\n" +
                "   - 3.4 1944年 (1944)\n" +
                "   - 3.5 1945年 (1945)\n" +
                "4. 战争的结束与影响 (Conclusion and Impact)\n" +
                "5. 参考资料 (References)\n" +
                "\n" +
                "### 1. 引言 (Introduction)\n" +
                "\n" +
                "第二次世界大战是人类历史上规模最大、持续时间最长的战争之一，涉及全球几乎所有大国。这场战争不仅改变了战斗方式，也深刻影响了全球的政治、经济和社会结构。本文将详细描述二战的经过，包括其起因、主要战役和影响。\n" +
                "\n" +
                "### 2. 战争的起因 (Causes of the War)\n" +
                "\n" +
                "二战的起因可以追溯到一战后的不平等条约，以及全球经济危机的影响。对德国的惩罚性条款激发了民族主义情绪，希特勒的崛起成为了战争的催化剂。此外，意大利和日本也寻找机会扩展其领土，导致了轴心国的形成。\n" +
                "\n" +
                "### 3. 主要战役与事件 (Key Battles and Events)\n" +
                "\n" +
                "#### 3.1 1939-1941年 (1939-1941)\n" +
                "\n" +
                "- **1939年9月1日**：德国入侵波兰，标志着二战的开始。\n" +
                "- **1940年5月**：德国闪电战击败挪威、丹麦及法国，迅速扩展其领土。\n" +
                "- **1940年6月**：法国投降，英国成为对抗轴心国的主要力量。\n" +
                "\n" +
                "#### 3.2 1942年 (1942)\n" +
                "\n" +
                "- **在北非的战斗**：英军与德意军队展开多次交锋，尤其是埃尔阿拉梅因战役，标志着盟军的转折点。\n" +
                "- **珍珠港事件**：1941年12月7日，美国遭到日本袭击，正式参战。\n" +
                "\n" +
                "#### 3.3 1943年 (1943)\n" +
                "\n" +
                "- **斯大林格勒战役**：苏联军队抵抗德军进攻，此役被视为战争的转折点。\n" +
                "- **盟军登陆西西里**：标志着欧洲大陆反攻的开始。\n" +
                "\n" +
                "#### 3.4 1944年 (1944)\n" +
                "\n" +
                "- **诺曼底登陆**：1944年6月6日，盟军展开了史上最大规模的海陆空联合进攻。\n" +
                "- **解放巴黎**：1944年8月，盟军成功解放巴黎，意味着对纳粹政权的逐步压制。\n" +
                "\n" +
                "#### 3.5 1945年 (1945)\n" +
                "\n" +
                "- **雅尔塔会议**：盟军领导人在战争末期进行战略会谈，决定战后布局。\n" +
                "- **德国无条件投降**：1945年5月，欧洲战场结束。\n" +
                "- **日本战败**：1945年8月，美国在广岛和长崎投下原子弹后，日本宣布无条件投降，二战全局结束。\n" +
                "\n" +
                "### 4. 战争的结束与影响 (Conclusion and Impact)\n" +
                "\n" +
                "第二次世界大战结束后，全球秩序发生了巨变。联合国的成立旨在防止未来的冲突，冷战格局的形成则引发了后续数十年的国际紧张关系。战后，欧洲重建和去殖民化的进程加速，世界开始朝向多极化发展。\n" +
                "\n" +
                "### 5. 参考资料 (References)\n" +
                "\n" +
                "1. 中华人民共和国语文出版社 (The People's Republic of China Language Literature Press)\n" +
                "2. 历史学理论及方法 (Theory and Method of History)\n" +
                "3. 相关历史书籍及文献 (Relevant Historical Books and Literature)\n" +
                "\n" +
                "---\n" +
                "\n" +
                "\n" +
                "- Convert Markdown to multiple formats:\n" +
                "  - Word (DOCX)\n" +
                "  - Excel (XLSX)\n" +
                "  - PDF\n" +
                "  - Plain Text (TXT)\n" +
                "  - Markdown (MD)\n" +
                "- Preserve Markdown structure and formatting\n" +
                "- Support for tables, lists, code blocks, and more\n" +
                "- Easy-to-use API\n" +
                "- Command-line interface\n";
    }
    
    @Test
    public void testConvertToWord() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.WORD, outputStream);
        
        byte[] result = outputStream.toByteArray();
        assertNotNull(result);
        // 写入文件
        Files.write(new File("output.docx").toPath(), result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testConvertToExcel() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.EXCEL, outputStream);
        
        byte[] result = outputStream.toByteArray();
        assertNotNull(result);

        Files.write(new File("output.xlsx").toPath(), result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testConvertToPdf() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.PDF, outputStream);
        
        byte[] result = outputStream.toByteArray();

        Files.write(new File("output.pdf").toPath(), result);
        assertNotNull(result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testConvertToText() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.TEXT, outputStream);
        
        String result = outputStream.toString("UTF-8");
        assertNotNull(result);
        Files.write(new File("output.txt").toPath(), result.getBytes());
        // Check that the text contains Chinese content from the sample markdown
        assertTrue(result.contains("二战"));
        assertTrue(result.contains("引言"));
        assertTrue(result.contains("1939"));
    }
    
    @Test
    public void testConvertToMarkdown() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        converter.convert(sampleMarkdown, FileType.MARKDOWN, outputStream);
        
        String result = outputStream.toString("UTF-8");
        Files.write(new File("output.md").toPath(), result.getBytes());
        assertEquals(sampleMarkdown, result);
    }
    
    @Test
    public void testConvertToFile(@TempDir Path tempDir) throws IOException {
        File outputFile = tempDir.resolve("test.docx").toFile();
        converter.convert(sampleMarkdown, FileType.WORD, outputFile);
        
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertFileToFile(@TempDir Path tempDir) throws IOException {
        Path inputFile = tempDir.resolve("input.md");
        Files.write(inputFile, sampleMarkdown.getBytes());
        
        File outputFile = tempDir.resolve("output.pdf").toFile();
        converter.convertFile(inputFile.toString(), FileType.PDF, outputFile.toString());
        
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertWithAutoDetection(@TempDir Path tempDir) throws IOException {
        Path inputFile = tempDir.resolve("input.md");
        Files.write(inputFile, sampleMarkdown.getBytes());
        
        String outputPath = tempDir.resolve("output.xlsx").toString();
        converter.convertFile(inputFile.toString(), outputPath);
        
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertToBytes() throws IOException {
        byte[] result = converter.convertToBytes(sampleMarkdown, FileType.WORD);
        
        assertNotNull(result);
        assertTrue(result.length > 0);
    }
    
    @Test
    public void testNullMarkdownThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert(null, FileType.WORD, new ByteArrayOutputStream());
        });
    }
    
    @Test
    public void testEmptyMarkdownThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert("  ", FileType.WORD, new ByteArrayOutputStream());
        });
    }
    
    @Test
    public void testNullFileTypeThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert(sampleMarkdown, null, new ByteArrayOutputStream());
        });
    }
    
    @Test
    public void testNullOutputStreamThrowsException() {
        ByteArrayOutputStream nullStream = null;
        assertThrows(IllegalArgumentException.class, () -> {
            converter.convert(sampleMarkdown, FileType.WORD, nullStream);
        });
    }
    
    @Test
    public void testInvalidFileExtensionThrowsException() {
        assertThrows(IllegalArgumentException.class, () -> {
            FileType.fromExtension("invalid");
        });
    }
}