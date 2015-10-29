package docxtemplateexpander;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.google.common.collect.ImmutableMap;
import com.google.common.io.Files;

import zipcomparator.ZipComparator;

public class DocxProcessorShould {
    private static final Map<String, String> EMPTY_MAP = ImmutableMap.of();
    private static final File testTemplate = loadReadOnlyFromClasspath("/template.docx");
    private static final File expected = loadReadOnlyFromClasspath("/expected.docx");
    private static final File tempFolder = prepareTempFolder();
    private static final File templateEquivalent = copyDocxThroughPOISerialization(testTemplate);
    private File resultDocx;

    public static File prepareTempFolder() {
        File result;
        try {
            result = java.nio.file.Files.createTempDirectory(
                    DocxProcessorShould.class.getSimpleName()).toFile();
            System.out.println(result.toString());
            
            return result;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static File copyDocxThroughPOISerialization(File original) {
        XWPFDocument doc;
        try {
            doc = new XWPFDocument(OPCPackage.open(original));
            File copy = new File(tempFolder, "./templateEquivalent.docx");
            copy.createNewFile();

            try (FileOutputStream out = new FileOutputStream(copy)) {
                doc.write(out);
            }

            copy.setWritable(false);
            
            return copy;
        } catch (InvalidFormatException | IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static File loadReadOnlyFromClasspath(String relPath) {
        URL urlTemplate = DocxProcessorShould.class.getResource(relPath);
        File f = new File(urlTemplate.getFile());
        f.setWritable(false);

        return f;
    }

    @Before
    public void setup() throws IOException {
        resultDocx = new File(tempFolder, "./result.docx");
        resultDocx.createNewFile();
    }

    @After
    public void teardown() throws IOException {
        resultDocx.setWritable(true);
        resultDocx.delete();
    }

    @Test(expected = NullPointerException.class)
    public void complain_if_template_is_null() {
        new DocxProcessor(null);
    }

    @Test(expected = IllegalStateException.class)
    public void complain_if_template_is_not_readable()
            throws IOException, InvalidFormatException {

        File unreadable = new File(tempFolder, "./unreadableTemplate.docx");
        Files.copy(testTemplate, unreadable);
        DocxProcessor dp = new DocxProcessor(unreadable);
        unreadable.setReadable(false);
        dp.process(EMPTY_MAP, resultDocx);
    }

    @Test(expected = NullPointerException.class)
    public void complain_if_destination_is_null()
            throws IOException, InvalidFormatException {
        DocxProcessor dp = new DocxProcessor(testTemplate);
        dp.process(EMPTY_MAP, null);
    }

    @Test(expected = NullPointerException.class)
    public void complain_if_map_is_null()
            throws IOException, InvalidFormatException {
        DocxProcessor dp = new DocxProcessor(testTemplate);
        dp.process(null, resultDocx);
    }

    @Test(expected = IllegalArgumentException.class)
    public void complain_if_destination_is_not_writable()
            throws IOException, InvalidFormatException {
        DocxProcessor dp = new DocxProcessor(testTemplate);
        resultDocx.setWritable(false);
        dp.process(EMPTY_MAP, resultDocx);
    }

    @Test
    public void copy_the_template_if_empty_map_is_passed() throws IOException, InvalidFormatException {
        DocxProcessor dp = new DocxProcessor(testTemplate);
        File target = new File(tempFolder, "./processedWithougSubstitutions.docx");
        target.createNewFile();
        dp.process(EMPTY_MAP, target);
        assertTrue(ZipComparator.equal(target, templateEquivalent));
    }

    @Test
    public void process_a_file() throws IOException, InvalidFormatException {
        DocxProcessor dp = new DocxProcessor(testTemplate);

        Map<String, String> map = ImmutableMap.of(
                Pattern.quote("{{SUB1}}"),
                "test value 1",
                Pattern.quote("{{SUB2}}"),
                "a_very_long_value: - The quick brown fox jumps over the lazy dog",
                Pattern.quote("{{SUB3}}"),
                "this\nis\na multiline\nvalue",
                Pattern.quote("{{SUB4}}"),
                "");
        File target = new File(tempFolder, "./processed_1.docx");
        target.createNewFile();
        dp.process(map, target);
        assertTrue(ZipComparator.equal(target, expected));
    }

}
