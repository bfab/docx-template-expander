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
import org.junit.BeforeClass;
import org.junit.Test;

import com.google.common.collect.ImmutableMap;
import com.google.common.io.Files;

import zipcomparator.ZipComparator;

public class DocxProcessorTest {
	private static final Map<String, String> EMPTY_MAP = ImmutableMap.of();
	private static final File testTemplate = loadReadOnlyFromClasspath("/template.docx");
	private static final File expected = loadReadOnlyFromClasspath("/expected.docx");
	private static File tempFolder;
	private static File templateEquivalent;
	private File resultDocx;
	
	@BeforeClass
	public static void prepareTempFolder() throws IOException {
		tempFolder = java.nio.file.Files.createTempDirectory(DocxProcessorTest.class.getSimpleName()).toFile();
		System.out.println(tempFolder.toString());
	}
	
	@BeforeClass
	public static void preparePristineCopy() throws InvalidFormatException, IOException {
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(testTemplate));
        
        templateEquivalent = new File(tempFolder, "./templateEquivalent.docx");
        templateEquivalent.createNewFile();
				
        try (FileOutputStream out = new FileOutputStream(templateEquivalent)) {
        	doc.write(out);
        }
        
        templateEquivalent.setWritable(false);
	}

	private static File loadReadOnlyFromClasspath(String relPath) {
		URL urlTemplate = DocxProcessorTest.class.getResource(relPath);
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

	@Test(expected=NullPointerException.class)
	public void testComplainsIfTemplateIsNull() {
		new DocxProcessor(null);
	}

	@Test(expected=IllegalStateException.class)
	public void testComplainsIfTemplateIsNotReadable() throws IOException, InvalidFormatException {
		
		File unreadable = new File(tempFolder, "./unreadableTemplate.docx");
		Files.copy(testTemplate, unreadable);
		DocxProcessor dp = new DocxProcessor(unreadable);
		unreadable.setReadable(false);
		dp.process(EMPTY_MAP, resultDocx);
	}

	@Test(expected=NullPointerException.class)
	public void testComplainsIfDestinationIsNull() throws IOException, InvalidFormatException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		dp.process(EMPTY_MAP, null);
	}

	@Test(expected=NullPointerException.class)
	public void testComplainsIfMapIsNull() throws IOException, InvalidFormatException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		dp.process(null, resultDocx);
	}

	@Test(expected=IllegalArgumentException.class)
	public void testComplainsIfDestinationIsNotWritable() throws IOException, InvalidFormatException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		resultDocx.setWritable(false);
		dp.process(EMPTY_MAP, resultDocx);
	}

	@Test
	public void testCanCopyAFile() throws IOException, InvalidFormatException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		File target = new File(tempFolder, "./processedWithougSubstitutions.docx");
		target.createNewFile();
		dp.process(EMPTY_MAP, target);
		assertTrue(ZipComparator.equal(target, templateEquivalent));
	}

	@Test
	public void testCanProcessAFile() throws IOException, InvalidFormatException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		
		Map<String, String> map = ImmutableMap.of(
				Pattern.quote("{{SUB1}}"), "test value 1",
				Pattern.quote("{{SUB2}}"), "a_very_long_value: - The quick brown fox jumps over the lazy dog",
				Pattern.quote("{{SUB3}}"), "this\nis\na multiline\nvalue",
				Pattern.quote("{{SUB4}}"), ""
				);
		File target = new File(tempFolder, "./processed_1.docx");
		target.createNewFile();
		dp.process(map, target);
		assertTrue(ZipComparator.equal(target, expected));
	}

}
