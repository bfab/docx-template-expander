package docxtemplateexpander;

import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

import com.google.common.collect.ImmutableMap;
import com.google.common.io.Files;

public class DocxProcessorTest {
	private static final Map<String, String> EMPTY_MAP = ImmutableMap.of();
	static File testTemplate;
	static File templateEquivalent;
	File resultDocx;
	
	@BeforeClass
	public static void preparePristineCopy() throws InvalidFormatException, IOException {
		URL url = DocxProcessorTest.class.getResource("/template.docx");
		testTemplate = new File(url.getFile());
		testTemplate.setWritable(false);
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(testTemplate));
        
//        templateEquivalent = File.createTempFile(DocxProcessorTest.class.getSimpleName(), ".docx");
        templateEquivalent = new File("./result.docx");
        templateEquivalent.createNewFile();
		
        try (FileOutputStream out = new FileOutputStream(templateEquivalent)) {
        	doc.write(out);
        }
        
        templateEquivalent.setWritable(false);
	}
	
	@Before
	public void setup() throws IOException {
		resultDocx = File.createTempFile(this.getClass().getSimpleName(), ".docx");
	}

	@Test(expected=NullPointerException.class)
	public void testComplainsIfTemplateIsNull() {
		new DocxProcessor(null);
	}

	@Test(expected=IllegalStateException.class)
	public void testComplainsIfTemplateIsNotReadable() throws IOException, InvalidFormatException {
		File unreadable = File.createTempFile(this.getClass().getSimpleName(), ".docx");
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
		resultDocx = new File("./equiv.docx");
		resultDocx.createNewFile();
		dp.process(EMPTY_MAP, resultDocx);
		assertTrue(Files.equal(resultDocx, templateEquivalent));
	}

	@Test
	public void testCanProcessAFile() throws IOException, InvalidFormatException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		
		Map<String, String> map = ImmutableMap.of(
				Pattern.quote("{{sub1}}"), "test value 1",
				Pattern.quote("{{sub2}}"), "a_very_long_value: - The quick brown fox jumps over the lazy dog",
				Pattern.quote("{{sub3}}"), "this\nis\na multiline\nvalue"
				);
		resultDocx = new File("./result.docx");
		resultDocx.createNewFile();
		dp.process(map, resultDocx);
	}

}
