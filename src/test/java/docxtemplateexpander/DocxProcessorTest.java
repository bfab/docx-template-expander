package docxtemplateexpander;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;

import com.google.common.collect.ImmutableMap;

public class DocxProcessorTest {
	private static final Map<String, String> EMPTY_MAP = ImmutableMap.of();
	File testTemplate;
	File resultDocxPath;
	
	@Before
	public void setup() throws IOException {
		URL url = this.getClass().getResource("/template.docx");
		testTemplate = new File(url.getFile());
		resultDocxPath = File.createTempFile(this.getClass().getSimpleName(), ".docx");
	}

	@Test(expected=NullPointerException.class)
	public void testComplainsIfTemplateIsNull() {
		new DocxProcessor(null);
	}

	@Test(expected=IllegalStateException.class)
	public void testComplainsIfTemplateIsNotReadable() throws IOException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		testTemplate.setReadable(false);
		dp.process(EMPTY_MAP, resultDocxPath);
	}

	@Test(expected=NullPointerException.class)
	public void testComplainsIfDestinationIsNull() throws IOException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		dp.process(EMPTY_MAP, null);
	}

	@Test(expected=NullPointerException.class)
	public void testComplainsIfMapIsNull() throws IOException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		dp.process(null, resultDocxPath);
	}

	@Test(expected=IllegalArgumentException.class)
	public void testComplainsIfDestinationIsNotWritable() throws IOException {
		DocxProcessor dp = new DocxProcessor(testTemplate);
		resultDocxPath.setWritable(false);
		dp.process(EMPTY_MAP, resultDocxPath);
	}

}
