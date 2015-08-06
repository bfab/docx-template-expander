package docxtemplateexpander;

import org.junit.Test;

public class DocxProcessorTest {
	@Test(expected=NullPointerException.class)
	public void testComplainsIfTemplateIsNull() {
		new DocxProcessor(null);
	}

}
