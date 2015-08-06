package docxtemplateexpander;

import java.io.File;
import java.util.Map;

import com.google.common.base.Preconditions;

public final class DocxProcessor {

	private final File template;

	public DocxProcessor(File testTemplate) {
		Preconditions.checkNotNull(testTemplate);
		
		template = testTemplate;
	}

	public void process(Map<String, String> substitutionMap, File resultDocxPath) {
		Preconditions.checkState(template.canRead());
	}

}
