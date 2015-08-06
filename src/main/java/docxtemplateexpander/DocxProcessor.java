package docxtemplateexpander;

import java.io.File;
import java.io.IOException;
import java.util.Map;

import com.google.common.base.Preconditions;
import com.google.common.io.Files;

public final class DocxProcessor {

	private final File template;

	public DocxProcessor(File testTemplate) {
		Preconditions.checkNotNull(testTemplate);
		
		template = testTemplate;
	}

	public void process(Map<String, String> substitutionMap, File resultDocx) throws IOException {
		Preconditions.checkNotNull(substitutionMap);
		Preconditions.checkNotNull(resultDocx);
		Preconditions.checkArgument(resultDocx.canWrite());

		Preconditions.checkState(template.canRead());
		
		Files.copy(template, resultDocx);
	}

}
