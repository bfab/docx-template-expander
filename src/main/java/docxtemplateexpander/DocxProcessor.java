package docxtemplateexpander;

import java.io.File;
import java.util.Map;

import com.google.common.base.Preconditions;

public final class DocxProcessor {

	public DocxProcessor(File testTemplate) {
		Preconditions.checkNotNull(testTemplate);
	}

	public void process(Map<String, String> substitutionMap, File resultDocxPath) {
		Preconditions.checkNotNull(substitutionMap);
		Preconditions.checkNotNull(resultDocxPath);
		if (substitutionMap.isEmpty());
	}

}
