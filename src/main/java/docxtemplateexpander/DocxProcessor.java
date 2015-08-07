package docxtemplateexpander;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.google.common.base.Preconditions;

public final class DocxProcessor {

	private final File template;

	public DocxProcessor(File testTemplate) {
		Preconditions.checkNotNull(testTemplate);
		
		template = testTemplate;
	}

	/**
	 * In the given map, keys are regexes whose matches will be replaced by their respective value.
	 * 
	 * Note: Key matches need to be ignored by the document grammar checking (having matches be upper-case without spaces is usually enough).
	 * 
	 * Note: comments are not processed (they're just copied over verbatim).
	 * 
	 * @param substitutionMap
	 * @param resultDocx
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public void process(Map<String, String> substitutionMap, File resultDocx) throws IOException, InvalidFormatException {
		Preconditions.checkNotNull(substitutionMap);
		Preconditions.checkNotNull(resultDocx);
		Preconditions.checkArgument(resultDocx.canWrite());

		Preconditions.checkState(template.canRead());
		
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(template));

        for (XWPFParagraph par : doc.getParagraphs())
        	processParagraph(substitutionMap, par);
        
        for (XWPFTable tbl : doc.getTables())
			processTable(substitutionMap, tbl);

        for (XWPFFootnote fNote : doc.getFootnotes())
			processFootnote(substitutionMap, fNote);
        
        for (XWPFHeader header : doc.getHeaderList()) {
        	processHeaderFooter(substitutionMap, header);
		}
        
        for (XWPFFooter footer : doc.getFooterList()) {
        	processHeaderFooter(substitutionMap, footer);
		}
        
        for (XWPFFootnote fNote : doc.getFootnotes())
			processFootnote(substitutionMap, fNote);
        
        try (FileOutputStream out = new FileOutputStream(resultDocx)) {
        	doc.write(out);
        }

    }

	private static void processHeaderFooter(Map<String, String> substitutionMap, XWPFHeaderFooter hf) {
		for (XWPFParagraph par : hf.getParagraphs()) {
			processParagraph(substitutionMap, par);
		}
	}

	private static void processFootnote(Map<String, String> substitutionMap, XWPFFootnote fNote) {
		for (XWPFParagraph par : fNote.getParagraphs())
			processParagraph(substitutionMap, par);
	}

	private static void processTable(Map<String, String> substitutionMap, XWPFTable tbl) {
		for (XWPFTableRow row : tbl.getRows())
			for (XWPFTableCell cell : row.getTableCells())
				for (XWPFParagraph p : cell.getParagraphs())
					processParagraph(substitutionMap, p);
	}

	private static void processParagraph(Map<String, String> substitutionMap, XWPFParagraph par) {
		for(XWPFRun run : par.getRuns()) {
			for (int i=0; i< run.getCTR().sizeOfTArray(); ++i) {
				String text = run.getText(i);
				for (Entry<String, String> subst : substitutionMap.entrySet()) {
					text = text.replaceAll(subst.getKey(), subst.getValue());
				}
				run.setText(text, i);
			}
		}
	}
}
