package docxtemplateexpander;

import com.google.common.base.Preconditions;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
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

public final class DocxProcessor {

    private interface OPCPackageProvider {
        OPCPackage getPOCPackage() throws InvalidFormatException, IOException;
    }

	private final OPCPackageProvider opcPackageProvider;

	public DocxProcessor(final File testTemplate) {
		Preconditions.checkNotNull(testTemplate);

		opcPackageProvider = new OPCPackageProvider() {

            @Override
            public OPCPackage getPOCPackage() throws InvalidFormatException {
                Preconditions.checkState(testTemplate.canRead());

                return OPCPackage.open(testTemplate);
            }
        };
	}

	public DocxProcessor(final InputStream is) {
	    Preconditions.checkNotNull(is);

        opcPackageProvider = new OPCPackageProvider() {

            @Override
            public OPCPackage getPOCPackage() throws InvalidFormatException, IOException {
                return OPCPackage.open(is);
            }
        };
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
	public void process(Map<String, Substitution> substitutionMap, File resultDocx) throws IOException, InvalidFormatException {
		Preconditions.checkNotNull(substitutionMap);
		Preconditions.checkNotNull(resultDocx);
		Preconditions.checkArgument(resultDocx.canWrite());

        XWPFDocument doc = new XWPFDocument(opcPackageProvider.getPOCPackage());

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

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(resultDocx);
        	doc.write(out);
        } finally {
            if (out != null) {
                out.flush();
                out.close();
            }
        }

    }

	private static void processHeaderFooter(Map<String, Substitution> substitutionMap, XWPFHeaderFooter hf) {
		for (XWPFParagraph par : hf.getParagraphs()) {
			processParagraph(substitutionMap, par);
		}
	}

	private static void processFootnote(Map<String, Substitution> substitutionMap, XWPFFootnote fNote) {
		for (XWPFParagraph par : fNote.getParagraphs())
			processParagraph(substitutionMap, par);
	}

	private static void processTable(Map<String, Substitution> substitutionMap, XWPFTable tbl) {
		for (XWPFTableRow row : tbl.getRows())
			for (XWPFTableCell cell : row.getTableCells())
				for (XWPFParagraph p : cell.getParagraphs())
					processParagraph(substitutionMap, p);
	}

	private static void processParagraph(Map<String, Substitution> substitutionMap, XWPFParagraph par) {
		for(XWPFRun run : par.getRuns()) {
			for (int i=0; i< run.getCTR().sizeOfTArray(); ++i) {
				String text = run.getText(i);
				for (Entry<String, Substitution> subst : substitutionMap.entrySet()) { //TODO:remove println
					String stringToReplace = subst.getKey();
                    String stringToReplaceItWith = subst.getValue().getValue();
                    text = text.replaceAll(stringToReplace, stringToReplaceItWith);
				}
				run.setText(text, i);
			}
		}
	}
}
