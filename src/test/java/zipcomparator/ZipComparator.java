package zipcomparator;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

import org.apache.commons.io.IOUtils;

import com.google.common.collect.Sets;

public final class ZipComparator {

	public static boolean equal(File zip1, File zip2) throws ZipException, IOException {
		try(ZipFile zf1 = new ZipFile(zip1); ZipFile zf2 = new ZipFile(zip2)){
			Set<String> set1 = getEntryNames(zf1);

			Set<String> set2 = getEntryNames(zf2);

			if (!Sets.difference(set1, set2).isEmpty())
				return false;

			for (String entryName : set1) if (!equalEntry(zf1, zf2, entryName))
				return false;

			return true;
		}
	}

	private static Set<String> getEntryNames(ZipFile zf) {
		Set<String> set1 = new LinkedHashSet<String>();
		for (Enumeration<? extends ZipEntry> e = zf.entries(); e.hasMoreElements();)
			set1.add(((ZipEntry) e.nextElement()).getName());
		return set1;
	}

	private static boolean equalEntry(ZipFile zf1, ZipFile zf2, String entryName) throws IOException {
		ZipEntry entry1 = zf1.getEntry(entryName);
		ZipEntry entry2 = zf2.getEntry(entryName);
		
		if (entry1.isDirectory() != entry2.isDirectory())
			return false;
		
		long crc2 = entry2.getCrc();
		long crc1 = entry1.getCrc();
		if (
				crc1 != -1 &&
				crc2 != -1 &&
				crc1 == crc2)
			return true;
		
		try (InputStream is1 = zf1.getInputStream(entry1); InputStream is2 = zf2.getInputStream(entry2)) {
			return IOUtils.contentEquals(is1, is2);
		}
	}

}
