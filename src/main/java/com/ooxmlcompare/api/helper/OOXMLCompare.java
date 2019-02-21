package com.ooxmlcompare.api.helper;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class OOXMLCompare {

	private List<String> messages = new ArrayList<String>();
	private boolean matches = true;

	public boolean isEqualExceptCore() {
		return matches;
	}

	public List<String> getMessages() {
		return messages;
	}

	public static OOXMLCompare compareFileContentExceptCoreOf(String facitFilename, String underTestFilename)
		throws FileNotFoundException, IOException {
		FileInputStream facitInputStream = new FileInputStream(facitFilename);
		FileInputStream underTestInputStream = new FileInputStream(underTestFilename);
		return OOXMLCompare.compareContentExceptCoreOf(facitInputStream, underTestInputStream);
	}

	public static OOXMLCompare compareContentExceptCoreOf(FileInputStream facitInputStream, FileInputStream underTestInputStream)
		throws IOException {
		OOXMLCompare compare = new OOXMLCompare();
		compare.compareContentExceptCoreOf_(facitInputStream, underTestInputStream);
		return compare;
	}

	private void compareContentExceptCoreOf_(FileInputStream facitInputStream, FileInputStream underTestInputStream)
		throws IOException {
		ZipInputStream facitStream = new ZipInputStream(facitInputStream);
		ZipInputStream underTestStream = new ZipInputStream(underTestInputStream);
		ZipEntry facitEntry = facitStream.getNextEntry();
		ZipEntry underTestEntry = underTestStream.getNextEntry();

		while (facitEntry != null) {
			matches = checkEntryNotNull(facitEntry, underTestEntry);
			if (!facitEntry.getName().equals("docProps/core.xml")) {
				matches = checkEntryName(facitEntry, underTestEntry);
				matches = checkContents(facitStream, underTestStream, facitEntry);
			}
			facitEntry = facitStream.getNextEntry();
			underTestEntry = underTestStream.getNextEntry();
		}
		facitStream.closeEntry();
		facitStream.close();
		underTestStream.close();
	}

	private boolean checkContents(ZipInputStream facitStream, ZipInputStream underTestStream, ZipEntry facitEntry)
		throws IOException {
		int len, underTestLen;
		byte[] facitBuffer = new byte[1024];
		byte[] actualBuffer = new byte[1024];
		while ((len = facitStream.read(facitBuffer)) > 0) {
			underTestLen = underTestStream.read(actualBuffer);
			if (underTestLen != len || !Arrays.equals(facitBuffer, actualBuffer)) {
				this.messages.add(facitEntry.getName() + " differs from facit");
				return false;
			}
		}
		return true;
	}

	private boolean checkEntryNotNull(ZipEntry facitEntry, ZipEntry underTestEntry) {
		if (underTestEntry == null) {
			this.messages.add(facitEntry.getName() + " not present in stream under test");
			return false;
		}
		return true;
	}

	private boolean checkEntryName(ZipEntry facitEntry, ZipEntry underTestEntry) {
		if (!facitEntry.getName().equals(underTestEntry.getName())) {
			this.messages.add("Where facit had " + facitEntry.getName() + " stream under test had " + underTestEntry.getName());
			return false;
		}
		return true;
	}
}
