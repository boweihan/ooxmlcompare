package com.ooxmlcompare.api.helper;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class OOXMLCompare {

	private List<String> messages = new ArrayList<String>();
	private boolean matches = true;

	public boolean isEqualExceptCore() {
		return matches;
	}

	public List<String> getMessages() {
		return messages;
	}

	public static OOXMLCompare compareContentExceptCoreOf(FileInputStream facitInputStream, FileInputStream underTestInputStream)
		throws IOException {
		OOXMLCompare compare = new OOXMLCompare();
		compare.compareContentExceptCoreOf_(facitInputStream, underTestInputStream);
		return compare;
	}

	private void compareContentExceptCoreOf_(FileInputStream facitInputStream, FileInputStream underTestInputStream)
		throws IOException {
		// TODO actually compare individual XML parts of an OOXML file
	}
}
