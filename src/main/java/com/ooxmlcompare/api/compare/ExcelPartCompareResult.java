package com.ooxmlcompare.api.compare;

import java.util.List;
import java.util.Map;

public class ExcelPartCompareResult {
	private List<String> partDifferences;

	public ExcelPartCompareResult() {}

	public ExcelPartCompareResult(List<String> partDifferences) {
		this.setPartDifferences(partDifferences);
	}

	public List<String> getPartDifferences() {
		return partDifferences;
	}

	public void setPartDifferences(List<String> partDifferences) {
		this.partDifferences = partDifferences;
	}
}
