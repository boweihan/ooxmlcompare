package com.ooxmlcompare.api.compare;

import java.util.List;
import java.util.Map;

public class ExcelCompareResult {
	private List<ComparisonDifference> sheetDifferences;
	private Map<String, List<ComparisonDifference>> sheetToCellDifferenceMap;

	public ExcelCompareResult() {}

	public ExcelCompareResult(List<ComparisonDifference> sheetDifferences,
			Map<String, List<ComparisonDifference>> sheetToCellDifferenceMap) {
		this.setSheetDifferences(sheetDifferences);
		this.setSheetToCellDifferenceMap(sheetToCellDifferenceMap);
	}

	public List<ComparisonDifference> getSheetDifferences() {
		return sheetDifferences;
	}

	public void setSheetDifferences(List<ComparisonDifference> sheetDifferences) {
		this.sheetDifferences = sheetDifferences;
	}

	public Map<String, List<ComparisonDifference>> getSheetToCellDifferenceMap() {
		return sheetToCellDifferenceMap;
	}

	public void setSheetToCellDifferenceMap(Map<String, List<ComparisonDifference>> sheetToCellDifferenceMap) {
		this.sheetToCellDifferenceMap = sheetToCellDifferenceMap;
	}
}
