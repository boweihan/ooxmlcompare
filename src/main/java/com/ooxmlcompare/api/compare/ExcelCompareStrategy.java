package com.ooxmlcompare.api.compare;

import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface ExcelCompareStrategy {
	
	ExcelCompareResult compare(XSSFWorkbook current, XSSFWorkbook backup);
	
	Map<String, List<ComparisonDifference>> getCellDifferences(XSSFWorkbook current, XSSFWorkbook backup);

	List<ComparisonDifference> getSheetDifferences(XSSFWorkbook current, XSSFWorkbook backup);
}
