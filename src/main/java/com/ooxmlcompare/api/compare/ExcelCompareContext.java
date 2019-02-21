package com.ooxmlcompare.api.compare;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCompareContext {
	private ExcelCompareStrategy strategy;
	
	public ExcelCompareContext(ExcelCompareStrategy strategy) {
		this.strategy = strategy;
	}
	
	public ExcelCompareResult executeStrategy(XSSFWorkbook current, XSSFWorkbook backup) {
		return strategy.compare(current, backup);
	}
}
