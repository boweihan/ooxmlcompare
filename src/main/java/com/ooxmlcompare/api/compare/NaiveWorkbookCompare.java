package com.ooxmlcompare.api.compare;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.google.common.collect.Sets;
import com.ooxmlcompare.api.helper.Difference;
import com.ooxmlcompare.api.helper.ExcelCellComparator;
import com.ooxmlcompare.api.service.ExcelWorkbookUtility;

@Component
public class NaiveWorkbookCompare implements ExcelCompareStrategy {

	@Autowired
	private ExcelWorkbookUtility excelWorkbookUtility;
	
	@Override
	public ExcelCompareResult compare(XSSFWorkbook current, XSSFWorkbook backup) {
		
		ExcelCompareResult result = new ExcelCompareResult();
		result.setSheetDifferences(getSheetDifferences(current, backup));
		result.setSheetToCellDifferenceMap(getCellDifferences(current, backup));
		return result;
	}
	
	@Override
	public Map<String, List<ComparisonDifference>> getCellDifferences(XSSFWorkbook current, XSSFWorkbook backup) {
		Map<String, List<ComparisonDifference>> cellDifferences = new HashMap<>();

		Set<String> currentSheets = excelWorkbookUtility.getSheetNames(current);
		Set<String> backupSheets = excelWorkbookUtility.getSheetNames(backup);

		Set<String> commonSheets = Sets.intersection(currentSheets, backupSheets);

		for (String sheetName : commonSheets) {
			List<ComparisonDifference> sheetCellDifferences = new ArrayList<>();
			ExcelCellComparator comparator = new ExcelCellComparator(sheetCellDifferences);
			
			Sheet currentSheet = current.getSheet(sheetName);
			Sheet backupSheet = backup.getSheet(sheetName);
			
			int startRow = Math.min(currentSheet.getFirstRowNum(), backupSheet.getFirstRowNum());
			int endRow = Math.max(currentSheet.getLastRowNum(), backupSheet.getLastRowNum());

			if (endRow - startRow > 0) {
				// endRow is zero indexed but we can't access the first row if data doesn't exist
				for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
					Row currentRow = currentSheet.getRow(rowNum);
					Row backupRow = backupSheet.getRow(rowNum);
					int rowLength = Math.max(currentRow.getLastCellNum(), backupRow.getLastCellNum());
					for (int cellNum = 0; cellNum < rowLength; cellNum++) {
						Cell sourceCell = currentRow.getCell(cellNum);
						Cell backupCell = backupRow.getCell(cellNum);
						comparator.compareDataInCell(sourceCell, backupCell);
					}
				}
			}
			
			cellDifferences.put(sheetName, sheetCellDifferences);
		}

		return cellDifferences;
	}

	@Override
	public List<ComparisonDifference> getSheetDifferences(XSSFWorkbook current, XSSFWorkbook backup) {
		List<ComparisonDifference> sheetDifferences = new ArrayList<>();

		Set<String> currentSheets = excelWorkbookUtility.getSheetNames(current);
		Set<String> backupSheets = excelWorkbookUtility.getSheetNames(backup);
		
		Set<String> removedSheets = Sets.difference(currentSheets, backupSheets);
		Set<String> addedSheets = Sets.difference(backupSheets, currentSheets);
		
		removedSheets.forEach(sheetName -> {
			sheetDifferences.add(new ComparisonDifference(sheetName, Difference.DELETE_SHEET));
		});
		addedSheets.forEach(sheetName -> {
			sheetDifferences.add(new ComparisonDifference(sheetName, Difference.INSERT_SHEET));
		});

		return sheetDifferences;
	}
}
