package com.ooxmlcompare.api.service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.ooxmlcompare.api.compare.ExcelCompareContext;
import com.ooxmlcompare.api.compare.ExcelCompareResult;
import com.ooxmlcompare.api.compare.NaiveWorkbookCompare;
import com.ooxmlcompare.api.helper.Pair;

@Service
public class ExcelCompareService {

	@Autowired
	private NaiveWorkbookCompare naiveWorkbookCompare;
	
	public ExcelCompareResult compareFiles(Pair<XSSFWorkbook, XSSFWorkbook> excelFilePair) {
		ExcelCompareContext context = new ExcelCompareContext(naiveWorkbookCompare);
		return context.executeStrategy(excelFilePair.first, excelFilePair.second);
	}

}
