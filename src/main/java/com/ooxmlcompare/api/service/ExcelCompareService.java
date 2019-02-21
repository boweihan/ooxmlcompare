package com.ooxmlcompare.api.service;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.ooxmlcompare.api.compare.ExcelCompareContext;
import com.ooxmlcompare.api.compare.ExcelCompareResult;
import com.ooxmlcompare.api.compare.ExcelPartCompareResult;
import com.ooxmlcompare.api.compare.NaiveWorkbookCompare;
import com.ooxmlcompare.api.helper.OOXMLCompare;
import com.ooxmlcompare.api.helper.Pair;
import com.ooxmlcompare.api.helper.PartCompare;

@Service
public class ExcelCompareService {

	@Autowired
	private NaiveWorkbookCompare naiveWorkbookCompare;
	
	public ExcelCompareResult compareViewableContents(Pair<XSSFWorkbook, XSSFWorkbook> excelFilePair) {
		ExcelCompareContext context = new ExcelCompareContext(naiveWorkbookCompare);
		return context.executeStrategy(excelFilePair.first, excelFilePair.second);
	}

	public ExcelPartCompareResult compareParts(Pair<FileInputStream, FileInputStream> excelFilePair) throws IOException {
		PartCompare compare = PartCompare.compareContentExceptCoreOf(excelFilePair.first, excelFilePair.second);
		return new ExcelPartCompareResult(compare.getMessages());
	}
	
	public ExcelPartCompareResult compareOOXMLContents(Pair<FileInputStream, FileInputStream> excelFilePair) throws IOException {
		OOXMLCompare compare = OOXMLCompare.compareContentExceptCoreOf(excelFilePair.first, excelFilePair.second);
		return new ExcelPartCompareResult(compare.getMessages());
	}
}
