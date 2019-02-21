package com.ooxmlcompare.api.service;

import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelWorkbookUtility {

	public Set<String> getSheetNames(XSSFWorkbook source) {
		Set<String> sourceSheets = new HashSet<>();
	    for (Sheet sheet : source ) {
	    	sourceSheets.add(sheet.getSheetName());
	    }
	    return sourceSheets;
	}

}
