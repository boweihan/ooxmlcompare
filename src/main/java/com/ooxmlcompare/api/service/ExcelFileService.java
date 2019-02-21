package com.ooxmlcompare.api.service;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelFileService {

	private static final Logger logger = LoggerFactory.getLogger(ExcelFileService.class);
	
	public XSSFWorkbook reifyExcelWorkbook(MultipartFile file) {
		XSSFWorkbook wb = null;
		try (OPCPackage opcPackage = OPCPackage.open(file.getInputStream())) {
			wb = new XSSFWorkbook(opcPackage);
			System.out.println(wb);
		} catch (IOException | InvalidFormatException e) {
			logger.info("Unable to create XSSFWorkbook from multipart file", e);
		}
		return wb;
	}

}
