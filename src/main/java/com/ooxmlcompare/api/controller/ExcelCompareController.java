package com.ooxmlcompare.api.controller;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.ooxmlcompare.api.compare.ExcelCompareResult;
import com.ooxmlcompare.api.compare.ExcelPartCompareResult;
import com.ooxmlcompare.api.exception.InvalidEntityException;
import com.ooxmlcompare.api.helper.Pair;
import com.ooxmlcompare.api.service.ExcelCompareService;
import com.ooxmlcompare.api.service.ExcelFileService;

@RestController
public class ExcelCompareController {

	private static final Logger logger = LoggerFactory.getLogger(ExcelCompareController.class);
	
	@Autowired
	private ExcelCompareService excelCompareService;
	
	@Autowired
	private ExcelFileService excelFileService;
	
    @PostMapping("/compare/contents")
    public ExcelCompareResult compareExcelFilesContents(@RequestParam("files") MultipartFile[] files) {
        List<XSSFWorkbook> excelFiles = Arrays.asList(files)
                .stream()
                .map(file -> excelFileService.reifyExcelWorkbook(file))
                .collect(Collectors.toList());

        excelFiles.removeAll(Collections.singleton(null));

        if (excelFiles.size() != 2) {
        	throw new InvalidEntityException("There must be exactly two excel files to compare");
        }
        
        Pair<XSSFWorkbook, XSSFWorkbook> excelFilePair = new Pair<XSSFWorkbook, XSSFWorkbook>(excelFiles.get(0), excelFiles.get(1));

        return excelCompareService.compareViewableContents(excelFilePair);
    }
    
	@PostMapping("/compare/ooxml")
	public ExcelPartCompareResult compareExcelFilesXML(@RequestParam("files") MultipartFile[] files) throws IOException {
		if (files.length != 2) {
			throw new InvalidEntityException("There must be exactly two excel files to compare");
		}

		Pair<FileInputStream, FileInputStream> excelFilePair = new Pair<FileInputStream, FileInputStream>(
				(FileInputStream) files[0].getInputStream(), (FileInputStream) files[1].getInputStream());

		return excelCompareService.compareOOXMLContents(excelFilePair);
	}
}
