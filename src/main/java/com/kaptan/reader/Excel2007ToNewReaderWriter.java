package com.kaptan.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2007ToNewReaderWriter extends AbstractExcelFileHelper {

	public void initWorkBook() {
		setWorkbook(new XSSFWorkbook());
	}

	public void initWorkBook(FileInputStream fis) throws IOException {
		setWorkbook(new XSSFWorkbook(fis));
	}

	public List<ListOfStringData> readDataAsString(FileInputStream fis) throws IOException {
		setWorkbook(new XSSFWorkbook(fis));

		List<ListOfStringData> excelData = readDataAsStringItems();

		fis.close();

		return excelData;
	}
}
