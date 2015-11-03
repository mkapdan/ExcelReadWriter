package com.kaptan.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * This one is created for XLXS files
 * @author mustafa.kapdan
 *
 */
public class Excel2003ReaderWriter extends AbstractExcelFileHelper {

	public Excel2003ReaderWriter() {
	}

	public void initWorkBook(FileInputStream fis) throws IOException {

		setWorkbook(new HSSFWorkbook(fis));

	}

	public void initWorkBook() {
		setWorkbook(new HSSFWorkbook());
	}

	public List<ListOfStringData> readDataAsString(FileInputStream fis) throws IOException {

		setWorkbook(new HSSFWorkbook(fis));

		List<ListOfStringData> excelData = readDataAsStringItems();

		fis.close();

		return excelData;
	}

}
