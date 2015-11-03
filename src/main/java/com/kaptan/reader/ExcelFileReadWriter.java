package com.kaptan.reader;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelFileReadWriter {

	void initWorkBook();

	void initWorkBook(FileInputStream fis) throws IOException;

	Workbook getWorkBook();

	String getFileName();

	List<Cell> prepareColumnHeaderName(Set<String> namesOfColumns);

	Sheet getSheetAt(int sheetNo);

	CreationHelper getCreationHelper();

	Row createNextRowOnSheet(Sheet givenSheet);

	Row createNextRowOnCurrentSheet();

	Row createRowOnSheetSpecificRow(int givenRowId, Sheet givenSheet);

	Row createRowOnCurrentSheetSpecificRow(int givenRowId);

	Cell createCellOnSpecificIndex(int cellIndex, Row row);

	Cell createNextCell(Row row);

	void write(FileOutputStream fileOut);

	List<ListOfStringData> readDataAsString(FileInputStream fis) throws IOException;

	CellStyle getCellStyle();

	Font getFont();

	FormulaEvaluator getFormulaEvaluator();

	String getSafeName(String proposalName);

	Sheet createSheet(String proposalSheetName);

	<T> Set<String> getColumnHeaders(Class<T> claszz);
}
