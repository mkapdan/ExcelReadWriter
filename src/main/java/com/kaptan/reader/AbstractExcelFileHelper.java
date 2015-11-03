package com.kaptan.reader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public abstract class AbstractExcelFileHelper implements ExcelFileReadWriter {

	private String DEFAULT_DATE_FORMAT = "dd/mm/yyyy";

	protected String fileName;

	protected String cellDateFormat;

	protected String FILE_EXTENSION;

	protected Workbook workbook;

	protected Font defaultFont;

	protected CellStyle defaultCellStyle;

	protected CreationHelper currentCreationHelper;

	protected Sheet currentSheet;

	private Map<Integer, Integer> sheetInxRowCountMap = null;

	public AbstractExcelFileHelper() {
		super();
		cellDateFormat = DEFAULT_DATE_FORMAT;
		sheetInxRowCountMap = new HashMap<Integer, Integer>();
		initWorkBook();
	}

	public Workbook getWorkBook() {
		return workbook;
	}

	public List<Cell> prepareColumnHeaderName(Set<String> namesOfColumns) {

		List<Cell> headerCells = new ArrayList<Cell>();

		Sheet newSheet = getCurrentSheet();

		CellStyle cellStyle = getCellStyle();

		Font font = getFont();

		font.setBoldweight((short) 2);

		cellStyle.setFont(font);

		int headerRowIndex = 0;
		// Header Row should be in 0 index
		Row headerRow = newSheet.createRow(headerRowIndex);

		Iterator<String> columnNames = namesOfColumns.iterator();

		int cellCount = 0;
		while (columnNames.hasNext()) {
			String columnName = columnNames.next();
			String safeName = getSafeName(columnName);
			Cell headerCell = headerRow.createCell(cellCount++);
			headerCell.setCellValue(safeName);
			headerCell.setCellStyle(cellStyle);
			headerCells.add(headerCell);
		}

		return headerCells;
	}

	public Sheet getSheetAt(int sheetNo) {

		int numberOfSheets = getWorkBook().getNumberOfSheets();
		if (sheetNo > numberOfSheets) {
			throw new RuntimeException("There is no page with given number : " + sheetNo);
		}

		setCurrentSheet(getWorkBook().getSheetAt(sheetNo));
		return getCurrentSheet();
	}

	public CreationHelper getCreationHelper() {

		if (null == getCurrentCreationHelper()) {
			setCurrentCreationHelper(getWorkBook().getCreationHelper());
		}

		return getCurrentCreationHelper();
	}

	public void write(FileOutputStream fileOut) {
		try {
			getWorkBook().write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (IOException e) {
			throw new RuntimeException(e.toString());
		}
	}

	public CellStyle getCellStyle() {

		if (null == getDefaultCellStyle()) {
			setDefaultCellStyle(getWorkBook().createCellStyle());
		}

		return getDefaultCellStyle();
	}

	public Font getFont() {

		if (null == getDefaultFont()) {
			setDefaultFont(getWorkBook().createFont());
		}

		return getDefaultFont();
	}

	public FormulaEvaluator getFormulaEvaluator() {
		return getCreationHelper().createFormulaEvaluator();
	}

	public String getSafeName(String proposalName) {
		return WorkbookUtil.createSafeSheetName(proposalName);
	}

	public void setDEFAULT_DATE_FORMAT(String dEFAULT_DATE_FORMAT) {
		DEFAULT_DATE_FORMAT = dEFAULT_DATE_FORMAT;
	}

	public String getFileName() {
		return this.fileName;
	}

	public String getCellDateFormat() {

		return cellDateFormat;
	}

	public void setCellDateFormat(String cellDateFormat) {
		this.cellDateFormat = cellDateFormat;
	}

	public Sheet createSheet(String proposalSheetName) {

		Sheet createdSheet = null;
		if (null == proposalSheetName || proposalSheetName.isEmpty()) {

			createdSheet = getWorkBook().createSheet();
		} else {
			createdSheet = getWorkBook().createSheet(getSafeName(proposalSheetName));
		}

		//
		int sheetIndex = getWorkBook().getSheetIndex(createdSheet);

		sheetInxRowCountMap.put(sheetIndex, 0);
		// Set current sheet
		setCurrentSheet(createdSheet);

		return createdSheet;
	}

	public Row createNextRowOnSheet(Sheet givenSheet) {

		Row createdRow = null;
		if (null == givenSheet) {

			givenSheet = getCurrentSheet();
		}

		int currentSheetIndex = getWorkBook().getSheetIndex(givenSheet);

		int currentRowIndOfSheet = sheetInxRowCountMap.get(currentSheetIndex);

		createdRow = givenSheet.createRow(++currentRowIndOfSheet);

		// Update Index
		sheetInxRowCountMap.put(currentSheetIndex, currentRowIndOfSheet);

		return createdRow;
	}

	public Row createNextRowOnCurrentSheet() {

		int currentSheetIndex = getWorkBook().getSheetIndex(getCurrentSheet());

		int currentRowIndOfSheet = sheetInxRowCountMap.get(currentSheetIndex);

		Row createdCCRow = getCurrentSheet().createRow(++currentRowIndOfSheet);

		// Update Index
		sheetInxRowCountMap.put(currentSheetIndex, currentRowIndOfSheet);

		return createdCCRow;
	}

	public Row createRowOnSheetSpecificRow(int givenRowId, Sheet givenSheet) {
		Row createdRow = null;
		if (null == givenSheet) {

			givenSheet = getCurrentSheet();
		}
		createdRow = givenSheet.createRow(givenRowId);
		return createdRow;
	}

	public Row createRowOnCurrentSheetSpecificRow(int givenRowId) {
		Row createdCCRow = getCurrentSheet().createRow(givenRowId);
		return createdCCRow;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

	private Font getDefaultFont() {
		return defaultFont;
	}

	public void setDefaultFont(Font defaultFont) {
		this.defaultFont = defaultFont;
	}

	private CellStyle getDefaultCellStyle() {
		return defaultCellStyle;
	}

	public void setDefaultCellStyle(CellStyle defaultCellStyle) {
		this.defaultCellStyle = defaultCellStyle;
	}

	private Sheet getCurrentSheet() {
		if (null == currentSheet) {
			this.currentSheet = createSheet("TEMP");
		}
		return currentSheet;
	}

	public void setCurrentSheet(Sheet currentSheet) {
		this.currentSheet = currentSheet;
	}

	private CreationHelper getCurrentCreationHelper() {
		return currentCreationHelper;
	}

	public void setCurrentCreationHelper(CreationHelper currentCreationHelper) {
		this.currentCreationHelper = currentCreationHelper;
	}

	public Cell createCellOnSpecificIndex(int cellIndex, Row row) {
		Cell newCell = null;
		newCell = row.createCell(cellIndex);
		return newCell;
	}

	public Cell createNextCell(Row row) {

		Cell newCell = null;
		int lastIndex = row.getPhysicalNumberOfCells();
		newCell = row.createCell(lastIndex++);
		return newCell;
	}

	protected List<ListOfStringData> readDataAsStringItems() {

		List<ListOfStringData> excelData = new ArrayList<ListOfStringData>();

		int numberOfSheets = getWorkBook().getNumberOfSheets();

		// loop through each of the sheets
		for (int i = 0; i < numberOfSheets; i++) {

			// Get the nth sheet from the workbook
			Sheet sheet = getWorkBook().getSheetAt(i);

			// every sheet has rows, iterate over them
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {

				// Get the row object
				Row row = rowIterator.next();

				// Every row has columns, get the column iterator and iterate
				// over them
				Iterator<Cell> cellIterator = row.cellIterator();

				ListOfStringData cellDataString = new ListOfStringData();

				while (cellIterator.hasNext()) {
					// Get the Cell object
					Cell cell = cellIterator.next();

					String cellData = "";
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						cellData = cell.getStringCellValue();
						break;
					case Cell.CELL_TYPE_NUMERIC:
						cellData = String.valueOf(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						cellData = String.valueOf(cell.getBooleanCellValue());

					}

					cellDataString.getStringCellDatas().add(cellData);

				} // end of cell iterator

				excelData.add(cellDataString);
			} // end of rows iterator

		} // end of sheets for loop

		return excelData;
	}

	public <T> Set<String> getColumnHeaders(Class<T> claszz) {

		Set<String> headers = new HashSet<String>();
		Field[] fields = claszz.getDeclaredFields();
		System.out.printf("%d fields:%n", fields.length);
		for (Field field : fields) {

			headers.add(field.getName());

			System.out.printf("%s %s %s%n", Modifier.toString(field.getModifiers()), field.getType().getSimpleName(),
					field.getName());
		}
		return headers;
	}
}
