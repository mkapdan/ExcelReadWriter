package com.kaptan.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class TestMain {

	public static void main(String[] args) throws IOException {
		String filePath = "C:\\Users\\mustafa.kapdan\\testExcel.xls";
		String fileoutPath = "C:\\Users\\mustafa.kapdan\\outExcel.xls";

		FileInputStream fis = new FileInputStream(new File(filePath));

		ExcelFileReadWriter xlxsOperator = new Excel2003ReaderWriter();

		List<ListOfStringData> cellDatas = xlxsOperator.readDataAsString(fis);

		ExcelFileReadWriter writerOperator = new Excel2003ReaderWriter();
		Set<String> header = new HashSet<String>();
		header.add("Head");
		header.add("First");
		header.add("Design");

		writerOperator.createSheet("WriteTest1");
		writerOperator.prepareColumnHeaderName(header);

		for (ListOfStringData cellData : cellDatas) {

			Row row = writerOperator.createNextRowOnCurrentSheet();

			for (String cell : cellData.getStringCellDatas()) {

				Cell nextCell = writerOperator.createNextCell(row);

				nextCell.setCellValue(cell);

				System.out.print(cell + "\t");
			}
			System.out.println();
		}

		writerOperator.createSheet("WriteTest2");
		writerOperator.prepareColumnHeaderName(header);

		for (ListOfStringData cellData : cellDatas) {

			Row row = writerOperator.createNextRowOnCurrentSheet();

			for (String cell : cellData.getStringCellDatas()) {

				Cell nextCell = writerOperator.createNextCell(row);

				nextCell.setCellValue(cell);

				System.out.print(cell + "\t");
			}
			System.out.println();
		}

		writerOperator.createSheet("WriteTest3");
		writerOperator.prepareColumnHeaderName(header);

		for (ListOfStringData cellData : cellDatas) {

			Row row = writerOperator.createNextRowOnCurrentSheet();

			for (String cell : cellData.getStringCellDatas()) {

				Cell nextCell = writerOperator.createNextCell(row);

				nextCell.setCellValue(cell);

				System.out.print(cell + "\t");
			}
			System.out.println();
		}

		writerOperator.getSheetAt(0);

		for (ListOfStringData cellData : cellDatas) {

			Row row = writerOperator.createNextRowOnCurrentSheet();

			for (String cell : cellData.getStringCellDatas()) {

				Cell nextCell = writerOperator.createNextCell(row);

				nextCell.setCellValue(cell);

				System.out.print(cell + "\t");
			}
			System.out.println();
		}

		FileOutputStream fos = new FileOutputStream(fileoutPath);
		writerOperator.write(fos);
	}
}
