package com.PDFToExcel.utility;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public void writeExcel(String excelFilePath, String excelName, List<List<List<String>>> list)
			throws FileNotFoundException, IOException {
		// check with 3.10 final
		// Handle exception,
		
		Workbook workbook = new XSSFWorkbook();
		for (int i = 0; i < list.size(); i++) {
			String sheetName = list.get(i).get(0).get(2);
			Sheet sheet = workbook.createSheet(sheetName);
			int rowCount = 0;
			int headerSize = 0;
			createHeader(sheetName, sheet, headerSize);
			for (List<String> rowdata : list.get(i)) {
				Row row = sheet.createRow(++rowCount);
				int columnCount = 0;

				for (Object columnData : rowdata) {
					Cell cell = row.createCell(++columnCount);
					if (columnData instanceof String) {
						cell.setCellValue((String) columnData);
					} else if (columnData instanceof Integer) {
						cell.setCellValue((Integer) columnData);
					}
					headerSize++;
				}
			}
			try (FileOutputStream outputStream = new FileOutputStream(excelFilePath+excelName+".xlsx")) {
				workbook.write(outputStream);
			} catch (Exception e) {
				System.out.println(e.getMessage());
			}
		}
	}

	private static void createHeader(String sheetName, Sheet sheet, int headerSize) {
		List<String> headerList = new ArrayList<>();
		Row rowHeader = sheet.createRow(0);

		switch (sheetName) {
		case "MABLN":
			headerList.add("A");
			headerList.add("B");
			headerList.add("Type");
			headerList.add("Something");
			headerList.add("D");
			headerList.add("D");
			break;
		case "MACND":
			headerList.add("B");
			headerList.add("B");
			headerList.add("B");
			headerList.add("B");
			headerList.add("B");
			break;
		case "MBAL":
			headerList.add("C");
			headerList.add("C");
			headerList.add("C");
			headerList.add("C");
			headerList.add("C");
			break;
		}

		for (int k = 0; k < headerList.size(); k++) {
			Cell cell = rowHeader.createCell(k + 1);
			cell.setCellValue(headerList.get(k));
		}

	}

}
