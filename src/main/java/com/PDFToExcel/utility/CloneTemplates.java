package com.PDFToExcel.utility;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CloneTemplates {

	public static void createTemplate(String templateName, String clonedFileName) throws InvalidFormatException, IOException {
		Workbook wb = new XSSFWorkbook(OPCPackage.open(templateName));
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			Sheet sheet = wb.getSheetAt(i);
			Row row = sheet.getRow(1);
			for (int j = 0; j < row.getLastCellNum() - 1; j++) {
				Cell cell = row.getCell(j);
				String data = cell.toString();
			}
			FileOutputStream fileout = new FileOutputStream(clonedFileName);
			wb.write(fileout);
			fileout.close();
		}
	}
}
