package com.PDFToExcel.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public void writeExcel(String excelFile, String excelName, List<List<List<String>>> list) throws Exception {
		// check with 3.10 final
		//Need to Handle exception,
		FileInputStream inputStream = new FileInputStream(new File(excelFile));
		Workbook workbook = WorkbookFactory.create(inputStream);

		for (int i = 0; i < list.size(); i++) {
			String eventType = list.get(i).get(0).get(2);
			Sheet sheet = workbook.getSheet(eventType);
			// if(sheet==null){ // Have some doubts, if sheet is not present in
			// excel, should we create one? If yes we will not be having headers
			// data
			// sheet=workbook.createSheet(eventType);
			// }

			if (sheet != null) {
				int rowCount = 0;
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
					}
				}
			}
			try (FileOutputStream outputStream = new FileOutputStream(excelFile)) {
				workbook.write(outputStream);
			} catch (Exception e) {
				System.out.println(e.getMessage());
			}
		}
	}
}
