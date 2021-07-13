package com.PDFToExcel.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public void writeExcel(String excelFilePath, String excelName, List<List<List<String>>> list, String templateMA,
			String templateMB, String metaData) throws FileNotFoundException, IOException, InvalidFormatException {
		// check with 3.10 final
		// Handle exception,
		if (excelName.contains("MA")) {
			cloneTemplate(templateMA, excelFilePath + excelName + ".xlsx");
		} else if (excelName.contains("MB")) {
			cloneTemplate(templateMB, excelFilePath + excelName + ".xlsx");
		}
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath + excelName + ".xlsx"));
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet mataDataSheet = workbook.getSheetAt(0);
		Cell cellMeta = mataDataSheet.createRow(1).createCell(2);
		cellMeta.setCellValue(metaData);

		for (int i = 0; i < list.size(); i++) {
			String eventType = list.get(i).get(0).get(3);
			Sheet sheet = workbook.getSheet(eventType);
			int rowCount = 1;
			if (sheet != null) {
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
		}
		try (FileOutputStream outputStream = new FileOutputStream(excelFilePath + excelName + ".xlsx")) {
			workbook.write(outputStream);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public void cloneTemplate(String templateName, String clonedFileName) throws InvalidFormatException, IOException {
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
