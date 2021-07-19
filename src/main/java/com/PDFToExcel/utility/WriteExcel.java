package com.PDFToExcel.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public int writeExcel(String excelFilePath, String excelName, List<List<List<String>>> list, String templateMA,
			String templateMB, List<String> metaDataList)
			throws FileNotFoundException, IOException, InvalidFormatException {
		int totalRecords = 0;
		// check with 3.10 final
		// Handle exception,
		if (excelName.contains("MA")) {
			CloneTemplates.createTemplate(templateMA, excelFilePath + excelName + ".xlsx");
		} else if (excelName.contains("MB")) {
			CloneTemplates.createTemplate(templateMB, excelFilePath + excelName + ".xlsx");
		}
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath + excelName + ".xlsx"));
		Workbook workbook = WorkbookFactory.create(inputStream);
		CellStyle cellStyle = workbook.createCellStyle();

		// Update Meta data values in MetaData sheet
		Sheet mataDataSheet = workbook.getSheetAt(0);
		Row metaRow = mataDataSheet.createRow(2);
		int columnCountMeta = 0;
		for (String metaData : metaDataList) {
			Cell cell = metaRow.createCell(columnCountMeta++);
			if (metaData instanceof String) {
				cell.setCellValue((String) metaData);
			}
		}
		for (int i = 0; i < list.size(); i++) {
			String eventType = list.get(i).get(0).get(3);
			Sheet sheet = workbook.getSheet(eventType);
			int rowCount = 1;

			// Creates single sheet
			if (sheet == null) {
				Sheet newSheet = workbook.getSheet("EXTRASHEET");
				if (newSheet == null) {
					newSheet = workbook.createSheet("EXTRASHEET");
				}
				for (List<String> rowdata : list.get(i)) {
					rowCount = newSheet.getLastRowNum();
					Row row = newSheet.createRow(++rowCount);
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
			} else {
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
						if (sheet.getRow(0).getCell(columnCount) != null) {
							if (sheet.getRow(0).getCell(columnCount).toString().equals("R")
									&& (columnData.toString().isEmpty() || columnData.toString().equalsIgnoreCase(" ")
											|| columnData.toString() == null)) {
								cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
								cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

								cellStyle.setBorderRight(BorderStyle.THIN);
								cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

								cellStyle.setBorderBottom(BorderStyle.THIN);
								cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

								cellStyle.setBorderLeft(BorderStyle.THIN);
								cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());

								cellStyle.setBorderTop(BorderStyle.THIN);
								cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

								cell.setCellStyle(cellStyle);
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
		for (int j = 1; j < workbook.getNumberOfSheets(); j++) {
			Sheet sheet = workbook.getSheetAt(j);
			{
				if (sheet.getLastRowNum() > 1) {
					int totalrows = 0;
					if (sheet.getSheetName().equalsIgnoreCase("EXTRASHEET")) {
						totalrows = sheet.getLastRowNum();
					} else {
						totalrows = sheet.getLastRowNum() - 1;
					}
					totalRecords = totalRecords + totalrows;
				}
			}
		}
		return totalRecords;
	}

	public void setCell() {

	}
}
