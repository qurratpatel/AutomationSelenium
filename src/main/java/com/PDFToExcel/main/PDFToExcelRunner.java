package com.PDFToExcel.main;

import com.aventstack.extentreports.ExtentTest;

public class PDFToExcelRunner {

	static String pdfFileName = System.getProperty("user.dir") + "/data/SampleData-converted (2).pdf";
	static String excelName = "workbook.xlsx";
	static String excelFilePath = System.getProperty("user.dir") + "/src/main/resources/documents/";
	public ExtentTest logger;

	public static void main(String[] args) {
		PDFToExcelProcess pdfToExcelProcess = new PDFToExcelProcess();
		pdfToExcelProcess.pdfToExcel(excelFilePath, pdfFileName, excelName);
	}
}
