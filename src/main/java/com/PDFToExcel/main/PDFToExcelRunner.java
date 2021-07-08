package com.PDFToExcel.main;

import com.aventstack.extentreports.ExtentTest;

public class PDFToExcelRunner {

	static String pdfFileName = System.getProperty("user.dir") + "/data/SampleDataNew.pdf";
	static String excelFilePath = System.getProperty("user.dir") + "/src/main/resources/documents/";
	public ExtentTest logger;

	public static void main(String[] args) {
		PDFToExcelProcess pdfToExcelProcess = new PDFToExcelProcess();
		pdfToExcelProcess.pdfToExcel(excelFilePath, pdfFileName);
	}
}
