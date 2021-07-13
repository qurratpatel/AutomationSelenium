package com.PDFToExcel.main;

import com.aventstack.extentreports.ExtentTest;

public class PDFToExcelRunner {

//	static String pdfFileName = System.getProperty("user.dir") + "/data/SampleData25Records.pdf";
	static String pdfFileName = System.getProperty("user.dir") + "/data/InputLatest_MultiPage_PDF.pdf";
	
	static String excelFilePath = System.getProperty("user.dir") + "/src/main/resources/documents/";
	static String templatesPath = System.getProperty("user.dir")+ "/src/main/resources/templates/";
	static String templateMA =templatesPath+ "MA_Template_2021.xlsx";
	static String templateMB =templatesPath+ "MB_Template_2021.xlsx";
	
	public static void main(String[] args) {
		PDFToExcelProcess pdfToExcelProcess = new PDFToExcelProcess();
		pdfToExcelProcess.pdfToExcel(excelFilePath, pdfFileName, templateMA,templateMB);
	}
}
