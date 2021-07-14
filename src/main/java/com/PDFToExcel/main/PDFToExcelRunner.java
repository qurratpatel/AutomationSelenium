package com.PDFToExcel.main;


import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;  

public class PDFToExcelRunner {
	
	  private static final Logger logger = LogManager.getLogger(PDFToExcelRunner.class); 
//	  String log4jConfPath = System.getProperty("user.dir")+"/src/main/resources/log4j/log4j.properties";

	// static String pdfFileName = System.getProperty("user.dir") +
	// "/data/SampleData25Records.pdf";
	static String pdfFileName = System.getProperty("user.dir") + "/data/InputLatest_MultiPage_PDF.pdf";

	static String excelFilePath = System.getProperty("user.dir") + "/src/main/resources/documents/";
	static String templatesPath = System.getProperty("user.dir") + "/src/main/resources/templates/";
	static String templateMA = templatesPath + "MA_Template_2021.xlsx";
	static String templateMB = templatesPath + "MB_Template_2021.xlsx";

	public static void main(String[] args) {
		BasicConfigurator.configure();
		logger.info("Inside PDF to excel runner");
		PDFToExcelProcess pdfToExcelProcess = new PDFToExcelProcess();
		pdfToExcelProcess.pdfToExcel(excelFilePath, pdfFileName, templateMA, templateMB);
	}
}
