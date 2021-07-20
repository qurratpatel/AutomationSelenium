package com.PDFToExcel.main;

//import org.apache.log4j.BasicConfigurator;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;

public class PDFToExcelRunner {

	private static Logger log = LogManager.getLogger(PDFToExcelRunner.class);

	static String pdfFileName = System.getProperty("user.dir") + "/data/Input.pdf";

	static String excelFilePath = System.getProperty("user.dir") + "/src/main/resources/documents/";
	static String templatesPath = System.getProperty("user.dir") + "/src/main/resources/templates/";
	static String templateMA = templatesPath + "MA_Template_2021.xlsx";
	static String templateMB = templatesPath + "MB_Template_2021.xlsx";

	public static void main(String[] args) {
		log.info("Inside PDF to excel runner");
		PDFToExcelProcess pdfToExcelProcess = new PDFToExcelProcess();
		pdfToExcelProcess.pdfToExcel(excelFilePath, pdfFileName, templateMA, templateMB);
	}
}
