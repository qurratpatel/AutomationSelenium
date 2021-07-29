package com.PDFToExcel.utility;

import java.io.File;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

public class ReadPDFData extends PDFTextStripper {
	public ReadPDFData() throws IOException {
		
	}
	static PDDocument document = null;
	static List<String> rows = new ArrayList<String>();
	public static List<String> readPDF(String folderName) throws IOException {
		File file =new File(folderName);
		File[] files = file.listFiles();
		for (File filenew:files) {
		try {
			document = PDDocument.load((filenew));
			System.out.println(filenew);
			PDFTextStripper stripper = new ReadPDFData();
			stripper.setSortByPosition(true);
			stripper.setStartPage(0);
			
			stripper.setEndPage(document.getNumberOfPages());
			stripper.setPageStart("NEW");
			Writer dummy = new OutputStreamWriter(new ByteArrayOutputStream());
			stripper.writeText(document, dummy);
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.getMessage());
		}
		finally {
			if (document != null) {
				document.close();
			}}}
			return rows;
	}
	@Override
	protected void writeString(String str, List<TextPosition> textPositions) throws IOException {
		rows.add(str);
	}
}
