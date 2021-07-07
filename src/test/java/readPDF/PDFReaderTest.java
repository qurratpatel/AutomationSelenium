package readPDF;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;


import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.testng.Assert;

public class PDFReaderTest{

static List<String> lines = new ArrayList<>();
	URL pdfURL;
	BufferedInputStream fileParse;
	PDDocument document = null;
	String pdfPath="file:///"+System.getProperty("user.dir")+"/data/sampleData1.pdf";
	
	//@Test
	public void readPDFTest() throws IOException{
		System.out.println(pdfPath);
		pdfURL =new URL(pdfPath);// file path
		InputStream inputStream = pdfURL.openStream(); // opens the connection for the url
		fileParse = new BufferedInputStream(inputStream);
		document = PDDocument.load(fileParse);
		String pdfData = new PDFTextStripper().getText(document);
		System.out.println(pdfData);
		Assert.assertTrue(pdfData.contains("MENO"));
	}
	
}

