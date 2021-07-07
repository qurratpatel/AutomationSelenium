package com.PDFToExcel.main;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

import com.PDFToExcel.utility.ReadPDFData;
import com.PDFToExcel.utility.WriteExcel;

public class PDFToExcelProcess {

	List<String> columnsFromRow = new ArrayList<>();
	HashMap<String, List<List<String>>> eventTypeMap = new HashMap<>();
	HashMap<String, List<List<List<String>>>> groupTypeMap = new HashMap();
	WriteExcel writeToExcel = new WriteExcel();
	
	public void pdfToExcel(String excelFilePath, String pdfFileName, String excelName) {
		List<String> rows = new ArrayList<>();
		try {
			// Reads the data from PDF
			rows = ReadPDFData.readPDF(pdfFileName);

			for (String row : rows) {
				// splits the row into columns
				columnsFromRow = Arrays.asList(row.split("\\s*,"));
				int rowsize = columnsFromRow.size();

				if (rowsize > 1) {// need to check
					// Sorting data based on eventype
					String eventTypeKey = columnsFromRow.get(2);
					if (!eventTypeMap.containsKey(eventTypeKey)) {
						List<List<String>> list = new ArrayList<>();
						list.add(columnsFromRow);
						// added in key value pair- key: eventTypeKey, value: columns,
						// eventypeKey will be our sheet
						eventTypeMap.put(eventTypeKey, list);
					} else {
						eventTypeMap.get(eventTypeKey).add(columnsFromRow);
					}
				}
			}
			// Workbook
			for (String key : eventTypeMap.keySet()) {
				String groupTypekey = key.toString().substring(0, 2);
				if (!groupTypeMap.containsKey(groupTypekey)) {
					List<List<List<String>>> list = new ArrayList<>();
					list.add(eventTypeMap.get(key));
					groupTypeMap.put(groupTypekey, list);
				} else {
					groupTypeMap.get(groupTypekey).add(eventTypeMap.get(key));
				}
			}
			
			for (String key : groupTypeMap.keySet()) {
				if(key.equalsIgnoreCase("MA")){
					excelName="Equity.xlsx";
					writeToExcel.writeExcel(excelFilePath+excelName,key + excelName, groupTypeMap.get(key));
				}
				else if (key.equalsIgnoreCase("MB")){
					excelName="Options.xlsx";
					writeToExcel.writeExcel(excelFilePath+excelName,key + excelName, groupTypeMap.get(key));
				}
				else{
					System.out.println(key);
					//If the workbook is not present should we create? Or simply update in the log file that workbook is not present
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
