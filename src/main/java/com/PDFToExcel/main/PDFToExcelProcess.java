package com.PDFToExcel.main;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;

import com.PDFToExcel.utility.ReadPDFData;
import com.PDFToExcel.utility.WriteExcel;

public class PDFToExcelProcess {

	List<String> columnsFromRow = new ArrayList<>();
	HashMap<String, List<List<String>>> eventTypeMap = new HashMap<>();
	HashMap<String, List<List<List<String>>>> groupTypeMap = new HashMap();
	WriteExcel writeToExcel = new WriteExcel();
	String date;
	String excelName;

	public void pdfToExcel(String excelFilePath, String pdfFileName) {
		List<String> rows = new ArrayList<>();
		try {
			// Reads the data from PDF
			rows = ReadPDFData.readPDF(pdfFileName);
			//extract RLIDType, SubType and Date from row 1 and 2
			String firstRow = rows.get(0);
			String secondRow = rows.get(1);
			String rlidTypeSubString = firstRow.substring(firstRow.indexOf("RLID:"), firstRow.indexOf("Sub")).trim();
			//doubt
			String subTypeSubstring = firstRow.substring(firstRow.indexOf("Sub ( "), firstRow.indexOf("C")).trim();
			String[] rlidArray = rlidTypeSubString.split(":");
			String rlidType = rlidArray[1].trim();
		//	String[] subArray = subTypeSubstring.split // Doubt
		//	String subType = subArray[1].trim();

			//Get date from row 2
			Pattern pattern = Pattern.compile("\\d{2}/\\d{2}/\\d{2}");
			Matcher matcher = pattern.matcher(secondRow);
			if (matcher.find()) {
				date = matcher.group().replace("/" , "-");
			}
			excelName = rlidType + "_" + date;
			
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
						// added in key value pair- key: eventTypeKey, value:
						// columns,
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
//				if (key.equalsIgnoreCase("MA")) {
//					// writeToExcel.writeExcel(excelFilePath+key+".xlsx",key +
//					// excelName, groupTypeMap.get(key));
//				} else if (key.equalsIgnoreCase("MB")) {
//					// writeToExcel.writeExcel(excelFilePath+key+".xlsx",key +
//					// excelName, groupTypeMap.get(key));
//				} else {
//					System.out.println(key);
//					// If the workbook is not present should we create? Or
//					// simply update in the log file that workbook is not
//					// present
//				}
				writeToExcel.writeExcel(excelFilePath, excelName+"_"+key, groupTypeMap.get(key));
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
