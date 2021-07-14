package com.PDFToExcel.main;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.PDFToExcel.utility.ReadPDFData;
import com.PDFToExcel.utility.WriteExcel;

public class PDFToExcelProcess {

	List<String> columnsFromRow = new ArrayList<>();
	Map<String, List<List<String>>> eventTypeMap = new LinkedHashMap<>();
	Map<String, List<List<List<String>>>> groupTypeMap = new LinkedHashMap<>();
	WriteExcel writeToExcel = new WriteExcel();
	String date;
	String excelName;
	String rlidType;
	String subType;
	String metaData;
	String actualDateFormat;
	String multipleLines = null; 

	public void pdfToExcel(String excelFilePath, String pdfFileName, String templateMA, String templateMB) {
		List<String> rows = new ArrayList<>();
		try {
			// Reads the data from PDF
			rows = ReadPDFData.readPDF(pdfFileName);

			// extract RLIDType, SubType and Date from row 1 and 2
			String firstRow = rows.get(0);
			String secondRow = rows.get(1);

			// Sub type
			// int indexOfSeparator = firstRow.lastIndexOf("–");
			int indexOfSeparator = firstRow.lastIndexOf("-");
			subType = firstRow.substring(indexOfSeparator - 3, indexOfSeparator);

			// RLID type
			String[] rlidTypeSubString = firstRow.substring(firstRow.indexOf("RLID:"), firstRow.indexOf(subType))
					.split(":");
			rlidType = rlidTypeSubString[1].trim();
			metaData = rows.get(2);

			// Get date from row 2
			Pattern pattern = Pattern.compile("\\d{2}/\\d{2}/\\d{2}");
			Matcher matcher = pattern.matcher(secondRow);
			if (matcher.find()) {
				actualDateFormat = matcher.group();
				date = matcher.group().replace("/", "-");
			}
			// Excel name using RLID type, sub type and date
			excelName = rlidType + "_" + subType + "_" + date;
			List<String> ls = new ArrayList<>();

			// concatenate multiple lines
			for (int i = 4; i < rows.size(); i++) {
				if (rows.get(i).contains("NEW") || rows.get(i).contains("COR") || rows.get(i).contains("DEL")) {
					multipleLines = null; 
					ls.add(rows.get(i));
				} else if (!(rows.get(i).contains("Report") || rows.get(i).contains(actualDateFormat)
						|| rows.get(i) == null || rows.get(i).toLowerCase().contains("end") || rows.get(i).equals(" ")
						|| rows.get(i).isEmpty())) {
					if (multipleLines != null) {
						multipleLines = multipleLines.concat(rows.get(i)); 
						ls.set(ls.size() - 1, multipleLines);
					} else {
						multipleLines = rows.get(i - 1).concat(rows.get(i)); // assign 2 concatenated lines to multipleLines 
						ls.set(ls.size() - 1, multipleLines);
					}
				}
			}
			rows = ls;
//			 for (String row: rows){
//			 System.out.println(row);
//			 }
//			
			for (String row : rows) {
				// splits the row into columns
				columnsFromRow = Arrays.asList(row.split("\\s*,"));
				int rowsize = columnsFromRow.size();

				if (rowsize > 1 && !row.contains("META") && !row.contains("Report")) {
					// Sorting data based on eventype
					String eventTypeKey = columnsFromRow.get(3);
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
				writeToExcel.writeExcel(excelFilePath, excelName + "_" + key, groupTypeMap.get(key), templateMA,
						templateMB, metaData);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
