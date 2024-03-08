package com.utility;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

public class DocumentReader {

    public static void main(String[] args) {
		
  
          
          Properties config = loadConfig();
  	
  	    String folderPath = config.getProperty("folderPath");
  	    String excelPath = config.getProperty("excelFilePath");
          try {
              List<SummaryDetails> allNatureOfFacility = new ArrayList<>();

              try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(folderPath), "*.docx")) {
                  for (Path entry : stream) {
                      Map<String, String> borrowersMap = readBorrowers(entry.toString());
                      List<SummaryDetails> natureOfFacility = readNatureOfFacility(entry.toString(), borrowersMap);
                      allNatureOfFacility.addAll(natureOfFacility);
                  }
              }

              createExcelFile(excelPath, allNatureOfFacility);
              
 
			


        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Properties loadConfig() {
		// TODO Auto-generated method stub
    	// TODO Auto-generated method stub
 	   Properties config = new Properties();
        String configFilePath = "./config.properties"; // Replace with your actual path
        try (InputStream input = new FileInputStream(configFilePath)) {
            config.load(input);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return config;
	}

	private static void createExcelFile(String filePath, List<SummaryDetails> natureOfFacility) throws IOException {
		// TODO Auto-generated method stub
    	  try (Workbook workbook = new XSSFWorkbook()) {
              Sheet sheet = workbook.createSheet("NatureOfFacility");

              // Create header row
              Row headerRow = sheet.createRow(0);
              String[] headers = {"Borrower Name", "CIF", "Nature Of Facility", "Type", "Facility Limit", "Tenor", "Availability Period"};
              for (int i = 0; i < headers.length; i++) {
                  Cell cell = headerRow.createCell(i);
                  cell.setCellValue(headers[i]);
              }

              // Populate data rows
              int rowNum = 1;
              for (SummaryDetails details : natureOfFacility) {
                  Row row = sheet.createRow(rowNum++);
                  row.createCell(0).setCellValue(details.getBorrowerName());
                  row.createCell(1).setCellValue(details.getCif());
                  row.createCell(2).setCellValue(details.getNatureOfFaciltiy());
                  row.createCell(3).setCellValue(details.getType());
                  row.createCell(4).setCellValue(details.getFacilityLimit());
                  row.createCell(5).setCellValue(details.getTenor());
                  row.createCell(6).setCellValue(details.getPeriod());
              }

              // Write the workbook content to the file
              try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                  workbook.write(fileOut);
              }
          }
      
		
	}

	private static List<SummaryDetails> readNatureOfFacility(String filePath,Map borrowersMap) throws FileNotFoundException, IOException {
    	 try (FileInputStream fis = new FileInputStream(filePath);
                 XWPFDocument document = new XWPFDocument(fis)) {
    		 List<SummaryDetails> summary = new ArrayList<SummaryDetails>();

                for (XWPFTable table : document.getTables()) {
                	SummaryDetails details = new SummaryDetails();
                	String boName = (String) borrowersMap.get("Borrowers");
                	details.setBorrowerName(boName);
                	details.setCif((String)borrowersMap.get("CIF"));
                    for (XWPFTableRow row : table.getRows()) {
                    
                    	
                    	
                        String firstCellText = row.getCell(0).getText().trim();
                        String secondCellText ="";
                        if(row.getCell(1)!=null)
                        {
                        	 secondCellText = row.getCell(1).getText().trim();

                             //Check if the first column contains 'Nature Of Facility'
                             if (firstCellText.contains("Nature Of Facility")) {
                             	details.setNatureOfFaciltiy(secondCellText);
                             }
                             if (firstCellText.contains("Type")) {
                             	details.setType(secondCellText);
                             }
                             if (firstCellText.contains("Facility Limit")) {
                             	details.setFacilityLimit(secondCellText);
                             }
                             if (firstCellText.contains("Tenor")) {
                             	details.setTenor(secondCellText);
                             }
                             if (firstCellText.contains("Availability Period")) {
                             	details.setPeriod(secondCellText);
                             }
                        	 
                        }
                        
                  
                    }
                    summary.add(details);
                }
                return summary; // Return null if the value is not found
            }
	}

	public static Map<String, String> readBorrowers(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            Map<String, String> borrowersMap = new HashMap<>();
            StringBuilder borrowersValue = new StringBuilder();
            String cifvalue ="";
            boolean borrowerFound =false;
            boolean cifFound =false ;
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText().trim();
                
                if (text.startsWith("Borrower(s):") && borrowerFound !=true) {
                    // Extract the value after "Borrower(s):"
                    String value = text.substring("Borrower(s):".length()).trim();
                    borrowersValue.append(value);
                    borrowerFound=true;
                 
                }
                if (text.startsWith("CIF Number") && cifFound !=true) {
                    // Extract the value after "Borrower(s):"
                	cifvalue = text.substring("CIF Number".length()).trim();
                    
                    cifFound=true;
                 
                }
                
            }

            // Put the final value in the map
            borrowersMap.put("Borrowers", borrowersValue.toString().trim());
            borrowersMap.put("CIF", cifvalue.trim());
            return borrowersMap;
        }
    }
}
