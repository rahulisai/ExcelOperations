package com.atmecs.excelopearation;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Purpose: To provide the last row number of particular excel sheet.
 */
public class RowNumberOperation {
	/**
	 * Purpose: This method provide the functionality to print the number of last
	 * row.
	 * 
	 * @param filepath provide the excel sheet path
	 * @param sheetnumber provide the sheet number
	 * @throws IOException IOException any exception occurred it will throw to calling method
	 */
	public void lastRow(String filepath, String sheetname) {
		
	try {
			FileInputStream inpustream = new FileInputStream(filepath);
			XSSFWorkbook workbook = new XSSFWorkbook(inpustream);
			XSSFSheet sheet = workbook.getSheet(sheetname);
			int lastRowNum = 0;
			Iterator<Row> rowIterator = sheet.rowIterator(); // Iterating each row
		
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				lastRowNum = row.getRowNum();  //Getting row number
			}
		
			System.out.println("Last row number is: " + (lastRowNum + 1));
			workbook.close();
			inpustream.close();
		
			}catch(NullPointerException e) {
				System.out.println("Sheet is not available");
			}catch(IOException e) {
				System.out.println("The system cannot find the file specified");
			}catch(OLE2NotOfficeXmlFileException e) {
				System.out.println("Only .xlsx file is supported");
			}catch(Exception e) {
				System.out.println(e.getMessage());
			   }
		}
	}
