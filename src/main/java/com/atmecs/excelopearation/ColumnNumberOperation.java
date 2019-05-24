package com.atmecs.excelopearation;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Purpose: ExcelFileOperation program implements an application to provide the
 * last row number and last column of given row number of particular excel
 * sheet(.xlsx).
 *
 */
public class ColumnNumberOperation {

	/**
	 * Purpose: To get the last column number of the row provided by user.
	 * 
	 * @param filepath provide the excel sheet path
	 * @param row provide the sheet number to get the last row
	 * @throws IOException any exception occurred it will throw to calling method
	 */
	public void lastColumnOfRow(String filepath, int rownum, String sheetnumber) {
	try {
		 FileInputStream inputstream = new FileInputStream(filepath);
		 XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		 XSSFSheet sheet = workbook.getSheet(sheetnumber);
		 int lastCellNum = 0;

		 Iterator<Row> rowIterator = sheet.rowIterator();
		 if (rownum > 0) {
		 rownum = rownum - 1;
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

			// Comparing the row number provided by user
			if (rownum == row.getRowNum()) {
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						lastCellNum = cell.getColumnIndex();
					}
					System.out.println("Last Column number of the "+(rownum+1)+" row is :" + (lastCellNum + 1));
				  }
			 }
			 if (lastCellNum == 0) {
				  System.out.println("Row is empty");
			     }
		    }else {
			System.out.println("Invalid row number");
		 }
		 workbook.close();
		 inputstream.close();
		 
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

