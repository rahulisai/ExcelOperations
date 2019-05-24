package com.atmecs.excelopearation;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;


/**
 * Purpose: ExcelFileOperation program implements an application to provide the
 * last row number and last column of given row number of particular excel
 * sheet.
 * 
 * @author rahul.isai
 *
 */
public class ExcelFileOperation {

	public static void main(String[] args){
		BufferedReader reader = null;

	try {
			System.out.println("Enter path of excel sheet: ");
			reader = new BufferedReader(new InputStreamReader(System.in));
			String filepath = reader.readLine();

			System.out.println("Enter excelsheet name: ");
			reader = new BufferedReader(new InputStreamReader(System.in));
			String sheetname =reader.readLine();

			RowNumberOperation rownumber = new RowNumberOperation();
			rownumber.lastRow(filepath, sheetname); // this method will display the last row of
														// the sheet having data

			System.out.println("Enter row number to get the last column: ");
			reader = new BufferedReader(new InputStreamReader(System.in));

			ColumnNumberOperation columnumber = new ColumnNumberOperation();
			int rownum = Integer.parseInt(reader.readLine());
			columnumber.lastColumnOfRow(filepath, rownum, sheetname); // this method will
																// display the last
																// column of given
																// row number
		} catch (NumberFormatException e) {
			System.out.println("Invalid data");
		} catch(IOException e) {
			System.out.println(e);
		} 
	}
}
